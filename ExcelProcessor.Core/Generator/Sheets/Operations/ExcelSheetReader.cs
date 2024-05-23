using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Exceptions;
using ExcelProcessor.Core.Pointers;
using System.Collections.Concurrent;
using System.Globalization;

namespace ExcelProcessor.Core.Generator.Sheets.Operations
{
    public class ExcelSheetReader<TEntityReaded> : ExcelSheetBase, IExcelSheetReader<TEntityReaded>
        where TEntityReaded: class, new()
    {
        private readonly WorkbookPart workbookPart;

        public IExcelReaderResult<TEntityReaded> Results { get; private set; }

        public ExcelSheetReader(WorkbookPart workbookPart, WorksheetPart worksheetPart, IExcelReaderResult<TEntityReaded> results) 
            : base(workbookPart, worksheetPart)
        {
            Results = results;
            this.workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        }

        public string ReadValue()
        {
            string value = ReadValueInternal(cursor.CellRef);
            return !string.IsNullOrEmpty(value) ? value.Trim() : value;
        }

        public DateTime ReadValueAsDateTime(string customError = null)
        {
            string value = ReadValue();
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dateAsDouble))
                return DateTime.FromOADate(dateAsDouble);
            else
            {
                string valueAsRepresentableString = value == null ? "Null" : value;
                Results.AddCellError(string.IsNullOrEmpty(customError) ? $"Value ({valueAsRepresentableString}) is not a Date" : customError, cursor.CellRef);
                return default;
            }
        }

        public int ReadValueAsInteger(string customError = null)
        {
            string value = ReadValue();
            if (int.TryParse(value, out int intValue))
                return intValue;
            else
            {
                string valueAsRepresentableString = value == null ? "Null" : value;
                Results.AddCellError(string.IsNullOrEmpty(customError) ? $"Value ({valueAsRepresentableString}) is not an integer" : customError, cursor.CellRef);
                return default;
            }
        }

        public bool ReadValueAsYesNo(string customError = null)
        {
            string value = ReadValue();            
            if (string.IsNullOrEmpty(value))
            {
                Results.AddCellError(GetYesNoErrorString(value, customError), cursor.CellRef);
                return default;
            }
            else
            {
                value = value.ToUpper();
                if (value == "YES")
                    return true;
                else if (value == "NO")
                    return false;
                else
                {
                    Results.AddCellError(GetYesNoErrorString(value, customError), cursor.CellRef);
                    return default;
                }
            }
        }

        private string GetYesNoErrorString(string value, string customError)
        {
            string valueAsRepresentableString = value == null ? "Null" : value;
            return string.IsNullOrEmpty(customError) ? $"Value ({valueAsRepresentableString}) is not a boolean (Yes / No)" : customError;
        }               

        public IEnumerable<TParallelReaded> ReadInParallel<TParallelReaded>(int startsAtRow, int columnCount, Func<string[], uint, TParallelReaded> processAction)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

            var rows = sheetData.Elements<Row>().AsEnumerable().Where(r => r.RowIndex >= startsAtRow);
            ConcurrentDictionary<uint, TParallelReaded> results = new ConcurrentDictionary<uint, TParallelReaded>();
            Parallel.ForEach(rows, row =>
            {
                if (row.RowIndex != null)
                {
                    bool hasData = false;
                    string[] rowData = new string[columnCount];
                    var cells = row.Elements<Cell>();

                    ICellReference cellRef = new CellReference((int)row.RowIndex.Value, "A");
                    int columnIndexToRead = 0;
                    for (int i = 0; i < columnCount; i++)
                    {
                        Cell currentCell = columnIndexToRead < cells.Count() ? cells.ElementAt(columnIndexToRead) : null;
                        if (currentCell != null)
                        {
                            if (currentCell.CellReference.Value == cellRef.ToString())
                            {
                                string cellValue;
                                if (currentCell.DataType != null && currentCell.DataType == CellValues.SharedString)
                                {
                                    int sharedId = int.Parse(currentCell.InnerText);
                                    cellValue = sharedStringTable.ElementAt(sharedId).InnerText;
                                }
                                else
                                    cellValue = currentCell.InnerText;

                                rowData[i] = cellValue;
                                if (!hasData && !string.IsNullOrEmpty(cellValue))
                                    hasData = true;

                                columnIndexToRead++;
                            }
                        }
                        else
                            rowData[i] = null;

                        cellRef = cellRef.NextColumn();
                    }
                    if (hasData)
                        results.TryAdd(row.RowIndex, processAction(rowData, row.RowIndex));
                }
            });
            // Sorted as expected
            return results.OrderBy(r => r.Key).Select(r => r.Value);
        }

    }
}
