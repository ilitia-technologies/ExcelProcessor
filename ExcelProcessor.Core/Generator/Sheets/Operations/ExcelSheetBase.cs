using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Exceptions;
using ExcelProcessor.Core.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations
{
    public abstract class ExcelSheetBase
    {
        protected IRowCursor cursor = new RowCursor(new CellReference(1, "A"));
        protected readonly WorksheetPart worksheetPart;
        protected readonly WorkbookPart workbookPart;

        protected ExcelSheetBase(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            this.workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
            this.worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
        }

        public IRowCursor InitializeCursor(ICellReference cellRef)
        {
            cursor = new RowCursor(cellRef);
            return cursor;
        }

        protected string ReadValueInternal(ICellReference cellRef)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = cellRef.ToString();

            // If the worksheet does not contain a row with the specified row index: exception
            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == cellRef.Row))
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == cellRef.Row);
            }
            else
                throw new RowNotExistsException(cellRef.Row);

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == cellReference))
            {

                Cell cell = row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    int id;
                    if (int.TryParse(cell.InnerText, out id))
                    {
                        var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (sharedStringTable != null)
                        {
                            return sharedStringTable.SharedStringTable.ElementAt(id).InnerText;
                        }
                    }
                }
                return cell.CellValue != null ?
                                cell.CellValue.InnerText :
                                null;
            }
            else
                return null;
        }
    }
}
