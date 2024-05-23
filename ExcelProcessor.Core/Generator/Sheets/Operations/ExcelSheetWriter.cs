using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Pointers;
using System.Drawing;
using System.Drawing.Imaging;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ExcelProcessor.Core.Generator.Sheets.Operations
{
    public class ExcelSheetWriter : ExcelSheetBase, IExcelSheetWriter
    {
        protected readonly IExcelStyles styles;
        public ExcelSheetWriter(WorkbookPart workbookPart, WorksheetPart worksheetPart, IExcelStyles styles = null)
            : base(workbookPart, worksheetPart)
        {            
            this.styles = styles;
        }

        #region Insert Value

        public void InsertValue(string value, ICellReference cellRef, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cellRef.Column, cellRef.Row, worksheetPart, CellValues.String);

            if (value != null)
                cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            ApplyFormat(cell, styleName);
        }

        public void InsertValue(string value, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cursor.CellRef.Column, cursor.CellRef.Row, worksheetPart, CellValues.String);

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            ApplyFormat(cell, styleName);
        }

        public void InsertValue(decimal value, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cursor.CellRef.Column, cursor.CellRef.Row, worksheetPart, CellValues.Number);
            var valueParse = cursor.CellRef.ToDoubleChange(value);
            cell.CellValue = new CellValue(valueParse);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            ApplyFormat(cell, styleName);
        }

        public void InsertValue(int value, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cursor.CellRef.Column, cursor.CellRef.Row, worksheetPart, CellValues.Number);
            cell.CellValue = new CellValue(value.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            ApplyFormat(cell, styleName);
        }

        public void InsertValue(DateTime value, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cursor.CellRef.Column, cursor.CellRef.Row, worksheetPart, CellValues.String);

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.Date);
            ApplyFormat(cell, styleName);
        }

        #endregion

        public void InsertFormula(IFormula formula, string styleName = null)
        {
            Cell cell = InsertCellInWorksheet(cursor.CellRef.Column, cursor.CellRef.Row, worksheetPart, CellValues.Number);
            cell.CellFormula = formula.Build();
            ApplyFormat(cell, styleName);
        }

        public void SetRowHeight(double rowHeight)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            int rowIndex = cursor.CellRef.Row;

            // If row does not exists: create
            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == rowIndex))
                row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            else
            {
                row = new Row() { RowIndex = (uint)rowIndex };
                sheetData.Append(row);
            }

            row.Height = new DoubleValue(rowHeight);
            row.CustomHeight = new BooleanValue(true);
        }

        private Cell InsertCellInWorksheet(string columnName, int rowIndex, WorksheetPart worksheetPart, CellValues dataType)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If row does not exists: insert
            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == rowIndex))
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            }
            else
            {
                row = new Row() { RowIndex = (uint)rowIndex };
                sheetData.Append(row);
            }

            // If cell not exists: insert
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
            {
                Cell cell = row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
                cell.DataType = new EnumValue<CellValues>(dataType);
                return cell;
            }
            else
            {
                // Find cell in sequential order
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length &&
                        string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference, DataType = new EnumValue<CellValues>(dataType) };
                row.InsertBefore(newCell, refCell);
                return newCell;
            }
        }

        public void CreateTableWithHeaders(ICellReference from, ICellReference to)
        {
            string refRange = $"{from.ToExcelString()}:{to.ToExcelString()}";
            int tableNumber = worksheetPart.TableDefinitionParts.Count() + 1;
            Table table = new Table()
            {
                Id = (uint)tableNumber,
                Name = $"Table{tableNumber}",
                Reference = refRange,
                AutoFilter = new AutoFilter()
                {
                    Reference = refRange,
                },
                TotalsRowShown = false,
                DisplayName = $"Table{tableNumber}"
            };
            // Columns
            uint columnIndex = 1;
            ICellReference cellIndex = new CellReference(from.Row, from.Column);
            bool createColumns = true;
            TableColumns columns = new TableColumns();
            while (createColumns)
            {
                columns.Append(new TableColumn()
                {
                    Id = columnIndex,
                    Name = ReadValueInternal(cellIndex),
                });

                if (cellIndex.Column == to.Column)
                    createColumns = false;
                else
                {
                    cellIndex = cellIndex.NextColumn();
                    columnIndex++;
                }
            }

            columns.Count = columnIndex;
            table.Append(columns);

            TableDefinitionPart tableDefParts = worksheetPart.AddNewPart<TableDefinitionPart>($"rId{tableNumber}");
            tableDefParts.Table = table;

            TableParts tableParts = (TableParts)worksheetPart.Worksheet.ChildElements.Where(ce => ce is TableParts).FirstOrDefault(); // Add table parts only once
            if (tableParts is null)
            {
                tableParts = new TableParts();
                worksheetPart.Worksheet.Append(tableParts);
            }
            TablePart tablePart = new TablePart() { Id = $"rId{tableNumber}" };
            tableParts.Append(tablePart);

            tableParts.Count = 1;
        }

        #region Insert Image

        public void InsertImage(byte[] imgData, string imgDescription, long? customWidth = null, long? customHeight = null)
        {
            if (imgData != null)
            {
                int colNumber = cursor.CellRef.GetColumnIndex();
                int rowNumber = cursor.CellRef.Row;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                    drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
                    worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

                if (drawingsPart.WorksheetDrawing == null)
                    drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

                var worksheetDrawing = drawingsPart.WorksheetDrawing;

                ImagePart imagePart = null;
                long extentsCx = 0;
                long extentsCy = 0;
                if (imgData != null)
                {
                    using (MemoryStream imageStream = new MemoryStream(imgData))
                    {
                        using (MemoryStream imageMemStream = new MemoryStream())
                        {
                            imageStream.Position = 0;
                            imageStream.CopyTo(imageMemStream);
                            imageStream.Position = 0;

                            using (Bitmap bitMap = new Bitmap(imageMemStream))
                            {
                                imagePart = drawingsPart.AddImagePart(GetImagePartTypeByBitmap(bitMap));
                                imagePart.FeedData(imageStream);

                                extentsCx = (customWidth.HasValue ?
                                                customWidth.Value :
                                                bitMap.Width) * (long)(914400 / bitMap.HorizontalResolution);

                                extentsCy = (customHeight.HasValue ?
                                                        customHeight.Value :
                                                        bitMap.Height) * (long)(914400 / bitMap.VerticalResolution);
                            }
                        }
                    }

                    int colOffset = 0;
                    int rowOffset = 0;

                    var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
                    var nvpId = nvps.Count() > 0
                        ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                        : 1U;

                    var oneCellAnchor = new Xdr.OneCellAnchor(
                        new Xdr.FromMarker
                        {
                            ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                            RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                            ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                            RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                        },
                        new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                        new Xdr.Picture(
                            new Xdr.NonVisualPictureProperties(
                                new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDescription },
                                new Xdr.NonVisualPictureDrawingProperties(new OpenXmlDrawing.PictureLocks { NoChangeAspect = true })
                            ),
                            new Xdr.BlipFill(
                                new OpenXmlDrawing.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = OpenXmlDrawing.BlipCompressionValues.Print },
                                new OpenXmlDrawing.Stretch(new OpenXmlDrawing.FillRectangle())
                            ),
                            new Xdr.ShapeProperties(
                                new OpenXmlDrawing.Transform2D(
                                    new OpenXmlDrawing.Offset { X = 0, Y = 0 },
                                    new OpenXmlDrawing.Extents { Cx = extentsCx, Cy = extentsCy }
                                ),
                                new OpenXmlDrawing.PresetGeometry { Preset = OpenXmlDrawing.ShapeTypeValues.Rectangle }
                            )
                        ),
                        new Xdr.ClientData()
                    );
                    worksheetDrawing.Append(oneCellAnchor);
                }
            }
        }

        public static PartTypeInfo GetImagePartTypeByBitmap(Bitmap image)
        {
            if (ImageFormat.Bmp.Equals(image.RawFormat))
                return ImagePartType.Bmp;
            else if (ImageFormat.Gif.Equals(image.RawFormat))
                return ImagePartType.Gif;
            else if (ImageFormat.Png.Equals(image.RawFormat))
                return ImagePartType.Png;
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
                return ImagePartType.Tiff;
            else if (ImageFormat.Icon.Equals(image.RawFormat))
                return ImagePartType.Icon;
            else if (ImageFormat.Jpeg.Equals(image.RawFormat))
                return ImagePartType.Jpeg;
            else if (ImageFormat.Emf.Equals(image.RawFormat))
                return ImagePartType.Emf;
            else if (ImageFormat.Wmf.Equals(image.RawFormat))
                return ImagePartType.Wmf;
            else
                throw new Exception("Image type could not be determined.");
        }        

        #endregion

        #region Merge

        public void MergeRows(int count, string styleName = null)
        {
            ICellReference to = cursor.CellRef;
            for (int i = 1; i < count; i++)
                to = to.NextRow();
            Merge(cursor.CellRef, to, styleName);
        }

        public void MergeColumns(int count, string styleName = null)
        {
            ICellReference to = cursor.CellRef;
            for (int i = 1; i < count; i++)
                to = to.NextColumn();
            Merge(cursor.CellRef, to, styleName);
        }

        public void Merge(int columns, int rows, string styleName = null)
        {
            ICellReference to = cursor.CellRef;
            for (int i = 1; i < rows; i++)
                to = to.NextRow();

            for (int i = 1; i < columns; i++)
                to = to.NextColumn();
            Merge(cursor.CellRef, to, styleName);
        }

        private void Merge(ICellReference from, ICellReference to, string styleName = null)
        {
            MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            }
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"{from.ToExcelString()}:{to.ToExcelString()}") });

            if (!string.IsNullOrEmpty(styleName))
                ApplyFormatRange(from, to, styleName);
        }

        #endregion        

        #region Apply Format

        private void ApplyFormat(Cell cell, string styleName)
        {
            if (!string.IsNullOrEmpty(styleName))
                cell.StyleIndex = styles.GetStyleId(styleName);
        }

        private void ApplyFormatRange(ICellReference from, ICellReference to, string styleName)
        {
            if (from.Row.Equals(to.Row))
            {
                int columnIndexFrom = from.GetColumnIndex();
                int columnIndexTo = to.GetColumnIndex();

                ICellReference cellRef = new CellReference(from.Row, from.Column);
                for (int i = columnIndexFrom; i <= columnIndexTo; i++)
                {
                    InsertValue(null, cellRef, styleName);
                    cellRef = cellRef.NextColumn();
                }
            }
            else
            {
                string charFrom = new string(from.Column.Where(char.IsLetter).ToArray());
                string charTo = new string(to.Column.Where(char.IsLetter).ToArray());

                if (charFrom.Equals(charTo))
                {
                    for (int i = from.Row; i <= to.Row; i++)
                        InsertValue(null, new CellReference(i, charFrom), styleName);
                }
                else
                {
                    int columnIndexFrom = from.GetColumnIndex();
                    int columnIndexTo = to.GetColumnIndex();

                    ICellReference cellRef = new CellReference(from.Row, from.Column);
                    for (int columnIndex = columnIndexFrom; columnIndex <= columnIndexTo; columnIndex++)
                    {
                        for (int rowIndex = from.Row; rowIndex <= to.Row; rowIndex++)
                        {
                            InsertValue(null, new CellReference(rowIndex, cellRef.Column), styleName);
                        }
                        cellRef = cellRef.NextColumn();
                    }


                }
            }
        }

        #endregion

        public void Save() => worksheetPart?.Worksheet?.Save();
        
    }
}
