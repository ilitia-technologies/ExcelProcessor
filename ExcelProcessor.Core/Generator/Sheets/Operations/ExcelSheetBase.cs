using DocumentFormat.OpenXml.Packaging;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations
{
    public abstract class ExcelSheetBase
    {
        protected IRowCursor cursor = new RowCursor(new CellReference(1, "A"));
        protected readonly WorksheetPart worksheetPart;        

        protected ExcelSheetBase(WorksheetPart worksheetPart)
        {
            this.worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
        }

        public IRowCursor InitializeCursor(ICellReference cellRef)
        {
            cursor = new RowCursor(cellRef);
            return cursor;
        }
    }
}
