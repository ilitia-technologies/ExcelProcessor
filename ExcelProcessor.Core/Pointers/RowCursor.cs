using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Pointers
{
    public class RowCursor : IRowCursor
    {
        public ICellReference RowRef
        {
            get; private set;
        }

        public ICellReference CellRef
        {
            get; private set;
        }

        public RowCursor(ICellReference cellRef)
        {
            CellRef = cellRef;
            RowRef = cellRef;
        }

        public void NextColumn()
        {
            CellRef = CellRef.NextColumn();
        }
        public void NextRowFromOrigin()
        {
            RowRef = RowRef.NextRow();
            CellRef = RowRef;
        }
    }
}
