namespace ExcelProcessor.Abstractions.Pointers
{
    /// <summary>
    /// Abstraction of row cursor
    /// </summary>
    public interface IRowCursor
    {
        /// <summary>
        /// Reference to current row at origin position
        /// </summary>
        ICellReference RowRef { get; }

        /// <summary>
        /// Reference to current position
        /// </summary>
        ICellReference CellRef { get; }

        /// <summary>
        /// Move CellRef to next column
        /// </summary>
        void NextColumn();

        /// <summary>
        /// Move CellRef to next row from origin
        /// </summary>
        void NextRowFromOrigin();
    }
}
