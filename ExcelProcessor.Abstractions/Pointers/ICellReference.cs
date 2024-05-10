namespace ExcelProcessor.Abstractions.Pointers
{
    /// <summary>
    /// Abstraction of a reference to a Excel cell
    /// </summary>
    public interface ICellReference
    {
        /// <summary>
        /// Row number
        /// </summary>
        int Row { get; }

        /// <summary>
        /// Column letter
        /// </summary>
        string Column { get; }

        /// <summary>
        /// Get next column
        /// </summary>
        /// <returns></returns>
        ICellReference NextColumn();

        /// <summary>
        /// Get next row
        /// </summary>
        /// <returns></returns>
        ICellReference NextRow();

        /// <summary>
        /// Cell position in Excel format as string
        /// </summary>
        /// <returns></returns>
        string ToExcelString();

        /// <summary>
        /// String as NumberDecimalSeparator as '.'
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        string ToDoubleChange(decimal number);

        /// <summary>
        /// Returns column index
        /// </summary>
        /// <returns></returns>
        int GetColumnIndex();
    }
}
