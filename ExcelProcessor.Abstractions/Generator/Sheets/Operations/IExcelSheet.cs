using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    /// <summary>
    /// Abstraction to Excel sheet operations
    /// </summary>
    public interface IExcelSheet
    {
        /// <summary>
        /// Initialize cursor at a position
        /// </summary>
        /// <param name="cellRef">Position</param>
        /// <returns>Instance of <see cref="IRowCursor"/></returns>
        IRowCursor InitializeCursor(ICellReference cellRef);
    }
}
