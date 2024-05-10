using ExcelProcessor.Abstractions.Generator.ReaderResults;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheetReader<TEntityReaded> : IExcelSheet
        where TEntityReaded : class, new()
    {
        /// <summary>
        /// Reader results
        /// </summary>
        IExcelReaderResult<TEntityReaded> Results { get; }

        /// <summary>
        /// Reads value as string at cursor position
        /// </summary>
        /// <returns></returns>
        string ReadValue();

        /// <summary>
        /// Reads value as DateTime at cursor position
        /// </summary>
        /// <param name="customError">Custom error message on error. Optional</param>
        /// <returns></returns>
        DateTime ReadValueAsDateTime(string customError = null);

        /// <summary>
        /// Reads value as integer at cursor position
        /// </summary>
        /// <param name="customError">Custom error message on error. Optional</param>
        /// <returns></returns>
        int ReadValueAsInteger(string customError = null);

        /// <summary>
        /// Reads value as bool at cursor position. Expected true o false
        /// </summary>
        /// <param name="customError">Custom error message on error. Optional</param>
        /// <returns></returns>
        bool ReadValueAsYesNo(string customError = null);

        /// <summary>
        /// Parallel reading of a block of content. Improve performance
        /// </summary>
        /// <typeparam name="TParallelReaded">Type of internal entity readed</typeparam>
        /// <param name="startsAtRow"></param>
        /// <param name="columnCount"></param>
        /// <param name="processAction"></param>
        /// <returns></returns>
        IEnumerable<TParallelReaded> ReadInParallel<TParallelReaded>(int startsAtRow, int columnCount, Func<string[], uint, TParallelReaded> processAction);
    }
}
