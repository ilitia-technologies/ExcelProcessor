using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    /// <summary>
    /// Excel engine to read data
    /// </summary>
    public interface IExcelReaderEngine : IDisposable
    {
        /// <summary>
        /// Read Excel file and parse data
        /// </summary>
        /// <typeparam name="TEntityReaded">Type of data entity</typeparam>
        /// <param name="sheetParsers">Parser operations</param>
        /// <returns></returns>
        IExcelReaderResult<TEntityReaded> ReadFile<TEntityReaded>(IExcelSheetParser<TEntityReaded>[] sheetParsers)
            where TEntityReaded : class, new();
    }
}
