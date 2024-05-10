using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    /// <summary>
    /// Excel engine to write data
    /// </summary>
    public interface IExcelWriterEngine : IDisposable
    {
        /// <summary>
        /// Create Excel fild and populate
        /// </summary>
        /// <typeparam name="TDataContext">Data context type</typeparam>
        /// <param name="sheetBuilders">Instances of <see cref="IExcelSheetBuilder{TDataContext}"/></param>
        /// <param name="dataContext">Data context</param>
        /// <returns>Byte array of excel file</returns>
        byte[] Create<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext)
            where TDataContext : class;

        /// <summary>
        /// Create Excel fild, populate and copy to directory
        /// </summary>
        /// <typeparam name="TDataContext">Data context type</typeparam>
        /// <param name="sheetBuilders">Instances of <see cref="IExcelSheetBuilder{TDataContext}"/></param>
        /// <param name="dataContext">Data context</param>
        /// <param name="outputFile">Path to output file</param>
        void CreateAndCopy<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext, string outputFile)
           where TDataContext : class;
    }
}