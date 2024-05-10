using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Definitions
{
    /// <summary>
    /// Sheet Builder
    /// </summary>
    /// <typeparam name="TDataContext">Type of data context</typeparam>
    public interface IExcelSheetBuilder<TDataContext>
        where TDataContext : class
    {
        /// <summary>
        /// Sheet name in Excel file
        /// </summary>
        string SheetName { get; }

        /// <summary>
        /// Populate Excel sheet
        /// </summary>
        /// <param name="sheet">Instances of <see cref="IExcelSheetWriter"/> </param>
        /// <param name="dataContext">Instance of data context</param>
        void Build(IExcelSheetWriter sheet, TDataContext dataContext);
    }
}
