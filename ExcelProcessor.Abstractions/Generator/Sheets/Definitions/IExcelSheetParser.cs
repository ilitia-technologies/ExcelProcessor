using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Definitions
{
    public interface IExcelSheetParser<TEntityReaded>
        where TEntityReaded : class, new()
    {
        /// <summary>
        /// Sheet name in Excel file
        /// </summary>
        string SheetName { get; }

        /// <summary>
        /// Parse Excel sheet
        /// </summary>
        /// <param name="sheet"><see cref="IExcelSheetReader{TEntityReaded}"/> instance</param>
        void Parse(IExcelSheetReader<TEntityReaded> sheet);
    }
}
