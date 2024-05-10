using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    /// <summary>
    /// Abstraction of a Formula
    /// </summary>
    public interface IFormula
    {
        /// <summary>
        /// Generate formula in Excel format
        /// </summary>
        /// <returns></returns>
        CellFormula Build();
    }
}
