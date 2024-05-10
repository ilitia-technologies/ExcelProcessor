using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public abstract class Formula : IFormula
    {
        public abstract CellFormula Build();
    }
}
