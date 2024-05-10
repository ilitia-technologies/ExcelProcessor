using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class Count : Formula
    {
        private readonly ICellReference from;
        private readonly ICellReference to;
        public Count(ICellReference from, ICellReference to)
        {
            this.from = from;
            this.to = to;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"COUNT({from.ToExcelString()}{":"}{to.ToExcelString()})");
        }
    }
}
