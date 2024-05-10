using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class Average : Formula
    {
        private readonly ICellReference from;
        private readonly ICellReference to;
        public Average(ICellReference from, ICellReference to)
        {
            this.from = from;
            this.to = to;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"AVERAGE({from.ToExcelString()}{":"}{to.ToExcelString()})");
        }
    }
}
