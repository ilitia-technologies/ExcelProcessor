using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class Rows : Formula
    {
        private readonly ICellReference from;
        private readonly ICellReference to;
        public Rows(ICellReference from, ICellReference to)
        {
            this.from = from;
            this.to = to;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"ROWS({from.ToExcelString()}{":"}{to.ToExcelString()})");
        }
    }
}
