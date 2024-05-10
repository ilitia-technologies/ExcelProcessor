using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class CountIf : Formula
    {
        private readonly ICellReference fromCondition;
        private readonly ICellReference toCondition;
        private readonly string condition;

        public CountIf(ICellReference fromCondition, ICellReference toCondition, string condition)
        {
            this.fromCondition = fromCondition;
            this.toCondition = toCondition;
            this.condition = condition;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"COUNTIF({fromCondition.ToExcelString()}:{toCondition.ToExcelString()},\"{condition}\")");
        }
    }
}
