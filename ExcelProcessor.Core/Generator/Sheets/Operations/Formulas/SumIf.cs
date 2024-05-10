using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class SumIf : Formula
    {
        private readonly ICellReference fromCondition;
        private readonly ICellReference toCondition;
        private readonly string condition;
        private readonly ICellReference fromSum;
        private readonly ICellReference toSum;

        public SumIf(ICellReference fromCondition, ICellReference toCondition, string condition, ICellReference fromSum, ICellReference toSum)
        {
            this.fromCondition = fromCondition;
            this.toCondition = toCondition;
            this.condition = condition;
            this.fromSum = fromSum;
            this.toSum = toSum;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"SUMIF({fromCondition.ToExcelString()}:{toCondition.ToExcelString()},\"{condition}\",{fromSum.ToExcelString()}:{toSum.ToExcelString()})");
        }
    }
}
