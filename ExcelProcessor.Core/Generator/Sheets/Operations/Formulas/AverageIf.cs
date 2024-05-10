using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations.Formulas
{
    public class AverageIf : Formula
    {
        private readonly ICellReference fromCondition;
        private readonly ICellReference toCondition;
        private readonly string condition;
        private readonly ICellReference fromAvg;
        private readonly ICellReference toAvg;
        public AverageIf(ICellReference fromCondition, ICellReference toCondition, string condition, ICellReference fromAvg, ICellReference toAvg)
        {
            this.fromCondition = fromCondition;
            this.toCondition = toCondition;
            this.condition = condition;
            this.fromAvg = fromAvg;
            this.toAvg = toAvg;
        }

        public override CellFormula Build()
        {
            return new CellFormula($"AVERAGEIF({fromCondition.ToExcelString()}:{toCondition.ToExcelString()},\"{condition}\",{fromAvg.ToExcelString()}:{toAvg.ToExcelString()})");
        }
    }
}
