using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Generator.Sheets.Operations.Formulas;
using ExcelProcessor.Core.Pointers;
using ExcelProcessor.Example.Writer.DataContext;

namespace ExcelProcessor.Example.Writer
{
    internal class ExcelSheetBuilderExample : IExcelSheetBuilder<WriterDataContext>
    {
        public string SheetName => "WriterSheetExample1";

        public void Build(IExcelSheetWriter sheet, WriterDataContext dataContext)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            InsertCustomFormats(sheet, 2);            

            InsertDataContextValues(sheet, dataContext, 6);
            
            InsertFormulas(sheet, dataContext, 9, 8);

            InsertImages(sheet, 18);
        }

        private void InsertCustomFormats(IExcelSheetWriter sheet, int initialRow)
        {
            IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "A"));
            sheet.InsertValue("This cell is red", ExampleExcelStyles.RedCell);
            cursor.NextColumn();
            sheet.InsertValue("This is green and use Arial font size 13", ExampleExcelStyles.GreenArialCell);
            cursor.NextRowFromOrigin();
            cursor.NextRowFromOrigin();
            sheet.InsertValue("This row has custom height: 40");
            sheet.SetRowHeight(40);
        }

        private void InsertDataContextValues(IExcelSheetWriter sheet, WriterDataContext dataContext, int initialRow)
        {
            // Write DataContext values
            IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "A"));
            sheet.InsertValue("Now DataContext values: (merge 4 columns and 2 rows)");
            sheet.Merge(4, 2, ExampleExcelStyles.BlueCell);
            cursor.NextRowFromOrigin();
            cursor.NextRowFromOrigin();

            sheet.InsertValue($"{dataContext.Title}: (merge 2 columns)");
            sheet.MergeColumns(2, ExampleExcelStyles.LightBlueCell);

            cursor.NextColumn();
            cursor.NextColumn();
            sheet.InsertValue($"{dataContext.SubTitle}: (merge 2 columns)");
            sheet.MergeColumns(2, ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            // User info
            sheet.InsertValue("Name", ExampleExcelStyles.HeaderTableCell);
            cursor.NextColumn();
            sheet.InsertValue("LastName", ExampleExcelStyles.HeaderTableCell);
            cursor.NextColumn();
            sheet.InsertValue("Age", ExampleExcelStyles.HeaderTableCell);
            cursor.NextColumn();
            sheet.InsertValue("Child count", ExampleExcelStyles.HeaderTableCell);
            cursor.NextRowFromOrigin();
            foreach (var user in dataContext.Users)
            {
                sheet.InsertValue(user.Name, ExampleExcelStyles.TableCell);
                cursor.NextColumn();
                sheet.InsertValue(user.LastName, ExampleExcelStyles.TableCell);
                cursor.NextColumn();
                sheet.InsertValue(user.Age, ExampleExcelStyles.TableCell);
                cursor.NextColumn();
                sheet.InsertValue(user.ChildCount, ExampleExcelStyles.TableCell);
                cursor.NextRowFromOrigin();
            }
        }

        private void InsertFormulas(IExcelSheetWriter sheet, WriterDataContext dataContext, int initialRow, int initialDataRow)
        {
            IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "F"));
            sheet.InsertValue("Formulas");
            sheet.MergeColumns(2, ExampleExcelStyles.HeaderTableCell);
            cursor.NextRowFromOrigin();

            ICellReference fromAge = new CellReference(initialDataRow, "C");
            ICellReference toAge = new CellReference(fromAge.Row + dataContext.Users.Count() - 1, fromAge.Column);

            ICellReference fromChildCount = fromAge.NextColumn();
            ICellReference toChildCount = new CellReference(fromChildCount.Row + dataContext.Users.Count() - 1, fromChildCount.Column);

            sheet.InsertValue("Average (children)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new Average(fromChildCount, toChildCount), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            sheet.InsertValue("Average (children with parent > 25)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new AverageIf(fromAge, toAge, ">25", fromChildCount, toChildCount), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            sheet.InsertValue("Sum (children)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new Sum(fromChildCount, toChildCount), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            sheet.InsertValue("Sum (children with parent > 25)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new SumIf(fromAge, toAge, ">25", fromChildCount, toChildCount), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            sheet.InsertValue("Count (parents)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new Count(fromChildCount, toChildCount), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();

            sheet.InsertValue("Count (parents with age > 25)", ExampleExcelStyles.BlueCell);
            cursor.NextColumn();
            sheet.InsertFormula(new CountIf(fromAge, toAge, ">25"), ExampleExcelStyles.LightBlueCell);
            cursor.NextRowFromOrigin();
        }

        private void InsertImages(IExcelSheetWriter sheet, int initialRow)
        {
            IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "A"));
            sheet.InsertValue("Insert a image:");
            sheet.MergeColumns(2, ExampleExcelStyles.Orange);
            cursor.NextRowFromOrigin();
            sheet.SetRowHeight(60);
            sheet.InsertValue("Custom size", ExampleExcelStyles.LightOrange);
            sheet.SetRowHeight(120);
            cursor.NextColumn();
            sheet.InsertImage(File.ReadAllBytes("Resources\\IlitiaLogo.png"), "Ilitia logo", 180, 121);
            cursor.NextRowFromOrigin();
            sheet.InsertValue("Full size", ExampleExcelStyles.LightOrange);
            sheet.SetRowHeight(120);
            cursor.NextColumn();
            sheet.InsertImage(File.ReadAllBytes("Resources\\IlitiaLogo.png"), "Ilitia logo");
            cursor.NextColumn();
        }
    }
}
