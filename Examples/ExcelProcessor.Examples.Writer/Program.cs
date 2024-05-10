using ExcelProcessor.Abstractions.Generator;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Core.Generator;
using ExcelProcessor.Example.Writer;
using ExcelProcessor.Example.Writer.DataContext;

namespace ExcelProcessor.Example
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Executing examples...");

            WriterExample(new ExcelGenerator());

            Console.WriteLine("Done!");
        }

        private static void WriterExample(IExcelGenerator excelGenerator)
        {
            Console.WriteLine("WriterExample...");

            using (IExcelWriterEngine writerEngine = excelGenerator.WriteFromTemplate<ExampleExcelStyles>("Resources\\WriterTemplateExample.xlsx"))
            {
                // Get Excel byte-array
                byte[] excelBytes = writerEngine.Create(new IExcelSheetBuilder<WriterDataContext>[]
                {
                    new ExcelSheetBuilderExample()
                },
                CreateDataContext());

                // Write to output
                File.WriteAllBytes("WriterTemplateOutput.xlsx", excelBytes);
            }
        }

        private static WriterDataContext CreateDataContext()
        {
            return new WriterDataContext()
            {
                Title = "DataContext Title. Example",
                SubTitle = "DataContext SubTitle. Another example",
                Users = new UserInfo[]
                    {
                        new UserInfo(){ Name = "Henry", LastName = "Spencer", Age = 22, ChildCount = 2},
                        new UserInfo(){ Name = "Keyla", LastName = "Brewer", Age = 35, ChildCount = 1},
                        new UserInfo(){ Name = "Juliana", LastName = "Johns", Age = 19, ChildCount = 3},
                        new UserInfo(){ Name = "Elliott", LastName = "Meadows", Age = 64, ChildCount = 2},
                        new UserInfo(){ Name = "Marina", LastName = "Pitts", Age = 51, ChildCount = 4},
                    }
            };
        }
    }
}