using ExcelProcessor.Abstractions.Generator;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Core.Generator;
using ExcelProcessor.Examples.Reader.Reader;
using ExcelProcessor.Examples.Reader.Reader.Entities;

namespace ExcelProcessor.Examples.Reader
{
    public class ReadExampleTest
    {
        [Fact]
        public void ReadOkTest()
        {
            IExcelGenerator excelGenerator = new ExcelGenerator();
            using (IExcelReaderEngine readerEngine = excelGenerator.ReadFromFile("Resources\\ReaderExampleOk.xlsx"))
            {
                IExcelReaderResult<StudentContext> result = readerEngine.ReadFile(new IExcelSheetParser<StudentContext>[]
                    {
                        new ExcelSheetParserExample()
                    });

                ValidateOk(result);
            }
        }

        [Fact]
        public void ReadOkInParallelTest()
        {
            IExcelGenerator excelGenerator = new ExcelGenerator();
            using (IExcelReaderEngine readerEngine = excelGenerator.ReadFromFile("Resources\\ReaderExampleOk.xlsx"))
            {
                IExcelReaderResult<StudentContext> result = readerEngine.ReadFile(new IExcelSheetParser<StudentContext>[]
                    {
                        new ExcelSheetParallelParserExample()
                    });
                ValidateOk(result);
            }
        }

        private void ValidateOk(IExcelReaderResult<StudentContext> result)
        {
            Assert.NotNull(result);

            // Without errors
            Assert.Empty(result.Errors);

            // Data is ok
            Assert.NotNull(result.EntityReaded);
            Assert.Equal("MIT", result.EntityReaded.University);
            Assert.Equal(new DateTime(2024, 04, 25, 11, 5, 14, DateTimeKind.Utc), result.EntityReaded.GeneratedAt);
            Assert.Equal(6, result.EntityReaded.Students.Count);
            ValidateStudent(result.EntityReaded.Students.ElementAt(0), "Sebastian", "Gil", 22);
            ValidateStudent(result.EntityReaded.Students.ElementAt(1), "Pedro Jose", "Castro", 25);
            ValidateStudent(result.EntityReaded.Students.ElementAt(2), "Inmaculada", "Amat", 21);
            ValidateStudent(result.EntityReaded.Students.ElementAt(3), "Margarita", "Palma", 33);
            ValidateStudent(result.EntityReaded.Students.ElementAt(4), "Isidoro", "Lobo", 34);
            ValidateStudent(result.EntityReaded.Students.ElementAt(5), "Marcela", "Prado", 29);
        }

        [Fact]
        public void ReadWithErrorsTest()
        {
            IExcelGenerator excelGenerator = new ExcelGenerator();
            using (IExcelReaderEngine readerEngine = excelGenerator.ReadFromFile("Resources\\ReaderExampleErrors.xlsx"))
            {
                IExcelReaderResult<StudentContext> result = readerEngine.ReadFile(new IExcelSheetParser<StudentContext>[]
                    {
                        new ExcelSheetParserExample()
                    });

                Assert.NotNull(result);

                // Errors detected
                Assert.NotEmpty(result.Errors);
                Assert.Equal(3, result.Errors.Count());

                // Errors
                Assert.Equal("B4", result.Errors.ElementAt(0).Cell);
                Assert.Equal("Generated at is not a valid date", result.Errors.ElementAt(0).ErrorDescription);

                Assert.Equal(8, result.Errors.ElementAt(1).RowNumError);
                Assert.Equal("Last name is mandatory", result.Errors.ElementAt(1).ErrorDescription);

                Assert.Equal(10, result.Errors.ElementAt(2).RowNumError);
                Assert.Equal("Age is mandatory and must be a number", result.Errors.ElementAt(2).ErrorDescription);

                // Data ok
                Assert.NotNull(result.EntityReaded);
                Assert.Equal("MIT", result.EntityReaded.University);
                Assert.Equal(6, result.EntityReaded.Students.Count);
                ValidateStudent(result.EntityReaded.Students.ElementAt(0), "Sebastian", null, 22);
                ValidateStudent(result.EntityReaded.Students.ElementAt(1), "Pedro Jose", "Castro", 25);
                ValidateStudent(result.EntityReaded.Students.ElementAt(2), "Inmaculada", "Amat", 0);
                ValidateStudent(result.EntityReaded.Students.ElementAt(3), "Margarita", "Palma", 33);
                ValidateStudent(result.EntityReaded.Students.ElementAt(4), "Isidoro", "Lobo", 34);
                ValidateStudent(result.EntityReaded.Students.ElementAt(5), "Marcela", "Prado", 29);
            }
        }

        private void ValidateStudent(Student student, string expectedName, string expectedLastName, int expectedAge)
        {
            Assert.NotNull(student);
            Assert.Equal(expectedName, student.Name);
            Assert.Equal(expectedLastName, student.LastName);
            Assert.Equal(expectedAge, student.Age);
        }
    }
}