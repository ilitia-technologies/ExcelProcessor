using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Exceptions;
using ExcelProcessor.Core.Pointers;
using ExcelProcessor.Examples.Reader.Reader.Entities;

namespace ExcelProcessor.Examples.Reader.Reader
{
    public class ExcelSheetParallelParserExample : IExcelSheetParser<StudentContext>
    {
        public string SheetName => "Students";

        public void Parse(IExcelSheetReader<StudentContext> sheet)
        {
            IRowCursor cursor = sheet.InitializeCursor(new CellReference(3, "B"));
            
            sheet.Results.EntityReaded.University = sheet.ReadValue();
            // University (mandatory)
            if (string.IsNullOrEmpty(sheet.Results.EntityReaded.University))
                sheet.Results.AddCellError("University is mandatory", cursor.CellRef);
            cursor.NextRowFromOrigin();

            // GeneratedAt
            sheet.Results.EntityReaded.GeneratedAt = sheet.ReadValueAsDateTime("Generated at is not a valid date");

            // Read students
            sheet.Results.EntityReaded.Students = sheet.ReadInParallel(8, 3,
                (string[] rowData, uint rowIndex) =>
                {
                    Student student = new Student();
                    student.Name = rowData[0];
                    student.LastName = rowData[1];
                    
                    if (string.IsNullOrEmpty(student.Name))
                        sheet.Results.AddRowError("Name is mandatory", (int)rowIndex);
                    if (string.IsNullOrEmpty(student.LastName))
                        sheet.Results.AddRowError("Last name is mandatory", (int)rowIndex);

                    if (int.TryParse(rowData[2], out int ageAsInteger))
                        student.Age = ageAsInteger;
                    else
                        sheet.Results.AddRowError("Age is mandatory and must be a number", (int)rowIndex);
                    return student;
                })?.ToList();
        }


    }
}
