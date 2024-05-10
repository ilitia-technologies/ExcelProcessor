using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Exceptions;
using ExcelProcessor.Core.Pointers;
using ExcelProcessor.Examples.Reader.Reader.Entities;

namespace ExcelProcessor.Examples.Reader.Reader
{
    public class ExcelSheetParserExample : IExcelSheetParser<StudentContext>
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
            cursor = sheet.InitializeCursor(new CellReference(8, "A"));
            bool end = false;
            while (!end)
            {
                try
                {
                    Student student = new Student();
                    student.Name = sheet.ReadValue();
                    cursor.NextColumn();

                    student.LastName = sheet.ReadValue();
                    cursor.NextColumn();

                    if (string.IsNullOrEmpty(student.Name) && string.IsNullOrEmpty(student.LastName))
                        end = true;
                    else
                    {
                        if (string.IsNullOrEmpty(student.Name))
                            sheet.Results.AddRowError("Name is mandatory", cursor.CellRef.Row);
                        if (string.IsNullOrEmpty(student.LastName))
                            sheet.Results.AddRowError("Last name is mandatory", cursor.CellRef.Row);
                        student.Age = sheet.ReadValueAsInteger("Age is mandatory and must be a number");
                    }
                    sheet.Results.EntityReaded.Students.Add(student);
                    cursor.NextRowFromOrigin(); // Next row
                }
                catch (RowNotExistsException)
                {
                    // End of file
                    end = true;
                }
            }
        }


    }
}
