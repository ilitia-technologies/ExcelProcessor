# ExcelProcessor

Wrapper over *OpenXML* for easy use. No paid license required. It offers the functionality to read and generate *Excel* files.

## Positioning
Position and movement through *Excel* is done through two abstractions:
- *ICellReference*. It defines a position: row and column
- *IRowCurso*r. Cursor with the ability to change cell reference position. The first position is marked as origin
  - *NextColumn()*: It allows to move to the next column
  - *NextRowFromOrigin()*: It allows positioning in the next row and origin column 

## Write operations
Steps:
1. Create a instance of *ExcelGenerator*. It has methods to generate an instance of *IExcelWriterEngine* from:
   - *Excel* Template file
   - Byte array that represents a *Excel* file
   - Empty file
  
2. Define *Excel* styles. You need a class that inherits from *ExcelStyles*. The objective is the definition of styles (*InjectStyle* method).
   Methods are available to:
   - Inyect fill, border and fonts
   - Generate solid fill with full border, it can be customized.

3. Use *IExcelWriterEngine* to perform operations and get Excel file as byte array (*Create* method) or save it to disk (*CreateAndSave* method). Two main parameters:
   - *TDataContext*. Data context. It will contain the data that you want to write in *Excel*
   - *IExcelSheetBuilder<TDataContext>[]*. Each instance of *IExcelSheetBuilder* must contain:
     - *SheetName*: the name of the *Excel* sheet referenced.
     - Implementation of the *Build* method. The operations to be performed on the *Excel* sheet will be indicated.

4. Implement *IExcelSheetBuilder*. Build method usually follows the following steps:
   - Initialize cursor in a cell
   - Perform operations in the cell: InsertValue, InsertFormula, InsertImage, Merge, SetRowHeight, etc
   - User cursor to move to another position: NextColumn or NextRowFromOrigin
   - Continue performing operations in a cell
  
For further information, see the complete example: project *ExcelProcessor.Examples.Writer*
```C#
IExcelGenerator excelGenerator = new ExcelGenerator();
using (IExcelWriterEngine writerEngine = excelGenerator.FromTemplate<ExampleExcelStyles>("Resources\\WriterTemplateExample.xlsx"))
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
```

Partial content of *ExcelSheetBuilderExample*
```C#

public void Build(IExcelSheetWriter sheet, WriterDataContext dataContext)
{
    InsertCustomFormats(sheet, 2);
    InsertDataContextValues(sheet, dataContext, 6);
    InsertFormulas(sheet, dataContext, 9, 8); // Available in source code
    InsertImages(sheet, 18); // Available in source code
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
```

### Extensibility
The system can be expanded in two ways:
- New formulae. It would be necessary to define more classes that inherit from *Formula* and implement the *Build* method with the necessary operations.
- Other operations at row and/or column level. The extension point is in the *IExcelSheetWriter* interface

## Read operations
Steps:
1. Create a instance of *ExcelGenerator*. It has methods to generate an instance of *IExcelReaderEngine* from:
   - *Excel* Template file
   - Byte array that represents a *Excel* file

2. Use *IExcelReaderEngine* to execute *ReadFile* method and get results as instance of *IExcelReaderResult<TEntityReaded>*. *TEntityReaded* must be a class with a constructor without parameters.
   *ReadFile* has one parameter
   - *IExcelSheetParser<TEntityReaded>[]*. Each instance of *IExcelSheetParser* must contain:
     - *SheetName*: the name of the Excel sheet referenced.
     - Implementation of the *Parse* method. The operations to be performed on the *Excel* sheet will be indicated.

3. Implement *IExcelSheetParser*. Parse method usually follows these steps:
   - Initialize cursor in a cell
   - Perform reading operations on the cell: *ReadValue*, *ReadValueAsDateTime*, *ReadValueAsInteger*, etc
   - User cursor to move to another position: *NextColumn* or *NextRowFromOrigin*
   - Continue performing operations on cell to load *TEntityReaded* instance
   - If a error is found you can register it as *CellError* (with cell reference), *RowError* (with row reference) or *GlobalError* (fatal error)

4. Validate *IExcelReaderResult<TEntityReaded>*. You can:
   - Check if it has errors: *result.Errors*
   - Get *TEntityReaded*. You probably want to validate the entity data

For further information, see the complete example: project *ExcelProcessor.Examples.Reader*

```C#
IExcelGenerator excelGenerator = new ExcelGenerator();
using (IExcelReaderEngine readerEngine = excelGenerator.ReadFromFile("Resources\\ReaderExampleOk.xlsx"))
{
    IExcelReaderResult<StudentContext> result = readerEngine.ReadFile(new IExcelSheetParser<StudentContext>[]
        {
            new ExcelSheetParserExample()
        });

    // Do operations with result: validate or process StudentContext instance
    // Example
    // No errors expected
    // Assert.Empty(result.Errors);
    // Entity must be readed
    // Assert.NotNull(result.EntityReaded);
    // Check expected value
    // Assert.Equal("MIT", result.EntityReaded.University);
}
```

Partial content of *ExcelSheetParserExample*
```C#
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
```

### Read as parallel
With the *IExcelSheetReader<TEntityReaded>.ReadInParallel* method you can perform a block reading line by line in parallel. 
It allows a better performance in iterative reading of rows.
It has 3 parameters:
- *StartsAtRow*: Row number where to start reading
- *ColumnCount*: Number of columns to read
- Func. Output: Entity readed. Inputs:
    - *string[]*. Line data readed as raw string. Length will be equal to *ColumnCount*.
    - *uint*. Row index readed

Partial content of *ExcelSheetParallelParserExample*
```C#
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
```

### Extensibility
The most obvious line of work to extend reading functionalities focuses on the extension of native *.NET* data types that can be read, avoiding external conversions from string to the desired type.
Currently the reading of integers, dates, text strings and flags as Yes/No is allowed. New methods in *IExcelSheetReader<TEntityReaded>* would allow these actions to be performed



