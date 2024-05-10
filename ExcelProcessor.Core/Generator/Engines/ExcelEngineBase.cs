using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelProcessor.Core.Generator.Engines
{
    public abstract class ExcelEngineBase : IDisposable
    {
        protected SpreadsheetDocument spreadSheetDocument;
        protected MemoryStream documentMs;

        protected virtual void LoadFrom(byte[] data)
        {
            documentMs = new MemoryStream();
            documentMs.Write(data, 0, data.Length);
            spreadSheetDocument = SpreadsheetDocument.Open(documentMs, true);
        }

        protected WorksheetPart GetWorksheetPart(string sheetName)
        {
            IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<Sheet>()
                                                                                 .Where(s => s.Name == sheetName);
            if (sheets?.Count() == 0)
                return null;

            string relationshipId = sheets?.First().Id.Value;
            return (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
        }


        public void Dispose()
        {
            spreadSheetDocument.Dispose();
            documentMs.Dispose();
        }
    }
}
