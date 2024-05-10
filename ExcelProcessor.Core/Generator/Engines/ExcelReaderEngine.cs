using DocumentFormat.OpenXml.Packaging;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Core.Generator.ReaderResults;
using ExcelProcessor.Core.Generator.Sheets.Operations;

namespace ExcelProcessor.Core.Generator.Engines
{
    public class ExcelReaderEngine : ExcelEngineBase, IExcelReaderEngine
    {

        internal ExcelReaderEngine(byte[] data)
        {
            LoadFrom(data);
        }

        internal ExcelReaderEngine(string file)
        {
            if (!File.Exists(file))
                throw new IOException($"ExcelEngine: File {file} not found");

            byte[] templateContent = File.ReadAllBytes(file);
            LoadFrom(templateContent);
        }

        public IExcelReaderResult<TEntityReaded> ReadFile<TEntityReaded>(IExcelSheetParser<TEntityReaded>[] sheetParsers) 
            where TEntityReaded : class, new()
        {
            IExcelReaderResult<TEntityReaded> result = new ExcelReaderResult<TEntityReaded>();
            if (sheetParsers != null)
            {
                foreach (IExcelSheetParser<TEntityReaded> parser in sheetParsers)
                    parser.Parse(GetSheet(parser.SheetName, result));
            }
            return result;
        }

        public IExcelSheetReader<TEntityReaded> GetSheet<TEntityReaded>(string sheetName, IExcelReaderResult<TEntityReaded> results)
            where TEntityReaded: class, new()
        {
            WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
            if (worksheetPart == null)
                return null;

            return new ExcelSheetReader<TEntityReaded>(spreadSheetDocument.WorkbookPart, worksheetPart, results);
        }
    }
}
