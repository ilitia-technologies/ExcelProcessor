using DocumentFormat.OpenXml.Packaging;
using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Core.Generator.Sheets.Operations;
using System.Reflection;

namespace ExcelProcessor.Core.Generator.Engines
{
    public class ExcelWriterEngine<ExcelStyles> : ExcelEngineBase, IExcelWriterEngine, IDisposable
        where ExcelStyles : IExcelStyles, new()
    {
        private ExcelStyles styles;

        internal ExcelWriterEngine(string template)
        {
            if (!File.Exists(template))
                throw new IOException($"ExcelEngine: File {template} not found");

            byte[] templateContent = File.ReadAllBytes(template);
            LoadFrom(templateContent);
        }

        internal ExcelWriterEngine(byte[] data)
        {
            LoadFrom(data);
        }

        internal ExcelWriterEngine()
        {
            byte[] emptyFile = ReadEmptyFile();
            LoadFrom(emptyFile);
        }

        private byte[] ReadEmptyFile()
        {
            Assembly assembly = Assembly.GetAssembly(GetType());
            using (Stream stream = assembly.GetManifestResourceStream($"{assembly.GetName().Name}.Resources.Empty.xlsx"))
            {
                byte[] data = new byte[stream.Length];
                stream.Read(data, 0, (int)stream.Length);
                return data;
            }            
        }

        protected override void LoadFrom(byte[] data)
        {
            base.LoadFrom(data);

            styles = new ExcelStyles();
            styles.Inyect(spreadSheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet);
        }

        public byte[] Create<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext)
            where TDataContext : class
        {
            Generate(sheetBuilders, dataContext);
            return GetFileAndClose();
        }

        public void CreateAndCopy<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext, string outputFile)
            where TDataContext : class
        {
            Generate(sheetBuilders, dataContext);
            SaveAs(outputFile);
        }

        private void Generate<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext)
            where TDataContext : class
        {
            if (sheetBuilders != null)
            {
                foreach (IExcelSheetBuilder<TDataContext> sheetBuilder in sheetBuilders)
                {
                    sheetBuilder.Build(GetSheet(sheetBuilder.SheetName), dataContext);
                }
            }
        }


        private IExcelSheetWriter GetSheet(string sheetName)
        {
            WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
            if (worksheetPart == null)
                return null;

            return new ExcelSheetWriter(worksheetPart, styles);
        }

        private byte[] GetFileAndClose()
        {
            spreadSheetDocument.Save();

            documentMs.Position = 0;
            return documentMs.ToArray();
        }

        private void SaveAs(string outputFile)
        {
            spreadSheetDocument.Save();

            documentMs.Position = 0;
            File.WriteAllBytes(outputFile, documentMs.ToArray());
        }

    }
}
