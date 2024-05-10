using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Core.Generator.Engines;

namespace ExcelProcessor.Core.Generator
{
    public class ExcelGenerator : IExcelGenerator
    {
        public IExcelWriterEngine WriteFromTemplate<TExcelStyles>(string templateFilePath)
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>(templateFilePath);
        }

        public IExcelWriterEngine WriteFromEmptyFile<TExcelStyles>() 
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>();
        }

        public IExcelWriterEngine WriteFromByteArray<TExcelStyles>(byte[] data)
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>(data);
        }

        public IExcelReaderEngine ReadFromByteArray(byte[] data)
        {
            return new ExcelReaderEngine(data);
        }

        public IExcelReaderEngine ReadFromFile(string filePath)
        {
            return new ExcelReaderEngine(filePath);
        }
    }
}
