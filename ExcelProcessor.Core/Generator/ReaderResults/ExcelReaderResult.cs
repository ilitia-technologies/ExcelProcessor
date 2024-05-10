using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Core.Generator.ReaderResults
{
    public class ExcelReaderResult<TEntityReaded> : IExcelReaderResult<TEntityReaded>
        where TEntityReaded : class, new()
    {
        private readonly List<IExcelReaderError> errors = new List<IExcelReaderError>();

        public TEntityReaded EntityReaded { get; set; } = new TEntityReaded();
        public IEnumerable<IExcelReaderError> Errors
        {
            get => errors;
        }

        public bool HasErrors
        {
            get => errors != null && errors.Any();
        }

        public void AddGlobalError(string error)
        {
            errors.Add(new ExcelReaderError()
            {
                IsGlobalError = true,
                ErrorDescription = error,
            });
        }

        public void AddCellError(string error, ICellReference cellRef)
        {
            errors.Add(new ExcelReaderError()
            {
                RowNumError = cellRef.Row,
                ErrorDescription = error,
                Cell = cellRef?.ToExcelString()
            });
        }

        public void AddRowError(string error, int numLine)
        {
            errors.Add(new ExcelReaderError()
            {
                RowNumError = numLine,
                ErrorDescription = error,
            });
        }

    }
}
