using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.ReaderResults
{
    /// <summary>
    /// Reader results
    /// </summary>
    /// <typeparam name="TEntityReaded">Underlying data type</typeparam>
    public interface IExcelReaderResult<TEntityReaded>
        where TEntityReaded : class
    {
        /// <summary>
        /// Errors detectged
        /// </summary>
        IEnumerable<IExcelReaderError> Errors { get; }

        /// <summary>
        /// Data readed
        /// </summary>
        public TEntityReaded EntityReaded { get; }

        /// <summary>
        /// Add a global error
        /// </summary>
        /// <param name="error">Error message</param>
        void AddGlobalError(string error);

        /// <summary>
        /// Add cell error
        /// </summary>
        /// <param name="error">Error message</param>
        /// <param name="cellRef">Cell reference</param>
        void AddCellError(string error, ICellReference cellRef);

        /// <summary>
        /// Add row error
        /// </summary>
        /// <param name="error">Error message</param>
        /// <param name="numLine">Line source of error</param>
        void AddRowError(string error, int numLine);
    }
}
