namespace ExcelProcessor.Abstractions.Generator.ReaderResults
{
    /// <summary>
    /// Excel reader error
    /// </summary>
    public interface IExcelReaderError
    {
        /// <summary>
        /// Indicates if it is a global error, it does not refer to a cell or a row
        /// </summary>
        public bool IsGlobalError { get; set; }

        /// <summary>
        /// Row that caused the error
        /// </summary>
        public int? RowNumError { get; set; }

        /// <summary>
        /// Error description
        /// </summary>
        public string ErrorDescription { get; set; }

        /// <summary>
        /// Cell that caused the error (Ej: B4)
        /// </summary>
        public string Cell { get; set; }
    }
}
