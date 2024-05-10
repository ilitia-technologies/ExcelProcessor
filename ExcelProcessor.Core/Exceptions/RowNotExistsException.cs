namespace ExcelProcessor.Core.Exceptions
{
    public class RowNotExistsException : Exception
    {
        public int RowNumber { get; set; }
        public RowNotExistsException(int rowNumber)
            : base()
        {
            RowNumber = rowNumber;
        }
    }
}
