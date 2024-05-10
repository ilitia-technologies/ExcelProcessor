namespace ExcelProcessor.Example.Writer.DataContext
{
    internal class WriterDataContext
    {
        public string Title { get; set; }
        public string SubTitle { get; set; }
        public IEnumerable<UserInfo> Users { get; set; }
    }
}
