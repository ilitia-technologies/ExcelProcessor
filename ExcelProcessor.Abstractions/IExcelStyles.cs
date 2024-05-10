using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelProcessor.Abstractions
{
    public interface IExcelStyles
    {
        /// <summary>
        /// Entry point to inyect styles
        /// </summary>
        /// <param name="stylesSheet">Reference to Excel Stylesheet</param>
        void Inyect(Stylesheet stylesSheet);

        /// <summary>
        /// Gets identifier of style
        /// </summary>
        /// <param name="styleName">Name of style</param>
        /// <returns></returns>
        uint GetStyleId(string styleName);
    }
}
