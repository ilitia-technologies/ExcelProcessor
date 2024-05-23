using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheetWriter : IExcelSheet
    {
        /// <summary>
        /// Insert a string value at cursor position
        /// </summary>
        /// <param name="value">Value to insert</param>
        /// <param name="styleName">Style name. Optional</param>
        void InsertValue(string value, string styleName = null);

        /// <summary>
        /// Insert a decimal value at cursor position
        /// </summary>
        /// <param name="value">Value to insert</param>
        /// <param name="styleName">Style name. Optional</param>
        void InsertValue(decimal value, string styleName = null);

        /// <summary>
        /// Insert a integer value at cursor position
        /// </summary>
        /// <param name="value">Value to insert</param>
        /// <param name="styleName">Style name. Optional</param>
        void InsertValue(int value, string styleName = null);

        /// <summary>
        /// Insert a datetime value at cursor position
        /// </summary>
        /// <param name="value">Value to insert</param>
        /// <param name="styleName">Style name. Optional</param>
        void InsertValue(DateTime value, string styleName = null);

        /// <summary>
        /// Create a table with headers in first row
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        void CreateTableWithHeaders(ICellReference from, ICellReference to);

        /// <summary>
        /// Insert a formula at cursor position
        /// </summary>
        /// <param name="formula">Instance of <see cref="IFormula"/></param>
        /// <param name="styleName">Style name. Optional</param>
        void InsertFormula(IFormula formula, string styleName = null);

        /// <summary>
        /// Insert an image at cursor position
        /// </summary>
        /// <param name="imgData">Image byte array</param>
        /// <param name="imgDescription">Image descripcion</param>
        /// <param name="customWidth">Custom width. Optional</param>
        /// <param name="customHeight">Custom height. Optional</param>
        void InsertImage(byte[] imgData, string imgDescription, long? customWidth = null, long? customHeight = null);

        /// <summary>
        /// Merge rows at cursor position
        /// </summary>
        /// <param name="count">Number of rows to merge</param>
        /// <param name="styleName">Style name. Optional</param>
        void MergeRows(int count, string styleName = null);

        /// <summary>
        /// Merge columns at cursor position
        /// </summary>
        /// <param name="count">Number or columns to merge</param>
        /// <param name="styleName">Style name. Optional</param>
        void MergeColumns(int count, string styleName = null);

        /// <summary>
        /// Merge rows and columns at cursor position
        /// </summary>
        /// <param name="columns">Number of columns to merge</param>
        /// <param name="rows">Number of rows to merge</param>
        /// <param name="styleName">Style name. Optional</param>
        void Merge(int columns, int rows, string styleName = null);

        /// <summary>
        /// Sets height of a row (cursor position)
        /// </summary>
        /// <param name="rowHeight">Row height</param>
        void SetRowHeight(double rowHeight);

        /// <summary>
        /// Save changes
        /// </summary>
        void Save();
    }
}
