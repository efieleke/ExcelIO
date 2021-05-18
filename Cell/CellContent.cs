namespace Sayer.Excel.Cell
{
    /// <summary>
    /// Represents the content of an Excel cell
    /// </summary>
    public abstract class CellContent
    {
        /// <summary>
        /// Generates cell content for a double, treating NaN as an error.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static CellContent Create(double value) => Create(value, null);

        /// <summary>
        /// Generates cell content for a double, treating NaN as an error.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="numberFormat">The number format for the cell (e.g. "0.00%"). Null means default.</param>
        /// <returns></returns>
        public static CellContent Create(double value, string numberFormat) =>
            double.IsNaN(value) ? (CellContent)new CellFormula("=NA()") : new CellValue(value, numberFormat);

        /// <summary>
        /// Possible cell content types
        /// </summary>
        public enum ContentType { Value, Formula, FormulaLocal, FormulaR1C1, FormulaR1C1Local }

        /// <summary>
        /// Type of data in the cell
        /// </summary>
        public ContentType TypeOfContent { get; }

        /// <summary>
        /// The content of this cell
        /// </summary>
        public object Content { get; }

        /// <summary>
        /// The number format for the cell (e.g. "0.00%"). Null means default.
        /// </summary>
        public string NumberFormat { get; }

        /// <summary>
        /// Constructs cell content
        /// </summary>
        /// <param name="contentType">type of content in the cell</param>
        /// <param name="content">the cell content, which must be a type Excel can digest</param>
        /// <param name="numberFormat">The number format for the cell (e.g. "###,##.00%"). Null means default.</param>
        protected CellContent(ContentType contentType, object content, string numberFormat)
        {
            TypeOfContent = contentType;
            Content = content;
            NumberFormat = numberFormat;
        }
    }
}
