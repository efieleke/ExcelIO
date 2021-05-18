namespace Sayer.Excel.Cell
{
    /// <summary>
    /// An Excel cell containing a value
    /// </summary>
    public class CellValue : CellContent
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Must be a type Excel can digest</param>
        public CellValue(object value) : this(value, null) { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Must be a type Excel can digest</param>
        /// <param name="numberFormat">The number format for the cell (e.g. "0.00%"). Null means default.</param>
        public CellValue(object value, string numberFormat) : base(ContentType.Value, value, numberFormat) { }
    }
}
