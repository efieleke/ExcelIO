namespace Sayer.Excel.Cell
{
    /// <summary>
    /// An Excel cell containing a formula
    /// </summary>
    public class CellFormula : CellContent
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="formula">Must be a string representing a legal Excel formula</param>
        public CellFormula(string formula) : this(formula, null) { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="formula">Must be a string representing a legal Excel formula</param>
        /// <param name="numberFormat">The number format for the cell (e.g. "0.00%"). Null means default.</param>
        public CellFormula(string formula, string numberFormat) : base(ContentType.Formula, formula, numberFormat) { }
    }
}
