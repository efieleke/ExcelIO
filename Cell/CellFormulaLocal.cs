namespace Sayer.Excel.Cell
{
    public class CellFormulaLocal : CellContent
    {
        public CellFormulaLocal(string formula) : this(formula, null)
        {
        }

        public CellFormulaLocal(string formula, string numberFormat) : base(ContentType.FormulaLocal, formula, numberFormat)
        {
        }
    }
}
