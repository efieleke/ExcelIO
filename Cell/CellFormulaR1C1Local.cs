namespace Sayer.Excel.Cell
{
    public class CellFormulaR1C1Local : CellContent
    {
        public CellFormulaR1C1Local(string formula) : this(formula, null)
        {
        }

        public CellFormulaR1C1Local(string formula, string numberFormat) : base(ContentType.FormulaR1C1Local, formula, numberFormat)
        {
        }
    }
}
