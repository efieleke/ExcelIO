namespace Sayer.Excel.Cell
{
    public class CellFormulaR1C1 : CellContent
    {
        public CellFormulaR1C1(string formula) : this(formula, null)
        {
        }

        public CellFormulaR1C1(string formula, string numberFormat) : base(ContentType.FormulaR1C1, formula, numberFormat)
        {
        }
    }
}
