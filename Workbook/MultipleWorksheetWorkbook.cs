using System.Collections.Generic;
using Sayer.Excel.Worksheet;

namespace Sayer.Excel.Workbook
{
    /// <summary>
    /// Represents an Excel workbook with multiple sheets
    /// </summary>
    public class MultipleWorksheetWorkbook : IWorkbook
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheets">The sheets in the workbook</param>
        public MultipleWorksheetWorkbook(IReadOnlyList<IWorksheet> sheets)
        {
            Worksheets = sheets;
        }

        /// <inheritdoc />
        public IReadOnlyList<IWorksheet> Worksheets { get; }
    }
}
