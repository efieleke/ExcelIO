using System.Collections.Generic;
using Sayer.Excel.Worksheet;

namespace Sayer.Excel.Workbook
{
    /// <summary>
    /// Represents an Excel workbook containing just one table or worksheet
    /// </summary>
    public class SingleWorksheetWorkbook : IWorkbook
    {
        /// <summary>
        /// Constructs a workbook containing just one table
        /// </summary>
        /// <param name="table"></param>
        public SingleWorksheetWorkbook(ITable table) : this(new SingleTableWorksheet(table)) { }

        /// <summary>
        /// Constructs a workbook containing just one worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        public SingleWorksheetWorkbook(IWorksheet worksheet) { Worksheets = new[] { worksheet }; }

        /// <inheritdoc />
        public IReadOnlyList<IWorksheet> Worksheets { get; }
    }
}
