using System.Collections.Generic;
using Sayer.Excel.Worksheet;

namespace Sayer.Excel.Workbook
{
    /// <summary>
    /// Represents an Excel workbook
    /// </summary>
    public interface IWorkbook
    {
        /// <summary>
        /// The worksheets belonging to this workbook
        /// </summary>
        IReadOnlyList<IWorksheet> Worksheets { get; }
    }
}
