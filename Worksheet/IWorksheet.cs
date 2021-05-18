using System.Collections.Generic;
using Sayer.Excel.Cell;

namespace Sayer.Excel.Worksheet
{
    /// <summary>
    /// Represents an Excel worksheet
    /// </summary>
    public interface IWorksheet
    {
        /// <summary>
        /// The name of the worksheet, as displayed on its tab
        /// </summary>
        string Name { get; }

        /// <summary>
        /// The collection of tables that this worksheet contains
        /// </summary>
        IReadOnlyCollection<ITable> Tables { get; }

        /// <summary>
        /// Non-table data that this worksheet contains. If none, return null (or an empty dictionary).
        /// </summary>
        IReadOnlyDictionary<CellLocation, CellContent> NonTableData { get; }
    }
}
