using System.Collections.Generic;
using Sayer.Excel.Cell;

namespace Sayer.Excel.Worksheet
{
    /// <summary>
    /// Represents a worksheet that holds multiple tables
    /// </summary>
    public class MultipleTableWorksheet : IWorksheet
    {
        /// <summary>
        /// Constructs a worksheet
        /// </summary>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="tables">The collection of tables</param>
        public MultipleTableWorksheet(string sheetName, IReadOnlyCollection<ITable> tables)
        {
            Name = sheetName;
            Tables = tables;
        }

        /// <inheritdoc />
        public string Name { get; }

        /// <inheritdoc />
        public IReadOnlyCollection<ITable> Tables { get; }

        /// <inheritdoc />
        public IReadOnlyDictionary<CellLocation, CellContent> NonTableData => null;
    }
}
