using System.Collections.Generic;
using Sayer.Excel.Cell;

namespace Sayer.Excel.Worksheet
{
    /// <summary>
    /// Represents a worksheet that contains a single table
    /// </summary>
    public class SingleTableWorksheet : IWorksheet
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="table">The single table within this sheet. This sheet's name is set to the table's name.</param>
        public SingleTableWorksheet(ITable table)
        {
            Name = table.Name;
            Tables = new[] { table };
        }

        /// <inheritdoc />
        public string Name { get; }

        /// <inheritdoc />
        public IReadOnlyCollection<ITable> Tables { get; }

        /// <inheritdoc />
        public IReadOnlyDictionary<CellLocation, CellContent> NonTableData => null;
    }
}
