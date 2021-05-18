using Sayer.Excel.Cell;

namespace Sayer.Excel
{
    /// <summary>
    /// Represents a table in Excel
    /// </summary>
    public interface ITable
    {
        /// <summary>
        /// The name of the table
        /// </summary>
        string Name { get; }

        /// <summary>
        /// The upper-left coordinate of the table, including the obligatory header row
        /// </summary>
        CellLocation UpperLeft { get; }

        /// <summary>
        /// Returns the list of column headers
        /// </summary>
        string[] ColumnHeaders { get; }

        /// <summary>
        /// The number of rows in the table, not including the header row
        /// </summary>
        int RowCount { get; }

        /// <summary>
        /// Given a 1-based column index, return the cell content for every cell in the column.
        /// </summary>
        /// <param name="offsetColumn">
        /// The 1-based column index. The first column will be 1. In other words, this method will be invoked
        /// as if UpperLeft is defined as 1,1.
        /// </param>
        /// <returns>
        /// The default content of every cell in the column. Contents for particular cells within the column can be overwritten
        /// via the GetCellContent method, which takes precedence. Return null if there is no suitable default value
        /// for each column in the row.
        /// </returns>
        CellContent GetColumnContent(int offsetColumn);

        /// <summary>
        /// Given a location, returns the content of that cell. This will never be invoked for the header row.
        /// </summary>
        /// <param name="offsetLocation">
        /// The cell location. The first row/column pair for data will be 1,1. In other words, this method will be invoked
        /// as if there is no header and UpperLeft is defined as 1,1.
        /// </param>
        /// <returns>The content of the cell, or null if the value defined by GetColumnContent is suitable.</returns>
        CellContent GetCellContent(CellLocation offsetLocation);
    }
}
