using System;

namespace Sayer.Excel.Cell
{
    /// <summary>
    /// Represents the 1-based location of an Excel cell
    /// </summary>
    public struct CellLocation
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="row">Must be > 0</param>
        /// <param name="column">Must be > 0</param>
        public CellLocation(int row, int column)
        {
            if (row < 1 || column < 1)
            {
                throw new ArgumentException("Row and column must be greater than 0");
            }

            Row = row;
            Column = column;
        }

        /// <summary>
        /// 1-based index of the cell's row
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 1-based index of the cell's column
        /// </summary>
        public int Column { get; }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            if (!(obj is CellLocation))
            {
                return false;
            }

            CellLocation location = (CellLocation)obj;
            return Row == location.Row && Column == location.Column;
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            int hashCode = 656739706;
            hashCode = hashCode * -1521134295 + Row.GetHashCode();
            hashCode = hashCode * -1521134295 + Column.GetHashCode();
            return hashCode;
        }
    }
}
