using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Sayer.Excel.Cell;
using Sayer.Excel.Workbook;
using Sayer.Excel.Worksheet;

namespace Sayer.Excel
{
    /// <summary>
    /// IO operations for Excel workbooks. Currently only allows for saving workbooks. Performance
    /// is poor for large workbooks, which might be easily fixed by using interop methods that
    /// work on ranges instead of individual cells.
    /// </summary>
    public static class IO
    {
        /// <summary>
        /// Saves an Excel workbook to a steam
        /// </summary>
        /// <param name="workbook">The workbook to save</param>
        /// <param name="tempFolder">The folder in which to create any required temporary files. If null, the Windows temp folder is used.</param>
        /// <returns>A stream. Caller is responsible for disposing it.</returns>
        public static MemoryStream SaveToStream(IWorkbook workbook, string tempFolder = null)
        {
            if (tempFolder == null)
            {
                tempFolder = Path.GetTempPath();
            }

            string tempFileName = Path.Combine(tempFolder, Path.GetRandomFileName() + ".xlsx");
            Save(workbook, tempFileName);

            try
            {
                using (FileStream fileStream = File.OpenRead(tempFileName))
                {
                    // Can't just return the fileStream because we are about to delete the file.
                    // So read it fully into a memory stream and return that.
                    MemoryStream memoryStream = new MemoryStream();
                    memoryStream.SetLength(fileStream.Length);
                    fileStream.Read(memoryStream.GetBuffer(), 0, (int)fileStream.Length);
                    return memoryStream;
                }
            }
            finally
            {
                File.Delete(tempFileName);
            }
        }

        /// <summary>
        /// Saves an Excel workbook to file
        /// </summary>
        /// <param name="workbook">The workbook to save</param>
        /// <param name="path">The full path of the file</param>
        /// <remarks>
        /// This is written to write one cell at a time. If saving a large workbook, this
        /// is painfully slow. I expect there are interop methods to save a range of cells,
        /// and that using those methods instead would result in a huge performance boost.
        /// </remarks>
        public static void Save(IWorkbook workbook, string path)
        {
            if (workbook.Worksheets == null || workbook.Worksheets.Count == 0)
            {
                throw new Exception("Workbook must have at least one worksheet");
            }

            using (var xlApp = new Marshaled<Application>(new Application()))
            {
                try
                {
                    using (var xlWorkBook = new Marshaled<Microsoft.Office.Interop.Excel.Workbook>(xlApp.Value.Workbooks.Add()))
                    {
                        try
                        {
                            List<IWorksheet> worksheets = workbook.Worksheets.Reverse().ToList();
                            var sheetAndTables = new List<Tuple<Microsoft.Office.Interop.Excel.Worksheet, IReadOnlyCollection<ITable>>>(worksheets.Count);

                            // First, create all the worksheets and tables. It's necessary to do this before we fill any cell data, because that data may
                            // reference other sheets and/or tables.
                            for (int i = 0; i < worksheets.Count; ++i)
                            {
                                try
                                {
                                    Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)(i == 0
                                        ? xlWorkBook.Value.Worksheets[1]
                                        : xlWorkBook.Value.Worksheets.Add());
                                    sheet.Name = worksheets[i].Name;

                                    if (worksheets[i].NonTableData != null)
                                    {
                                        foreach (KeyValuePair<CellLocation, CellContent> cellData in worksheets[i].NonTableData)
                                        {
                                            SetContent(cellData.Value, (Range)sheet.Cells[cellData.Key.Row, cellData.Key.Column]);
                                        }
                                    }

                                    foreach (ITable table in worksheets[i].Tables)
                                    {
                                        var bottomRight = new CellLocation(table.UpperLeft.Row + table.RowCount, table.UpperLeft.Column + table.ColumnHeaders.Length - 1);
                                        Range range = sheet.Range[sheet.Cells[table.UpperLeft.Row, table.UpperLeft.Column], sheet.Cells[bottomRight.Row, bottomRight.Column]];
                                        sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = table.Name;
                                    }

                                    sheetAndTables.Add(Tuple.Create(sheet, worksheets[i].Tables));
                                }
                                catch (Exception e)
                                {
                                    throw new Exception($"Failed to create worksheet {worksheets[i].Name}. Reason: {e.Message}", e);
                                }
                            }

                            // Now set the cell data
                            foreach (Tuple<Microsoft.Office.Interop.Excel.Worksheet, IReadOnlyCollection<ITable>> entry in sheetAndTables)
                            {
                                Microsoft.Office.Interop.Excel.Worksheet sheet = entry.Item1;

                                foreach (ITable table in entry.Item2)
                                {
                                    try
                                    {
                                        foreach (int col in Enumerable.Range(0, table.ColumnHeaders.Length))
                                        {
                                            sheet.Cells[table.UpperLeft.Row, table.UpperLeft.Column + col] = table.ColumnHeaders[col];
                                            CellContent content = table.GetColumnContent(col + 1);
                                            Range cells = sheet.Range[sheet.Cells[table.UpperLeft.Row + 1, table.UpperLeft.Column + col], sheet.Cells[table.UpperLeft.Row + table.RowCount, table.UpperLeft.Column + col]];
                                            SetContent(content, cells);

                                            foreach (int row in Enumerable.Range(1, table.RowCount))
                                            {
                                                content = table.GetCellContent(new CellLocation(row, col + 1));
                                                Range cell = (Range)sheet.Cells[table.UpperLeft.Row + row, table.UpperLeft.Column + col];
                                                SetContent(content, cell);
                                            }
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        throw new Exception($"Failed to create table {table.Name}. Reason: {e.Message}", e);
                                    }
                                }

                                sheet.Columns.AutoFit();
                            }

                            xlWorkBook.Value.ForceFullCalculation = true;
                            xlWorkBook.Value.SaveAs(path);
                        }
                        finally
                        {
                            xlWorkBook.Value.Close(false);
                        }
                    }
                }
                finally
                {
                    xlApp.Value.Quit();
                }
            }
        }

        private static void SetContent(CellContent cellContent, Range cellRange)
        {
            if (cellContent != null)
            {
                switch (cellContent.TypeOfContent)
                {
                    case CellContent.ContentType.Value:
                        cellRange.Value = cellContent.Content;
                        break;
                    case CellContent.ContentType.Formula:
                        cellRange.Formula = cellContent.Content;
                        break;
                    case CellContent.ContentType.FormulaLocal:
                        cellRange.FormulaLocal = cellContent.Content;
                        break;
                    case CellContent.ContentType.FormulaR1C1:
                        cellRange.FormulaR1C1 = cellContent.Content;
                        break;
                    case CellContent.ContentType.FormulaR1C1Local:
                        cellRange.FormulaR1C1Local = cellContent.Content;
                        break;
                    default:
                        throw new NotImplementedException($"ContentType {cellContent.TypeOfContent} not implemented.");
                }

                if (cellContent.NumberFormat != null)
                {
                    cellRange.NumberFormat = cellContent.NumberFormat;
                }
            }
        }

        private class Marshaled<T> : IDisposable
        {
            internal Marshaled(T item) { Value = item; }
            internal T Value { get; }
            public void Dispose() => Marshal.ReleaseComObject(Value);
        }
    }
}
