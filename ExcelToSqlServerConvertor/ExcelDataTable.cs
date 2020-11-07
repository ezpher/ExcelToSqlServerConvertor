using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServerConvertor
{
    class ExcelDataTable
    {
        public DataTable DataTable { get; private set; }

        public ExcelDataTable()
        {
            DataTable = new DataTable();
        }

        public void FillDataTable(ref ExcelSpreadsheet spreadSheet, ref IEnumerable<Row> worksheetRows)
        {
            SetHeaderRows(ref spreadSheet, ref worksheetRows);
            SetDataRows(ref spreadSheet, ref worksheetRows);
        }


        private enum XMLFormats
        {
            DateShort = 14,
            DateLong = 165,
        }

        private void SetHeaderRows(ref ExcelSpreadsheet spreadSheet, ref IEnumerable<Row> worksheetRows)
        {
            // Set columns by getting the cell values in the first row, assuming that the first row has headers
            foreach (Cell cell in worksheetRows.ElementAt(0))
            {
                // add column names i.e. the cell names to the columns
                DataTable.Columns.Add(cell.CellValue.InnerXml);
            }
        }

        private void SetDataRows(ref ExcelSpreadsheet spreadSheet, ref IEnumerable<Row> worksheetRows)
        {
            //Write data to datatable; skip first row i.e. header row
            foreach (Row row in worksheetRows.Skip(1))
            {
                IEnumerable<Cell> rowCells = row.Descendants<Cell>();
                DataRow newRow = DataTable.NewRow();

                for (int i = 0; i < rowCells.Count(); i++)
                {
                    if (rowCells.ElementAt(i).CellValue != null && rowCells.ElementAt(i).DataType != null)
                    {

                        switch (rowCells.ElementAt(i).DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For strings, look up the value in the shared strings table and handle nulls
                                newRow[i] = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault()?
                                    .SharedStringTable
                                    .ElementAt(int.Parse(rowCells.ElementAt(i).InnerText)).InnerText;

                                break;

                            case CellValues.Boolean:
                                switch (rowCells.ElementAt(i).InnerText)
                                {
                                    case "0":
                                        newRow[i] = "FALSE";
                                        break;
                                    default:
                                        newRow[i] = "TRUE";
                                        break;
                                }

                                break;
                        }
                    }
                    else if (rowCells.ElementAt(i).CellValue != null && rowCells.ElementAt(i).DataType == null)
                    {
                        int styleIndex;
                        uint? formatId = null;

                        if (rowCells.ElementAt(i).StyleIndex != null)
                        {
                            styleIndex = (int)rowCells.ElementAt(i).StyleIndex.Value;

                            CellFormat cellFormat = spreadSheet.WorkbookPart.WorkbookStylesPart
                                .Stylesheet
                                .CellFormats
                                .ChildElements[int.Parse(rowCells.ElementAt(i).StyleIndex.InnerText)] as CellFormat;

                            formatId = cellFormat.NumberFormatId.Value;
                        }

                        // for dates, use the NumberFormatId on the CellFormat and convert the number in the Date cell to Date Format
                        if (formatId != null && formatId == (uint)XMLFormats.DateShort || formatId == (uint)XMLFormats.DateLong)
                        {
                            if (double.TryParse(rowCells.ElementAt(i).InnerText, out double oaDate))
                            {
                                newRow[i] = DateTime.FromOADate(oaDate).ToShortDateString();
                            }
                        }
                        else
                        {
                            // If the cell represents an integer number
                            newRow[i] = rowCells.ElementAt(i).CellValue.InnerXml;
                        }
                    }
                    else
                    {
                        newRow[i] = DBNull.Value;
                    }
                }

                DataTable.Rows.Add(newRow);
            }
        }

    }
}
