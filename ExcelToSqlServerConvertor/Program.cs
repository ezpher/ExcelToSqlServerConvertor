using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServerConvertor
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create an OpenXMLSpreadsheet
                ExcelSpreadsheet spreadSheet = new ExcelSpreadsheet();

                // Fill the spreadsheet with the data from opened Excel Spreadsheet
                string filePath = @"C:\_MyDotNetApplications\ExcelToSqlServerConvertor\Test.xlsx";
                spreadSheet.FillSpreadSheet(filePath);

                // Initialize ConnectionString to Db
                string connectionStr = ConfigurationManager.ConnectionStrings["ExcelToSqlServerConvertorTestDb"].ToString();

                // use sql bulk copy for bulk sql operations to db; set keep identity if you want to keep the identity values of the source table
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionStr, SqlBulkCopyOptions.KeepIdentity))
                {
                    try
                    {
                        // for each worksheet in spreadsheet fill datatable and write datatable to sql server
                        foreach (KeyValuePair<string, Worksheet> namedWorksheet in ExcelSpreadsheet.GetNamedWorksheets(spreadSheet.WorkbookPart))
                        {
                            IEnumerable<Row> worksheetRows = ExcelSpreadsheet.GetWorkSheetRows(namedWorksheet.Value);
                            var dataTable = new ExcelDataTable();
                            dataTable.FillDataTable(ref spreadSheet, ref worksheetRows);

                            bulkCopy.DestinationTableName = namedWorksheet.Key;
                            bulkCopy.WriteToServer(dataTable.DataTable);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception thrown when using connecting or writing to server: " + ex.Message + "\n");
                        Console.WriteLine(ex.StackTrace);
                        Console.Read();
                    }
                }

                spreadSheet.SpreadSheetDocument.Close();
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
