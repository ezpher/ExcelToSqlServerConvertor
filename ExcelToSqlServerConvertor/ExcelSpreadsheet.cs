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
    class ExcelSpreadsheet
    {
        public string FilePath { get; set; }
        public SpreadsheetDocument SpreadSheetDocument { get; private set; }
        public WorkbookPart WorkbookPart { get; private set; }
        public IEnumerable<Row> SheetDataRows { get; private set; }

        // For filling spreadsheet and use case of 1 worksheet to 1 datatable mapping
        public bool FillSpreadSheet(string filePath)
        {
            FilePath = filePath;

            if (!string.IsNullOrEmpty(FilePath))
            {
                SpreadSheetDocument = SpreadsheetDocument.Open(FilePath, false);

                try
                {
                    //Get sheets
                    WorkbookPart = SpreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = SpreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                    //Get the sheet id i.e. relationship id to find the sheet data (this is just the way the api works)
                    string relationshipId = sheets.First().Id.Value;

                    //Get the sheet data for first sheet
                    WorksheetPart worksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                    // Get the rows for the first sheet
                    SheetDataRows = sheetData.Descendants<Row>();
                }
                catch (OpenXmlPackageException OEx)
                {
                    Console.WriteLine("Error when opening connection to spreadsheet: " + OEx.Message);
                    Console.Read();
                }
                catch (ArgumentException AEx)
                {
                    Console.WriteLine("Error when opening connection to spreadsheet: " + AEx.Message);
                    Console.Read();
                }
                catch (Exception)
                {
                    throw;
                }

                return true;
            }

            return false;
        }

        public static IEnumerable<KeyValuePair<string, Worksheet>> GetNamedWorksheets(WorkbookPart workbookPart)
        {
            IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            return sheets.Select(sheet => new KeyValuePair<string, Worksheet>(sheet.Name, GetWorkSheetFromSheet(workbookPart, sheet)));
        }

        public static Worksheet GetWorkSheetFromSheet(WorkbookPart workbookPart, Sheet sheet)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            return worksheetPart.Worksheet;
        }

        public static IEnumerable<Row> GetWorkSheetRows(Worksheet worksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            return sheetData.Descendants<Row>();
        }
    }

}
