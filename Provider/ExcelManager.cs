using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using Provider.Models;

namespace Provider
{
    public class ExcelManager
    {
        private          Excel.Application     xlApp;
        private          Excel.Workbook        xlWorkbook;
        private readonly List<Excel.Worksheet> xlSheets;
        private Excel.Worksheet xlWorksheet;
        private Excel.Range xlRange;

        public ExcelManager()
        {
            var filePath = ConfigurationManager.AppSettings["filePath"];
            xlSheets = new List<Excel.Worksheet>();
            var 
            InitExcelWorkbook(filePath);
        }

        private void InitExcelWorkbook(string path)
        {
            xlApp       = new Excel.Application();
            xlWorkbook  = xlApp.Workbooks.Open(path);
            xlWorksheet = xlWorkbook.Sheets[1] as Excel.Worksheet;
            xlRange = xlWorksheet.UsedRange;


            //foreach (Worksheet sheet in xlWorkbook.Sheets)
            //{
            //    xlSheets.Add(sheet);
            //}
        }

        public ExcelModel GetExcelData()
        {
            return null;
        }

        //private string ReadCell(int i, int j)
        //{
        //    if (xlWorksheet.Cells[i, j].Value2 != null)
        //        return xlWorksheet.Cells[i, j].Value2.ToString();

        //    return $"Cell [{i}, {j}] is empty";
        //}

        public Excel.Workbook GetWorkbook()
        {
            return xlWorkbook;
        }

        public List<Excel.Worksheet> GetSheets()
        {
            return xlSheets;
        }
    }
}