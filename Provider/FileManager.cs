using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace Provider
{
    public class FileManager
    {
        private          Application     xlApp;
        private          Workbook        xlWorkbook;
        private          Worksheet       xlWorksheet;
        private readonly List<Worksheet> xlSheets;

        public FileManager()
        {
            var filePath = ConfigurationManager.AppSettings["filePath"];
            xlSheets = new List<Worksheet>();

            InitExcelWorkbook(filePath);
        }

        private void InitExcelWorkbook(string path)
        {
            xlApp       = new Application();
            xlWorkbook  = xlApp.Workbooks.Open(path);

            foreach (Worksheet sheet in xlWorkbook.Sheets)
            {
                xlSheets.Add(sheet);
            }
        }

        public void GetExcelData()
        {
            var col = 1;
            var row = 1;

            //var asd = ReadCell(row, col);
        }

        //private string ReadCell(int i, int j)
        //{
        //    if (xlWorksheet.Cells[i, j].Value2 != null)
        //        return xlWorksheet.Cells[i, j].Value2.ToString();

        //    return $"Cell [{i}, {j}] is empty";
        //}

        public Workbook GetWorkbook()
        {
            return xlWorkbook;
        }

        public List<Worksheet> GetSheets()
        {
            return xlSheets;
        }
    }
}