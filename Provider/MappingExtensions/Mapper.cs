//using Microsoft.Office.Interop.Excel;
//using Provider.Models;

//namespace Provider.MappingExtensions
//{
//    public static class Mapper
//    {
//        public static ExcelModel ToModel(this Workbook origin)
//        {
//            var excelModel = new ExcelModel();

//            excelModel.Name = origin.Name;

//            foreach (Worksheet sheet in origin.Sheets)
//            {
//                var sheetModel = new SheetModel();

//                sheetModel.Name = sheet.Name;


                
//                excelModel.Sheets.Add(sheetModel);
//            }

//            return excelModel;
//        }

//        public static SheetModel ToModel(this Worksheet origin)
//        {
//            var sheetModel = new SheetModel();
//            sheetModel.Name = origin.Name;

//            for (int i = 1; i <= 13; i+=2)
//            {
//                var columnModel = new ColumnModel();
//                string cellValue = origin.Cells[1, i].ToString();
//                columnModel.Description = cellValue;

//                for (int j = 2; j < origin.Rows.Count; j++)
//                {
//                    var description = origin.Cells[j, i].ToString();
//                    var price = origin.Cells[j, i + 1].ToString();
//                    columnModel.Entrys.Add(description, price);
//                }

//                sheetModel.Columns.Add(columnModel);
//            }

//            return sheetModel;
//        }

        //private string ReadCell(int i, int j)
        //{
        //    if (xlWorksheet.Cells[i, j].Value2 != null)
        //        return xlWorksheet.Cells[i, j].Value2.ToString();

        //    return $"Cell [{i}, {j}] is empty";
//        //}
//    }
//}
