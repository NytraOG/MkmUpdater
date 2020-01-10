using System.Collections.Generic;

namespace Provider.Models
{
    public class ExcelModel
    {
        public ExcelModel()
        {
            Sheets = new List<SheetModel>();
        }

        public string           Name   { get; set; }
        public List<SheetModel> Sheets { get; set; }
    }
}