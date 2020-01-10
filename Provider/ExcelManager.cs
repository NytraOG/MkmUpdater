using System.Configuration;
using System.IO;
using ExcelDataReader;
using Provider.Models;

namespace Provider
{
    public class ExcelManager
    {
        public ExcelManager()
        {
            var filePath = ConfigurationManager.AppSettings["filePath"];
            InitExcelWorkbook(filePath);
        }

        private void InitExcelWorkbook(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        for (var i = 0; i <= 14; i += 2)
                        {
                            while (reader.Read())
                            {
                                var name  = reader.GetValue(i);
                                var price = reader.GetValue(i + 1);

                                var xxx = name + ": " + price + "€";
                            }
                        }




                    } while (reader.NextResult());

                    //var result = reader.AsDataSet();

                    //var x = result.Tables[0]
                    //              .Rows[1];
                }
            }
        }

        public ExcelModel GetExcelData()
        {
            return null;
        }
    }
}