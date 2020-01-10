using System.Configuration;
using System.IO;
using ExcelDataReader;
using Provider.Models;

namespace Provider
{
    public class ExcelManager
    {
        private readonly FileStream excelStream;

        public ExcelManager()
        {
            var filePath = ConfigurationManager.AppSettings["filePath"];
            excelStream = InitExcelWorkbook(filePath);
        }

        private FileStream InitExcelWorkbook(string path)
        {
            var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite);

            return stream;
        }

        public ExcelModel GetExcelData()
        {
            using (var reader = ExcelReaderFactory.CreateReader(excelStream))
            {
                do
                {
                    var sheet = new SheetModel();
                    sheet.Name = reader.Name;


                    while (reader.Read())
                    {
                        var landColumn = new ColumnModel();
                        var landName  = reader.GetValue(0).ToString();
                        var landPrice = reader.GetDecimal(1);
                        if (landColumn.Description == null)
                            landColumn.Description = landName;
                        else
                            landColumn.Entrys.Add(landName, landPrice);

                        var creatureColumn = new ColumnModel();
                        var creatureName  = reader.GetValue(2).ToString();
                        var creaturePrice = reader.GetDecimal(3);
                        if (creatureColumn.Description == null)
                            creatureColumn.Description = creatureName;
                        else
                            creatureColumn.Entrys.Add(creatureName, creaturePrice);

                        var sorceryColumn = new ColumnModel();
                        var sorceryName  = reader.GetValue(4).ToString();
                        var sorceryPrice = reader.GetDecimal(5);
                        if (sorceryColumn.Description == null)
                            sorceryColumn.Description = sorceryName;
                        else
                            sorceryColumn.Entrys.Add(sorceryName, sorceryPrice);

                        var instantColumn = new ColumnModel();
                        var instantName  = reader.GetValue(6).ToString();
                        var instantPrice = reader.GetDecimal(7);
                        if (instantColumn.Description == null)
                            instantColumn.Description = instantName;
                        else
                            instantColumn.Entrys.Add(instantName, instantPrice);

                        var enchantmentColumn = new ColumnModel();
                        var enchantmentName  = reader.GetValue(8).ToString();
                        var enchantmentPrice = reader.GetDecimal(9);
                        if (enchantmentColumn.Description == null)
                            enchantmentColumn.Description = enchantmentName;
                        else
                            enchantmentColumn.Entrys.Add(enchantmentName, enchantmentPrice);

                        var artifactColumn = new ColumnModel();
                        var artifactName  = reader.GetValue(10).ToString();
                        var artifactPrice = reader.GetDecimal(11);
                        if (artifactColumn.Description == null)
                            artifactColumn.Description = artifactName;
                        else
                            artifactColumn.Entrys.Add(artifactName, artifactPrice);

                        var planeswalkerColumn = new ColumnModel();
                        var planeswalkerName  = reader.GetValue(12).ToString();
                        var planeswalkerPrice = reader.GetDecimal(13);
                        if (planeswalkerColumn.Description == null)
                            planeswalkerColumn.Description = planeswalkerName;
                        else
                            planeswalkerColumn.Entrys.Add(planeswalkerName, planeswalkerPrice);
                    }
                } while (reader.NextResult());
            }

            return null;
        }
    }
}