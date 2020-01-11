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
                var excelWorkbook = new ExcelModel {Name = Constants.WorkbookName};

                do
                {
                    var sheet = new SheetModel {Name = reader.Name};

                    var landColumn         = new ColumnModel();
                    var creatureColumn     = new ColumnModel();
                    var sorceryColumn      = new ColumnModel();
                    var instantColumn      = new ColumnModel();
                    var enchantmentColumn  = new ColumnModel();
                    var artifactColumn     = new ColumnModel();
                    var planeswalkerColumn = new ColumnModel();


                    while (reader.Read())
                    {
                        var landName  = reader.GetValue(0)?.ToString();
                        var landPrice = reader.GetValue(1)?.ToString();
                        if (landColumn.Description == null)
                            landColumn.Description = landName;
                        else
                            if (!string.IsNullOrWhiteSpace(landName) && !string.IsNullOrWhiteSpace(landPrice))
                                landColumn.Entrys.Add(landName, decimal.Parse(landPrice));


                        var creatureName  = reader.GetValue(2)?.ToString();
                        var creaturePrice = reader.GetValue(3)?.ToString();
                        if (creatureColumn.Description == null)
                            creatureColumn.Description = creatureName;
                        else
                            if (!string.IsNullOrWhiteSpace(creatureName) && !string.IsNullOrWhiteSpace(creaturePrice))
                                creatureColumn.Entrys.Add(creatureName, decimal.Parse(creaturePrice));

                        var sorceryName  = reader.GetValue(4)?.ToString();
                        var sorceryPrice = reader.GetValue(5)?.ToString();
                        if (sorceryColumn.Description == null)
                            sorceryColumn.Description = sorceryName;
                        else
                            if (!string.IsNullOrWhiteSpace(sorceryName) && !string.IsNullOrWhiteSpace(sorceryPrice))
                                sorceryColumn.Entrys.Add(sorceryName, decimal.Parse(sorceryPrice));

                        var instantName  = reader.GetValue(6)?.ToString();
                        var instantPrice = reader.GetValue(7)?.ToString();
                        if (instantColumn.Description == null)
                            instantColumn.Description = instantName;
                        else
                            if (!string.IsNullOrWhiteSpace(instantName) && !string.IsNullOrWhiteSpace(instantPrice))
                                instantColumn.Entrys.Add(instantName, decimal.Parse(instantPrice));

                        var enchantmentName  = reader.GetValue(8)?.ToString();
                        var enchantmentPrice = reader.GetValue(9)?.ToString();
                        if (enchantmentColumn.Description == null)
                            enchantmentColumn.Description = enchantmentName;
                        else
                            if (!string.IsNullOrWhiteSpace(enchantmentName) && !string.IsNullOrWhiteSpace(enchantmentPrice))
                                enchantmentColumn.Entrys.Add(enchantmentName, decimal.Parse(enchantmentPrice));

                        var artifactName  = reader.GetValue(10)?.ToString();
                        var artifactPrice = reader.GetValue(11)?.ToString();
                        if (artifactColumn.Description == null)
                            artifactColumn.Description = artifactName;
                        else
                            if(!string.IsNullOrWhiteSpace(artifactName) && !string.IsNullOrWhiteSpace(artifactPrice))
                                artifactColumn.Entrys.Add(artifactName, decimal.Parse(artifactPrice));

                        var planeswalkerName  = reader.GetValue(12)?.ToString();
                        var planeswalkerPrice = reader.GetValue(13)?.ToString();
                        if (planeswalkerColumn.Description == null)
                            planeswalkerColumn.Description = planeswalkerName;
                        else
                            if(!string.IsNullOrWhiteSpace(planeswalkerName) && !string.IsNullOrWhiteSpace(planeswalkerPrice))
                                planeswalkerColumn.Entrys.Add(planeswalkerName, decimal.Parse(planeswalkerPrice));
                    }

                    sheet.Columns.AddRange(new []{ landColumn, creatureColumn, sorceryColumn, instantColumn, enchantmentColumn, artifactColumn, planeswalkerColumn });
                    excelWorkbook.Sheets.Add(sheet);
                } while (reader.NextResult());

                return excelWorkbook;
            }
        }
    }
}