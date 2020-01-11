
using System.Collections.Generic;

namespace Provider.Models
{
    public class SheetModel
    {
        public SheetModel()
        {
            Columns = new List<ColumnModel>();
        }

        public string Name { get; set; }
        public List<ColumnModel> Columns { get; set; }

        public override string ToString()
        {
            return $"{Name}";
        }
    }
}