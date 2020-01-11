using System.Collections.Generic;

namespace Provider.Models
{
    public class ColumnModel
    {
        public ColumnModel()
        {
            Entrys = new Dictionary<string, decimal>();
        }

        public string                      Description { get; set; }
        public Dictionary<string, decimal> Entrys      { get; set; }

        public override string ToString()
        {
            return $"{Description}";
        }
    }
}