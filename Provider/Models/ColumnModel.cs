using System.Collections.Generic;

namespace Provider.Models
{
    public class ColumnModel
    {
        public ColumnModel()
        {
            Entrys = new Dictionary<string, string>();
        }

        public string                      Description { get; set; }
        public Dictionary<string, string> Entrys      { get; set; }
    }
}