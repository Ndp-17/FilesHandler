using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesHandler.Models.Excel
{
    public class Row
    {
        public Row()
        {
            Columns = new List<Column>();
        }
        public int RowNumber { get; set; }
        public bool NewLine { get; set; }   
        public List<Column> Columns { get; set; }
    }
}
