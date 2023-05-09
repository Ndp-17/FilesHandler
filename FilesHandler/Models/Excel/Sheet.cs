using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesHandler.Models.Excel
{
    public class Sheet
    {
        public Sheet()
        {
            Rows = new List<Row>();
        }
        public string SheetName { get; set; }
        public List<Row> Rows { get; set; }
    }
}
