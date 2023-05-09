using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesHandler.Models.Excel
{
    public class Cell
    {
        public int RowIndex { get; set; }

        public int ColIndex { get; set; }

        public string Value { get; set; }
    }
}
