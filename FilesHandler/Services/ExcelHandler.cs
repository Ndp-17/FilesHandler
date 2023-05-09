using FilesHandler.Interfaces;
using FilesHandler.Models.Excel;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace FilesHandler.Services
{
    public class ExcelHandler : IExcelHandler
    {
        private readonly Application excel;
        private string Path { get; set; }
        private bool PathExists { get => File.Exists(Path); }
        public ExcelHandler()
        {
            excel = new Application();
        }
        /// <summary>
        /// Read the first sheet or the first one that matches in excel and you can specify which cell to read.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        /// <returns></returns>
        public async Task<List<Row>> ReadSheet(string sheet = null, string startRange = null, string endRange = null)
        {
            if (!PathExists)
                throw new Exception("Add the file path");

            Workbook workbook = excel.Workbooks.Open(Path);
            if (string.IsNullOrEmpty(sheet))
                sheet = ((Worksheet)workbook.Sheets[1]).Name;

            Worksheet worksheet = workbook.Sheets[sheet];
            Range Range = null;

            if (!string.IsNullOrEmpty(startRange))
                Range = worksheet.Range[startRange, endRange];
            else
                Range = worksheet.UsedRange;


            var Rows = new List<Row>();

            for (int row = 1; row <= Range.Rows.Count; row++)
            {

                var Row = new Row();
                Row.RowNumber = row;
                for (int col = 1; col <= Range.Columns.Count; col++)
                {
                    Row.Columns.Add(new Column { ColumnNumber = col, ColumnName = ((Range)worksheet.Cells[row, col]).Address, Value = Range.Cells[row, col].Value2?.ToString() });
                }
                Rows.Add(Row);
            }
            workbook.Close();
            excel.Quit();
            return Rows;
        }
        /// <summary>
        /// Read all excel sheets.
        /// </summary>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        /// <returns></returns>
        public async Task<List<Sheet>> ReadAll(string startRange = null, string endRange = null)
        {
            if (!PathExists)
                throw new Exception("Add the file path");

            Workbook workbook = excel.Workbooks.Open(Path);
            var Sheets = new List<Sheet>();
            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                string SheetName = ((Worksheet)workbook.Sheets[i + 1]).Name;

                Worksheet worksheet = workbook.Sheets[SheetName];
                Range Range = null;

                if (!string.IsNullOrEmpty(startRange))
                    Range = worksheet.Range[startRange, endRange];
                else
                    Range = worksheet.UsedRange;
                var Rows = new List<Row>();
                for (int row = 1; row <= Range.Rows.Count; row++)
                {

                    var Row = new Row();
                    Row.RowNumber = row;
                    for (int col = 1; col <= Range.Columns.Count; col++)
                    {
                        Row.Columns.Add(new Column { ColumnNumber = col, ColumnName = ((Range)worksheet.Cells[row, col]).Address, Value = Range.Cells[row, col].Value2?.ToString() });
                    }
                    Rows.Add(Row);
                }
                Sheets.Add(new Sheet { SheetName = SheetName, Rows = Rows });
            }
            workbook.Close();
            excel.Quit();
            return Sheets;
        }
        /// <summary>
        ///  Read all excel sheets and read by column name
        /// </summary>
        /// <param name="columnNames"></param>
        /// <returns></returns>
        public async Task<List<Sheet>> ReadColummByName(List<string> columnNames)
        {
            if (!PathExists)
                throw new Exception("Add the file path");

            Workbook workbook = excel.Workbooks.Open(Path);
            var Sheets = new List<Sheet>();
            List<int> columnIndexs = new List<int> { };
            var ColumnNames = columnNames.ToArray();
            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                string SheetName = ((Worksheet)workbook.Sheets[i + 1]).Name;

                Worksheet worksheet = workbook.Sheets[SheetName];
                Range range = worksheet.UsedRange;
                var Rows = new List<Row>();

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    if (columnNames.Contains(range.Cells[1, col].Value2.ToString()))
                    {
                        columnIndexs.Add(col);

                    }
                }
                if (columnIndexs.Any())
                {
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        var Row = new Row();
                        Row.RowNumber = row;
                        int j = 0;
                        foreach (int col in columnIndexs)
                        {
                            Row.Columns.Add(new Column { ColumnNumber = col, ColumnName = ColumnNames[j], Value = range.Cells[row, col].Value2?.ToString() });
                            j++;
                        }
                        Rows.Add(Row);
                    }
                    Sheets.Add(new Sheet { SheetName = SheetName, Rows = Rows });
                }
            }
            workbook.Close();
            excel.Quit();
            return Sheets;

        }
        /// <summary>
        /// Read specific excel sheet and read by column name
        /// </summary>
        /// <param name="columnNames"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public async Task<List<Sheet>> ReadColummByName(List<string> columnNames, string sheet = null)
        {
            if (!PathExists)
                throw new Exception("Add the file path");

            Workbook workbook = excel.Workbooks.Open(Path);
            var Sheets = new List<Sheet>();
            List<int> columnIndexs = new List<int> { };
            var ColumnNames = columnNames.ToArray(); 

            Worksheet worksheet = workbook.Sheets[sheet];
            Range range = worksheet.UsedRange;
            var Rows = new List<Row>();

            for (int col = 1; col <= range.Columns.Count; col++)
            {
                if (columnNames.Contains(range.Cells[1, col].Value2.ToString()))
                {
                    columnIndexs.Add(col);

                }
            }
            if (columnIndexs.Any())
            {
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    var Row = new Row();
                    Row.RowNumber = row;
                    int j = 0;
                    foreach (int col in columnIndexs)
                    {
                        Row.Columns.Add(new Column { ColumnNumber = col, ColumnName = ColumnNames[j], Value = range.Cells[row, col].Value2?.ToString() });
                        j++;
                    }
                    Rows.Add(Row);
                }
                Sheets.Add(new Sheet { SheetName = sheet, Rows = Rows });
            }

            workbook.Close();
            excel.Quit();
            return Sheets;

        }
        /// <summary>
        ///   Convert to PDF or CSV
        /// </summary>
        /// <param name="destinationFilePath"></param>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        /// <exception cref="Exception"></exception>
        public async void ConvertTo(string destinationFilePath, string filename, string format)
        {

            if (!PathExists)
                throw new Exception("Add the file path");

            Workbook workbook = excel.Workbooks.Open(Path);

            switch (format)
            {
                case "PDF":
                    workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, $"{destinationFilePath}\\{filename}.pdf");
                    break;
                case "CSV":
                    workbook.SaveAs($"{destinationFilePath}\\{filename}.csv", XlFileFormat.xlCSV);
                    break;
            }
            workbook.Close(false);
            excel.Quit();
        }
        /// <summary>
        /// Write in specific excel sheet or the first sheet 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cells"></param>
        public async void Write(string sheet = null, List<Cell> cells = null)
        {

            Workbook workbook = excel.Workbooks.Open(Path);
            if (string.IsNullOrEmpty(sheet))
                sheet = ((Worksheet)workbook.Sheets[1]).Name;
            Worksheet worksheet = workbook.ActiveSheet[sheet];

            foreach (var cell in cells)
            {
                worksheet.Cells[cell.RowIndex, cell.ColIndex] = cell.Value;
            }

            workbook.Save();
            workbook.Close();
            excel.Quit();
        }
        /// <summary>
        /// Set excel file path
        /// </summary>
        /// <param name="path"></param>
        public void SetPath(string path)
        {
            Path = $"{path}";
        }
    }
}
