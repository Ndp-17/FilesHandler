using FilesHandler.Models.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesHandler.Interfaces
{
    public interface IExcelHandler
    {
        /// <summary>
        /// Read the first sheet or the first one that matches in excel and you can specify which cell to read.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        /// <returns></returns>
        Task<List<Row>> ReadSheet(string sheet = null, string startRange = null, string endRange = null);
        /// <summary>
        /// Read all excel sheets.
        /// </summary>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        /// <returns></returns>
        Task<List<Sheet>> ReadAll(string startRange = null, string endRange = null); 
        /// <summary>
        ///  Read all excel sheets and read by column name
        /// </summary>
        /// <param name="columnNames"></param>
        /// <returns></returns>
        Task<List<Sheet>> ReadColummByName(List<string> columnNames);
        /// <summary>
        /// Read specific excel sheet and read by column name
        /// </summary>
        /// <param name="columnNames"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        Task<List<Sheet>> ReadColummByName(List<string> columnNames, string sheet = null);
        /// <summary>
        ///   Convert to PDF or CSV
        /// </summary>
        /// <param name="destinationFilePath"></param>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        /// <exception cref="Exception"></exception>
        void ConvertTo(string filePath, string filename, string format);
        /// <summary>
        /// Write in specific excel sheet or the first sheet 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cells"></param>
        void Write(string sheet = null, List<Cell> cells = null);
        /// <summary>
        /// Set excel file path
        /// </summary>
        /// <param name="path"></param>
        void SetPath(string path);

    }
}
