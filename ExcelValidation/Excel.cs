using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelValidation
{
    public class Excel
    {
        string path = "";
        Application excel = new Application();
        private Workbook wb;
        private Worksheet ws;
        HashSet<string> emailHashSet = new HashSet<string>();
        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCellsFromColumn(int i, int j)
        {
            //i++;
            //j++;

            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            else
            {
                return " ";
            }
        }

        public void WriteInCell(int i, int j, string value)
        {

            ws.Cells[i, j].Value = value;
        }

        public void SaveWorkBook()
        {
            wb.Save();
            wb.Close();
        }

        public void CloseFile()
        {
            wb.Close();
        }
    }
}
