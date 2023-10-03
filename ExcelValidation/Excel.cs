using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelValidation
{
    public class Excel : IDisposable
    {
        private Application excel;
        private Workbook wb;
        private Worksheet ws;

        public Excel(string path, int sheet)
        {
            try
            {
                excel = new Application();
                wb = excel.Workbooks.Open(path);
                ws = (Worksheet)wb.Worksheets[sheet];
            }
            catch (Exception ex)
            {
                throw new Exception($"Error opening Excel file: {ex.Message}");
            }
        }

        public string ReadCellsFromColumn(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2.ToString();
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
            try
            {
                wb.Save();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving Excel workbook: {ex.Message}");
            }
            finally
            {
                Dispose();
            }
        }

        public void Dispose()
        {
            // Clean up resources
            if (wb != null)
            {
                wb.Close();
                Marshal.ReleaseComObject(wb);
                wb = null;
            }

            if (excel != null)
            {
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                excel = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}