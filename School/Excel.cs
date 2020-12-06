using System;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace School
{

    class Excel
    {
        String path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(String path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        public bool cellIsNotEmpty(int row, int column)
        {
            var value = ws.Cells[row, column].Value2;

            if (value != null && !String.Equals(value, ""))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public int cellsCount (Excel excel,int row, int column)
        {
            row = 1;
            column = 11;
            int cellsNumber = 0;
            while (true)
            {
                if(excel.readCell(row, column) != null)
                    if (String.Equals(excel.readCell(row, column).Trim(), "الختم"))
                        return cellsNumber;
                cellsNumber++;
                row++;
            }
        }
        public String readCell(int row, int column)
        {
            var value = ws.Cells[row, column].Value2;
                if (value is double)
                {
                    return value.ToString("0.######");
                }
                return value;
        }
        public void closeFile()
        {
            if(wb != null && excel != null) { 
            wb.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }
    }
}
