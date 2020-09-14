using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace ExcelToVkShop_v1._0
{
    class Excel
    {
        string path = string.Empty;
        _Application excel = new _Excel.Application();

        Workbook wb;
        Worksheet ws;
        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public string Cells(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value != null)
                return ws.Cells[i, j].Value;
            else
                return string.Empty;


        }
    }
}
