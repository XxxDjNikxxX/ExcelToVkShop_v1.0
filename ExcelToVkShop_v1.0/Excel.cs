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

        public string[,] R_Range(int star_i, int star_x, int end_i, int end_x)
        {
            Range range = (Range)ws.Range[ws.Cells[star_i, star_x], ws.Cells[end_i, end_x]];
            object[,] holder = range.Value;
            string[,] returnstring = new string[end_i-star_i,end_x-star_x];
            for(int p = 1; p <= end_i - star_i; p++)
            {
                for(int q = 1; q < end_x - star_x; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;
        }
    }
}
