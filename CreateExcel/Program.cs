using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace CreateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //string[] lines = System.IO.File.ReadAllLines(@"C:\Users\frankhuang\Desktop\data.txt");
            //lines.Sort();
           var lines =  File.ReadLines(@"C:\Users\frankhuang\Desktop\datasort2.txt").Select(line => line).Distinct().OrderBy(s => s);
           int i1 = 1;
           int i2 = 1;
           int i3 = 1;
            
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;
            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            
            Worksheet ws = (Worksheet)wb.Worksheets[1];
            Worksheet ws2 ;
            Worksheet ws3 ;
            ws2 = (Worksheet)wb.Worksheets.Add();
            ws3 = (Worksheet)wb.Worksheets.Add();

            //資料結構 string=網址 index
            foreach (string s in lines )
            {
                string[] sOut = s.Split (new char [] {' '});
                if (sOut[0].Contains("Component"))
                {
                    ws.Cells[i1, "A"] = sOut[0];
                    ws.Cells[i1, "B"] = sOut[1];
                    i1++;
                }
                else if (sOut[0].Contains("ascx"))
                {
                    ws2.Cells[i2, "A"] = sOut[0];
                    ws2.Cells[i2, "B"] = sOut[1];
                    i2++;
                }
                else
                {
                    ws3.Cells[i3, "A"] = sOut[0];
                    ws3.Cells[i3, "B"] = sOut[1];
                    i3++;
                }
            }

            wb.SaveAs(@"C:\Users\frankhuang\Desktop\gg.xls");
            wb.Close();
            xlApp.Quit();
        }
    }
}
