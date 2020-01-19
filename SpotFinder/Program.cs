using System;
using System.Data;
using ClosedXML;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpotFinder
{
    internal class Program
    {
        public static void Main(string[] args)
        {
           System.Console.WriteLine("Witam");
           Record record = new Record();
           ReadSample();
        }

        public static void ReadSample()
        {
            
            Excel.Application xlApp = new Excel.Application();
            
            if (xlApp != null)
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Program Files\P_Olton\SpotFinder\raw_data.xlsx",
                    0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet) xlWorkbook.Worksheets.get_Item(1);
                Excel.Range xlRange = xlWorkSheet.UsedRange;
               
                List<Record> records = new List<Record>();
                
                
                int rowCount = 1; //xlRange.Rows.Column;
                int colCount = 10; //xlRange.Columns.Count;
                string str = (string)(xlRange.Cells[rowCount, colCount] as Excel.Range).Value2;
                System.Console.WriteLine(str);
               
                
                records.Add(new Record{LP=212, miasto = "test"});
              /*  records.Add(new Record
                {
                    LP = 21
                    //(int)(xlRange.Cells[rowCount, colCount] as Excel.Range).Value2,
                  
                    
                    miasto = (string)(xlRange.Cells[2, 2] as Excel.Range).Value2,
                    adres = (string)(xlRange.Cells[2, 3] as Excel.Range).Value2,
                    firma = (string)(xlRange.Cells[2, 4] as Excel.Range).Value2, 
                    nr_Tabeli = (string)(xlRange.Cells[2, 5] as Excel.Range).Value2,
                    forma = (string)(xlRange.Cells[2, 6] as Excel.Range).Value2,
                    latitude_X = (int)(xlRange.Cells[2, 7] as Excel.Range).Value2,
                    longitude_Y = (int)(xlRange.Cells[2, 8] as Excel.Range).Value2,
                    nr_tab_Format = (string)(xlRange.Cells[2, 9] as Excel.Range).Value2
                
                });
                */
                System.Console.Writecord.LP;
                
                string xas;
                for (int i = 1; i <= 10; i++)
                {
                    xas = (string) (xlRange.Cells[1, i] as Excel.Range).Value2;
                    System.Console.WriteLine(xas);
                    
                }
                xlWorkbook.Close();
                xlApp.Quit();
            }
        }
        
    }
}