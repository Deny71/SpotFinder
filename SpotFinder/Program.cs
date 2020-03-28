using System;
using System.Data;
using ClosedXML;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
using System.Collections.Generic;
using SpotFinder.Properties;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace SpotFinder
{
    internal class Program
    {
        public static void Main(string[] args)
        {
           System.Console.WriteLine("Working on it...");
           System.Console.WriteLine("\nWhen all operations will be finished window will close automatically");
           System.Console.WriteLine("\nResult will be visible in 'Test1.xlsx' file ");
           Record record = new Record();
           
           ReadSample();
        }

        public static void ReadSample()
        {
            Rand random = new Rand();
            ListaFormatow formaty = new ListaFormatow();
            Excel.Application xlApp = new Excel.Application();
            string currentDirectory = Directory.GetCurrentDirectory();
            string rawpath = System.IO.Path.Combine(currentDirectory,"ExcelData", "raw_Data.xlsx");
            
//dodana ścieżka z folderu kompilowanego raze z programem (ExcelData)
            if (xlApp != null)
            {//@"D:\Program Files\P_Olton\ExcelData\raw_data.xlsx",
                
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(rawpath,
                    0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet) xlWorkbook.Worksheets.get_Item(1);
                Excel.Range xlRange = xlWorkSheet.UsedRange;

                List<Record> records = new List<Record>();
    
                NumberFormatInfo nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ",";
                
                int rw = 0;
                int cw = 0;
                //limiter rekordów
                rw = xlRange.Rows.Count;
                cw = xlRange.Columns.Count;//Int32.Parse(xlRange.Columns.Count.ToString());
                
                
                for (int i = 2; i <= rw; i++)
                {
                    try
                    {
                        records.Add(new Record
                    { 
                        LP = Int32.Parse((xlRange.Cells[i, 1] as Excel.Range).Value2.ToString()),
                        miasto = (xlRange.Cells[i, 2] as Excel.Range).Value2.ToString(),
                          adres = (xlRange.Cells[i ,3] as Excel.Range).Value2.ToString(),
                        firma = (xlRange.Cells[i ,4] as Excel.Range).Value2.ToString(),
                        nr_Tablicy = (xlRange.Cells[i ,5] as Excel.Range).Value2.ToString(),
                        forma = (xlRange.Cells[i ,6] as Excel.Range).Value2.ToString(),
                        latitude_X = Double.Parse((xlRange.Cells[i ,7] as Excel.Range).Value2.ToString().Replace(".",",")),
                        longitude_Y = Double.Parse((xlRange.Cells[i ,8] as Excel.Range).Value2.ToString().Replace(".",",")),
                        nr_tab_Format = (xlRange.Cells[i ,9] as Excel.Range).Value2.ToString(),
                      //  form = buf
                        //potrzebny warunek złapanie przecinków
                    });
                    }
                    catch (System.FormatException)
                    {
                        Console.WriteLine("Masz tu błąd separator: " +i);
                        throw;
                    }
                }
                List<Format52> tymczasowy = formaty.set52();
                
                int lastUsedRow = xlWorkSheet.Cells.Find("*",System.Reflection.Missing.Value, 
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value, 
                    Excel.XlSearchOrder.xlByRows,Excel.XlSearchDirection.xlPrevious, 
                    false,System.Reflection.Missing.Value,System.Reflection.Missing.Value).Row;

                int lastUsedColumn = xlWorkSheet.Cells.Find("*", System.Reflection.Missing.Value, 
                    System.Reflection.Missing.Value,System.Reflection.Missing.Value, 
                    Excel.XlSearchOrder.xlByColumns,Excel.XlSearchDirection.xlPrevious, 
                    false,System.Reflection.Missing.Value,System.Reflection.Missing.Value).Column;
                
                random.losowanie52(records,tymczasowy,lastUsedColumn,lastUsedRow);
                
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkbook);
                    xlWorkbook = null;
                }
                if (xlApp != null)
                {
                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                
            }
            else
            {
                System.Console.WriteLine("Problem z otwarciem pliku");
            }
            
        }
        
    }
}

