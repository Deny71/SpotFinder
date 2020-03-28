using DocumentFormat.OpenXml.Drawing.Diagrams;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace SpotFinder.Properties

{
    public class ListaFormatow
    {
        
        public List<Format52>  set52()
        {
            Excel.Application excelFormat = new Excel.Application();

            if (excelFormat != null)
            {
                //wczytywanie excela
                string currentDirectory = Directory.GetCurrentDirectory();
                string rawpath = System.IO.Path.Combine(currentDirectory,"ExcelData", "podzial.xlsx");
                Excel.Workbook excelWorkbook = excelFormat.Workbooks.Open(
                    rawpath,
                    0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorkSheet = (Excel.Worksheet) excelWorkbook.Worksheets.get_Item(1);
                Excel.Range excelRange = excelWorkSheet.UsedRange;
                List<Format52> format52s = new List<Format52>();
                ////////////////////////
                //testowe sprawdzenie get_range
                //Excel.Range cell1 = excelWorkSheet.get_Range("A1", "A1");
                //Excel.Range cell2 = excelWorkSheet.Rows;
                ///////////////////////////////
            
                int rw = excelRange.Rows.Count;
                int cw = excelRange.Columns.Count;
                
                Excel.Range ran;
                
                double holder2;
                
                for (int j = 2; j<= rw; j++){
                    for (int z = 2; z <= cw; z++)
                    {
                        ran  =
                            (Excel.Range) excelWorkSheet.Range[excelWorkSheet.Cells[j, z], excelWorkSheet.Cells[j, z]];
                        try
                        {
                            holder2 = (double) ran.Value2;
                        }
                        catch
                        {
                            ran.Value2 = 0;
                        }
                    }
                }
                /*
                //aktywne edytowe skoroszyty
                Excel.Application excel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                Excel.Workbooks wbs = excel.Workbooks;
                foreach (Excel.Workbook wb in wbs)
                {
                    Console.WriteLine(wb.Name); // print the name of excel files that are open
                    wb.Save();
                    wb.Close(true);
                }
                excelFormat.ActiveWorkbook.SaveAs(@"D:\Program Files\P_Olton\SpotFinder\podzial2.xlsx", Excel.XlFileFormat.xlWorkbookNormal);
                excelWorkbook.Close();
                excelFormat.Quit();

                
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkSheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelFormat);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                excelWorkbook.Save();
               excelWorkbook.Close(1);
               excelFormat.Quit();
               */
                
                for (int i = 2; i <= rw; i++)
                {
                    try
                    {
                       
                        format52s.Add(item: new Format52
                        {
                            format = (excelRange.Cells[i, 1] as Excel.Range).Value2.ToString(),
                            suma = Int32.Parse((excelRange.Cells[i, 2] as Excel.Range).Value2.ToString()),
                            tv_Sam = Int32.Parse((excelRange.Cells[i, 3] as Excel.Range).Value2.ToString()),
                            tv_LG = Int32.Parse((excelRange.Cells[i, 4] as Excel.Range).Value2.ToString()),
                            tv_Sony = Int32.Parse((excelRange.Cells[i, 5] as Excel.Range).Value2.ToString()),
                             tv_Sha = Int32.Parse((excelRange.Cells[i, 6] as Excel.Range).Value2.ToString()),
                           laptop = Int32.Parse((excelRange.Cells[i, 7] as Excel.Range).Value2.ToString()),
                            tel_Sam = Int32.Parse((excelRange.Cells[i, 8] as Excel.Range).Value2.ToString()),
                            tel_Mot = Int32.Parse((excelRange.Cells[i, 9] as Excel.Range).Value2.ToString()),
                            pra_Sam = Int32.Parse((excelRange.Cells[i, 10] as Excel.Range).Value2.ToString()),
                            pra_Whi = Int32.Parse((excelRange.Cells[i, 11] as Excel.Range).Value2.ToString()),
                            kuc_Ami = Int32.Parse((excelRange.Cells[i, 12] as Excel.Range).Value2.ToString()),
                           lod_Sam = Int32.Parse((excelRange.Cells[i, 13] as Excel.Range).Value2.ToString()), 
                             lod_Bek = Int32.Parse((excelRange.Cells[i, 14] as Excel.Range).Value2.ToString()),
                              susz = Int32.Parse((excelRange.Cells[i, 15] as Excel.Range).Value2.ToString()),
                             oczysz = Int32.Parse((excelRange.Cells[i, 16] as Excel.Range).Value2.ToString()),
                             odk = Int32.Parse((excelRange.Cells[i, 17] as Excel.Range).Value2.ToString()),
                              eksp = Int32.Parse((excelRange.Cells[i, 18] as Excel.Range).Value2.ToString()),
                             szczot = Int32.Parse((excelRange.Cells[i, 19] as Excel.Range).Value2.ToString()),
                             paro = Int32.Parse((excelRange.Cells[i, 20] as Excel.Range).Value2.ToString()),
                          
                        });
                    }
                    catch (System.FormatException)
                    {
                        Console.WriteLine("Masz tu błąd : " +i);
                        throw;
                    }

                } 
                if (excelWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(excelWorkSheet);
                    excelWorkSheet = null;
                }
                if (excelWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(excelWorkbook);
                    excelWorkbook = null;
                }
                if (excelFormat != null)
                {
                    Marshal.FinalReleaseComObject(excelFormat);
                    excelFormat = null;
                }
                
                
             //   excelWorkbook.Close();
             //   excelFormat.Quit();
                return format52s;

            }

            return null;
        }


    }
}