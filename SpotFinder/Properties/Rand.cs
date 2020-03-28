using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using Excel = Microsoft.Office.Interop.Excel;
using SpotFinder.Properties;
using System.Timers;

namespace SpotFinder.Properties
{
    public class Rand
    {

        public void losowanie52(List<Record> formatka, List<Format52> ilosci,int kolumny,int wiersze)
        {
            
            //var excelFile = Path.GetFullPath(@"D:\Program Files\P_Olton\SpotFinder\Test1.xlsx");
            var excelFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "FOLDERdamiana", "FOLDERwewnatrzFolderuDamiana");
            Excel.Application excel = new Excel.ApplicationClass();
            excel.DisplayAlerts = false;
            Excel.Workbook wb = excel.Workbooks.Add(Missing.Value);
            Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            ws.Name = "Test";
            ws.Cells[1, 1] = "LP";
            ws.Cells[1, 2] = "MIASTO";
            ws.Cells[1, 3] = "Adres";
            ws.Cells[1, 4] = "FIRMA";
            ws.Cells[1, 5] = "Nr tablicy";
            ws.Cells[1, 6] = "Forma";
            ws.Cells[1, 7] = "X";
            ws.Cells[1, 8] = "Y";
            ws.Cells[1, 9] = "Nr tablicy z formatem";
            ws.Cells[1, 10] = "kategoria";
            ws.Cells[1, 11] = "Typ reklamy";
            ws.Cells[1, 12] = "test";

            
                //trzeba bedzie pokombinowac z napisanem metody która zwroci wartosc z bezposrednio z klasy
                //
            
           //ws.Cells[2, 1];
           int zmienna = 1;
               var cos = formatka.GetRange(2, 1);
               var test = formatka[1].adres;
               

               int j = 2;
            for (int i = 0; i <= wiersze-2; i++)
            {
                ws.Cells[j, 1] = formatka[i].LP;
                ws.Cells[j, 2] = formatka[i].miasto;
                ws.Cells[j, 3] = formatka[i].adres;
                ws.Cells[j, 4] = formatka[i].miasto;
                ws.Cells[j, 5] = formatka[i].nr_Tablicy;
                ws.Cells[j, 6] = formatka[i].forma;
                ws.Cells[j, 7] = formatka[i].latitude_X;
                ws.Cells[j, 8] = formatka[i].longitude_Y;
                ws.Cells[j, 9] = formatka[i].nr_tab_Format;
                ws.Cells[j, 10] = formatka[i].kategoria;
                ws.Cells[j, 11] = "Typ reklamy";
                ws.Cells[j, 12] = "test";
                j++;
            }

            int licznik = 0;
            int los;
            string typRek;
            j = 2;
            Random rnd = new Random(); //potrzebne do wylosowania typu reklamy
            
            ///////////////////////////////////////////////////////////////////////////////////////////////////////
            for (int i = 0; i <= wiersze-2; i++)
            {
                if (formatka[i].forma == "5x2")//losowanie dla 5x2
              {
                  for (;;)
                  {
                      licznik++;
                      typRek = Randv2.losowacz52(los = rnd.Next(1, 19), ilosci, formatka, i);
                      if (typRek == "again")
                      {
                      }
                      else
                      {
                          break;
                      }

                      if (FormatCheck.check52(ilosci) == "empty")
                      {
                          typRek = "brak_pozycji";
                          break;
                      }

                      if (licznik > 2000)
                      {
                          System.Console.WriteLine("KRECE SIE TU NA 52   " +licznik);
                          
                      }
                  }

                  ws.Cells[j, 11] = typRek;
                  j++;
              }
                else if (formatka[i].forma == "6x3")// losowanie dla 6x3
                {
                    for (;;)
                    {
                        typRek = Randv3.losowacz63(los = rnd.Next(1, 19), ilosci, formatka, i); 
                        if (typRek == "again")
                        {
                        }
                        else
                        {
                            break;
                        }
                        if (FormatCheck.check63(ilosci) == "empty")
                        {
                            typRek = "brak_pozycji";
                            break;
                        }
                        
                        if (licznik > 2000)
                        {
                            System.Console.WriteLine("KRECE SIE TU NA 63    " +licznik);
                        }
                    }

                    ws.Cells[j, 11] = typRek;
                    j++;
                    
                }
                else if (formatka[i].forma == "12x3" || formatka[i].forma == "12x4")// losowanie dla 6x3
                {
                    for (;;)
                    {
                        typRek = Randv4.losowacz1234(los = rnd.Next(1, 19), ilosci, formatka, i); 
                        if (typRek == "again")
                        {
                        }
                        else
                        {
                            break;
                        }
                        if (FormatCheck.check1234(ilosci) == "empty")
                        {
                            typRek = "brak_pozycji";
                            break;
                        }
                        
                        if (licznik > 2000)
                        {
                            System.Console.WriteLine("KRECE SIE TU NA 1234    " +licznik);
                        }
                    }

                    ws.Cells[j, 11] = typRek;
                    j++;
                    
                }
                else
                {
                    j++;
                }

            }

         /*   foreach (var a in ilosci)
            {
                Console.WriteLine(a.tv_Sam+a.tv_LG+a.tv_Sony+a.tv_Sha+a.laptop+a.tel_Sam+a.tel_Mot+a.pra_Sam+a.pra_Whi+a.kuc_Ami+a.lod_Sam+a.lod_Bek+a.susz+a.oczysz+a.odk+a.eksp+a.szczot+a.paro);
                
            }
*/
            //ustawiony jest tryb odczytu (Patrz ponizej)'
            wb.SaveAs(excelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            
            Killprocess kill = new Killprocess();
         kill.btnExport_Click();
         
        }
        
    }
    
}