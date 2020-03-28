using System;
using System.Diagnostics;
using System.Collections;

// TO JEST PIEKNA SPRAWA

namespace SpotFinder.Properties
{
    public class Killprocess
    {
        Hashtable myHashtable;
        public void btnExport_Click()
        {
            // get process ids before running the excel codes
            CheckExcellProcesses();

            // export to excel
            ExportDataToExcel();

            // kill the right process after export completed
            KillExcel();
        }

        private void ExportDataToExcel()
        {
            // your export process is here...
        }

        private void CheckExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach ( Process ExcelProcess in AllProcesses) {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        private void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
    
            // check to kill the right process
            foreach ( Process ExcelProcess in AllProcesses) {
                if (myHashtable.ContainsKey(ExcelProcess.Id) == true)
                    ExcelProcess.Kill();
            }
            Console.Beep(2000,1000);
            AllProcesses = null;
        }

    }
}