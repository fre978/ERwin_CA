using ERwin_CA.T;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
//using Interop.ERXMLLib;
//using Interop.SCAPI;



namespace ERwin_CA
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.Initialize(ConfigFile.LOG_FILE);
            ConfigFile.FOLDERDESTINATION = Path.Combine(ConfigFile.FOLDERDESTINATION_GENERAL, Timer.GetTimestampFolder(DateTime.Now));
            Logger.PrintLC("AVVIO ESECUZIONE");
            ExcelOps Accesso = new ExcelOps();

            int result = MngProcesses.StartProcess();
            switch (result)
            {
                case 0:
                    break;
                case 1:
                    break;
                case 6:
                    Logger.PrintLC("Program stopped abruptly with this error: ");
                    break;
                default:
                    break;
            }

            //####################################

        FINE_PROGRAMMA:
            MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));

            Logger.PrintLC("TERMINE ESECUZIONE");
            Timer.SetSecondTime(DateTime.Now);
            Logger.PrintLC("Tempo esecuzione: " + Timer.GetTimeLapseFormatted(Timer.GetFirstTime(), Timer.GetSecondTime()) + Environment.NewLine);
        }
    }
}
