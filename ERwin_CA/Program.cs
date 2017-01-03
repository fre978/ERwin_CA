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
            ConfigFile.RefreshLogLevel();
            Logger.PrintLC("** STARTING EXECUTION **");
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

            MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));

            Logger.PrintLC("** FINISHED EXECUTION **");
            Timer.SetSecondTime(DateTime.Now);
            Logger.PrintLC("Execution time: " + Timer.GetTimeLapseFormatted(Timer.GetFirstTime(), Timer.GetSecondTime()) + Environment.NewLine);
        }
    }
}
