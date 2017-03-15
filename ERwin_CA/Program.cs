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
            ConfigFile.TIMESTAMPFOLDER = Timer.GetTimestampFolder(DateTime.Now);
            ConfigFile.FOLDERDESTINATION = Path.Combine(ConfigFile.FOLDERDESTINATION_GENERAL, ConfigFile.TIMESTAMPFOLDER);
            ConfigFile.PERCORSOCOPIEERWINDESTINATION = Path.Combine(ConfigFile.PERCORSOCOPIEERWIN, Timer.GetTimestampFolder(DateTime.Now));
            Logger.PrintLC("** STARTING EXECUTION **");
            ExcelOps Accesso = new ExcelOps();

            int result = MngProcesses.StartProcess();
            switch (result)
            {
                case 0:
                    Logger.PrintLC("Process exited successfully.", 1);
                    break;
                case 1:
                    break;
                case 2:
                    Logger.PrintLC("Exited because no file was found to be processed.", 1, ConfigFile.WARNING);
                    break;
                case 4:
                    Logger.PrintLC("Finished copying process.", 1);
                    break;
                case 5:
                    Logger.PrintLC("Program exited without execution because it wasn't possible to copy remote structure locally.", 1, ConfigFile.ERROR);
                    break;
                case 51:
                    Logger.PrintLC("Program exited without copying files from local to remote structure.", 1, ConfigFile.ERROR);
                    break;
                case 6:
                    Logger.PrintLC("Program stopped abruptly.",1, ConfigFile.ERROR);
                    break;
                case 7:
                    Logger.PrintLC("Templates are missing. Clean exit.", 1, ConfigFile.ERROR);
                    break;
                default:
                    break;
            }
            MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));
            Logger.PrintLC("** FINISHED EXECUTION **");
            Timer.SetSecondTime(DateTime.Now);
            Logger.PrintLC("Execution time: " + Timer.GetTimeLapseFormatted(Timer.GetFirstTime(), Timer.GetSecondTime()) + Environment.NewLine);
        }
    }
}
