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
            Logger.PrintLC("AVVIO ESECUZIONE", 1);
            ExcelOps Accesso = new ExcelOps();
            
            //string nomeFile = @"C:\ERWIN\CODICE\Extra\" + fileDaAprire.Name.ToString();
            //bool testBool = Accesso.ConvertXLStoXLSX(nomeFile);
            //testBool = ExcelOps.FileValidation(nomeFile);
            string[] ElencoExcel = DirOps.GetFilesToProcess(ConfigFile.FILETEST, "*.xls");

            //####################################
            //Ciclo MAIN
            foreach (var file in ElencoExcel)
            {
                string TemplateFile = null;
                if (ExcelOps.FileValidation(file))
                {
                    FileT fileT = Parser.ParseFileName(file);
                    string destERFile = null;
                    if (fileT != null)
                    {
                        switch (fileT.TipoDBMS)
                        {
                            case ConfigFile.DB2_NAME:
                                TemplateFile = ConfigFile.ERWIN_TEMPLATE_DB2;
                                break;
                            case ConfigFile.ORACLE:
                                TemplateFile = ConfigFile.ERWIN_TEMPLATE_ORACLE;
                                break;
                            default:
                                TemplateFile = ConfigFile.ERWIN_TEMPLATE_DB2;
                                break;
                        }
                        FileInfo origin = new FileInfo(file);
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        destERFile = Path.Combine(ConfigFile.FOLDERDESTINATION, fileName + Path.GetExtension(TemplateFile));
                        if (!FileOps.CopyFile(TemplateFile, destERFile))
                            continue;
                    }
                    else
                        continue;

                    ConnMng connessione = new ConnMng();
                    if (!connessione.openModelConnection(destERFile))
                        continue;

                    connessione.SetRootObject();
                    connessione.SetRootCollection();

                    FileInfo fInfo = new FileInfo(file);
                    List<EntityT> DatiFile = ExcelOps.ReadXFile(fInfo);
                    foreach(var dati in DatiFile)
                        connessione.CreateEntity(dati, TemplateFile);

                    Logger.PrintLC("File " + file + " not valid for processing.");

                    //Chiusura delle Sessioni aperte in elaborazione
                    //SCAPI.Sessions cc = connessione.scERwin.Sessions;
                    //foreach (SCAPI.Session ses in cc)
                    //    ses.Close();
                    connessione.CloseModelConnection();
                    //connessione = null;

                    //Chiusura di tutti i file EXCEL aperti in elaborazione
                }
            }
            //####################################

        FINE_PROGRAMMA:
            MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));

            Logger.PrintL("TERMINE ESECUZIONE");
            Timer.SetSecondTime(DateTime.Now);
            Logger.PrintL("Tempo esecuzione: " + Timer.GetTimeLapseFormatted(Timer.GetFirstTime(), Timer.GetSecondTime()) + Environment.NewLine);
        }
    }
}
