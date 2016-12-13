using ERwin_CA.T;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    static class MngProcesses
    {
        /// <summary>
        /// MAIN process
        /// </summary>
        /// <returns></returns>
        public static int StartProcess()
        {
            try
            {
                string[] ElencoExcel = DirOps.GetFilesToProcess(ConfigFile.ROOT, "*.xls|.xlsx");
                List<string> gg = FileOps.GetTrueFilesToProcess(ElencoExcel);
                //####################################
                //Ciclo MAIN
                foreach (var file in gg)
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
                        //Apertura connessione per il file attuale
                        ConnMng connessione = new ConnMng();
                        if (!connessione.openModelConnection(destERFile))
                            continue;
                        //Aggiornamento della struttura dati per il file attuale
                        if (!connessione.SetRootObject())
                            continue;
                        if (!connessione.SetRootCollection())
                            continue;

                        FileInfo fInfo = new FileInfo(file);
                        List<EntityT> DatiFile = ExcelOps.ReadXFileEntity(fInfo, fileT.TipoDBMS);
                        foreach (var dati in DatiFile)
                            connessione.CreateEntity(dati, TemplateFile);
                        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                        List<AttributeT> AttrFile = null;
                        if (File.Exists(fInfo.FullName))
                        {
                            AttrFile = ExcelOps.ReadXFileAttribute(fInfo, fileT.TipoDBMS);
                        }
                        //ATTRIBUTI - PASSAGGIO UNO
                        //Aggiornamento dati struttura
                        if (!connessione.SetRootObject())
                            continue;
                        if (!connessione.SetRootCollection())
                            continue;
                        //############################
                        foreach (var dati in AttrFile)
                            connessione.CreateAttributePassOne(dati, TemplateFile);

                        ////ATTRIBUTI - PASSAGGIO DUE
                        ////Aggiornamento dati struttura
                        //if (!connessione.SetRootObject())
                        //    continue;
                        //if (!connessione.SetRootCollection())
                        //    continue;
                        ////############################
                        //foreach (var dati in AttrFile)
                        //    connessione.CreateAttributePassTwo(dati, TemplateFile);

                        //Chiusura connessione per il file attuale.
                        connessione.CloseModelConnection();
                    }
                }
                return 0;
            }
            catch (Exception exp)
            {
                //return exp.HResult;
                return 6;
            }
            return 6;
        }

        public static Process[] ProcList(string procName)
        {
            Process[] processes = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(procName))
                {
                    processes = Process.GetProcessesByName(procName);
                    return processes;
                }
            }
            catch (Exception exp)
            {

            }
            return processes;
        }

        public static void KillAllOf(Process[] processes)
        {
            try
            {
                foreach (Process proc in processes)
                {
                    if(proc.MainWindowTitle == "")
                    {
                        proc.Kill();
                        proc.WaitForExit();
                    }
                }
            }
            catch (System.NullReferenceException)
            {

            }
        }

    }
}
