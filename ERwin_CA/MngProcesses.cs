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
                    Logger.PrintLC("** START PROCESSING FILE: " + file, 2);
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

                        Logger.PrintLC("** START PROCESSING - TABLES parsing from Excel", 2);
                        FileInfo fInfo = new FileInfo(file);
                        List<EntityT> DatiFile = ExcelOps.ReadXFileEntity(fInfo, fileT.TipoDBMS);
                        Logger.PrintLC("** FINISH PROCESSING - TABLES parsing from Excel", 2);

                        Logger.PrintLC("** START PROCESSING - TABLES to ERwin Model", 2);
                        foreach (var dati in DatiFile)
                            connessione.CreateEntity(dati, TemplateFile);
                        Logger.PrintLC("** FINISH PROCESSING - TABLES to ERwin Model", 2);

                        Logger.PrintLC("** START PROCESSING - ATTRIBUTES parsing from Excel", 2);
                        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                        List<AttributeT> AttrFile = null;
                        if (File.Exists(fInfo.FullName))
                        {
                            AttrFile = ExcelOps.ReadXFileAttribute(fInfo, fileT.TipoDBMS);
                        }
                        Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES parsing from Excel", 2);
                        
                        //ATTRIBUTI - PASSAGGIO UNO
                        //Aggiornamento dati struttura
                        Logger.PrintLC("** START PROCESSING - ATTRIBUTES to ERwin model (pass one)", 2);
                        if (!connessione.SetRootObject())
                            continue;
                        if (!connessione.SetRootCollection())
                            continue;
                        //############################
                        foreach (var dati in AttrFile)
                            connessione.CreateAttributePassOne(dati, TemplateFile);
                        Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES to ERwin model (pass one)", 2);

                        Logger.PrintLC("** START PROCESSING - RELATIONS parsing from Excel", 2);
                        List<RelationT> DatiFileRelation = ExcelOps.ReadXFileRelation(fInfo, fileT.TipoDBMS);
                        GlobalRelationStrut globalRelationStrut = Funct.CreaGlobalRelationStrut(DatiFileRelation);
                        Logger.PrintLC("** FINISH PROCESSING - RELATIONS parsing from Excel", 2);

                        Logger.PrintLC("** START PROCESSING - RELATIONS to ERwin Model", 2);
                        object temp = connessione.trID;
                        connessione.CommitAndSave(temp);
                        foreach (var dati in globalRelationStrut.GlobalRelazioni)
                            connessione.CreateRelation(dati, TemplateFile);
                        Logger.PrintLC("** FINISH PROCESSING - RELATIONS to ERwin Model", 2);

                        //ATTRIBUTI - PASSAGGIO DUE
                        //Aggiornamento dati struttura
                        Logger.PrintLC("** START PROCESSING - ATTRIBUTES to ERwin model (pass two)", 2);
                        if (!connessione.SetRootObject())
                            continue;
                        if (!connessione.SetRootCollection())
                            continue;
                        //############################
                        foreach (var dati in AttrFile)
                            connessione.CreateAttributePassTwo(dati, TemplateFile);
                        Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES to ERwin model (pass two)", 2);
                        
                        //Chiusura connessione per il file attuale.
                        connessione.CloseModelConnection();
                        //Eliminazione file originale
                        bool OriginalXLS = false;
                        if (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xlsx")))
                            File.Delete(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                        if (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls")))
                        {
                            OriginalXLS = true;
                            File.Delete(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls"));
                        }
                        //Conversione file di destinazione nel formato XLS
                        if (OriginalXLS == true)
                        {
                            if (File.Exists(fInfo.FullName))
                            {
                                ExcelOps.ConvertXLSXtoXLS(fInfo.FullName);
                                File.Delete(fInfo.FullName);
                            }
                        }
                    }
                    Logger.PrintLC("** FINISH PROCESSING FILE: " + file, 2);
                }
                return 0;
            }
            catch (Exception exp)
            {
                //return exp.HResult;
                return 6;
            }
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
