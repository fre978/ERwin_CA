using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ERwin_CA.T;
using System.IO;

namespace ERwin_CA
{
    class Funct
    {

        public static string GetTemplate(FileT file)
        {
            string TemplateFile = string.Empty;
            switch (file.TipoDBMS)
            {
                case ConfigFile.DB2_NAME:
                    TemplateFile = ConfigFile.ERWIN_TEMPLATE_DB2;
                    break;
                case ConfigFile.ORACLE:
                    TemplateFile = ConfigFile.ERWIN_TEMPLATE_ORACLE;
                    break;
                case ConfigFile.SQLSERVER:
                    TemplateFile = ConfigFile.ERWIN_TEMPLATE_SQLSERVER;
                    break;
                default:
                    string fileName = file.SSA + "_" + file.Acronimo + "_" + file.NomeModello + "_" + file.TipoDBMS + "_" + file.Estensione;
                    Logger.PrintLC(fileName + ": DB descriptor is invalid. Skipping it.");
                    TemplateFile = null;
                    break;
            }
            return TemplateFile;
        }

        public static List<string> RemoteGet(List<string> FileDaElaborare)
        {
            DirectoryInfo local = new DirectoryInfo(ConfigFile.LOCAL_DIR_FULL);
            if (local.Exists)
                DirOps.TraverseDirectory(local);

            Logger.PrintLC("## Starting elaboration of remote/local file structure.", 2, ConfigFile.INFO);
            try
            {
                string[] listSQL = DirOps.GetFilesToProcess(ConfigFile.ROOT, "*.sql");
                List<string> listSQLdaCopiare = listSQL.ToList();
                DirOps.Copy(ConfigFile.ROOT, ConfigFile.LOCAL_DIR_FULL, FileDaElaborare);
                DirOps.Copy(ConfigFile.ROOT, ConfigFile.LOCAL_DIR_FULL, listSQLdaCopiare);
            }
            catch
            {
                Logger.PrintLC("Error while copying remote structure to local temporary structure.", 2, ConfigFile.ERROR);
                return new List<string>();
            }
            // SEZIONE AGGIORNAMENTO PATH CARTELLE DI SISTEMA
            ConfigFile.TEMP_REMOTE_ROOT = ConfigFile.ROOT;
            ConfigFile.FOLDERDESTINATION_GENERAL = Path.Combine(ConfigFile.LOCAL_DIR_FULL, ConfigFile.DEST_FOLD_NAME);
            ConfigFile.FOLDERDESTINATION = Path.Combine(ConfigFile.FOLDERDESTINATION_GENERAL, ConfigFile.TIMESTAMPFOLDER);
            //###############################################
            if (ConfigFile.LOCAL_DIR_FULL.Substring(ConfigFile.LOCAL_DIR_FULL.Length - 1) != @"\")
            {
                ConfigFile.ROOT = ConfigFile.LOCAL_DIR_FULL + @"\";
            }
            else
            {
                ConfigFile.ROOT = ConfigFile.LOCAL_DIR_FULL;
            }

            //FileDaElaborareRemoto = FileDaElaborare;
            string[] elencoLocale = DirOps.GetFilesToProcess(ConfigFile.ROOT, "*.xls|.xlsx");
            List<string> FileDaElaborareCompletoLocale = FileOps.GetTrueFilesToProcess(elencoLocale);
            FileDaElaborare = Parser.ParseListOfFileNames(FileDaElaborareCompletoLocale);
            Logger.PrintLC("## Finished elaboration of remote/local file structure.", 2, ConfigFile.INFO);
            return FileDaElaborare;
        }

        /// <summary>
        /// Copy the LOCAL structure of files and directories TO the original remote location,
        /// eliminating superflous Excel files.
        /// </summary>
        /// <param name="FileElaborati">List of files originally intended to be elaborated</param>
        /// <returns>State of the elaboration</returns>
        public static bool RemoteSet(List<string> FileElaborati)
        {
            Logger.PrintLC("## Copying back files from local to remote.", 2, ConfigFile.INFO);

            //ConfigFile.ROOT = ConfigFile.TEMP_REMOTE_ROOT;
            try
            {
                DirOps.Copy(ConfigFile.LOCAL_DIR_FULL, ConfigFile.TEMP_REMOTE_ROOT);
            }
            catch
            {
                Logger.PrintLC("Error while copying local temporary structure to remote structure.", 2, ConfigFile.ERROR);
                return false;
            }
            DirectoryInfo local = new DirectoryInfo(ConfigFile.LOCAL_DIR_FULL);
            if (local != null)
            {
                DirOps.TraverseDirectory(local);
            }

            foreach (var elem in FileElaborati)
            {
                string file = elem.Replace(ConfigFile.LOCAL_DIR_FULL + @"\", ConfigFile.TEMP_REMOTE_ROOT);
                FileInfo fileI = new FileInfo(file);
                FileInfo fileIU = null;
                switch (fileI.Extension)
                {
                    case ".xls":
                        fileIU = new FileInfo(Path.Combine(fileI.DirectoryName, Path.GetFileNameWithoutExtension(fileI.FullName) + ".XLS"));
                        //fileI = fileIU;
                        break;
                    case ".xlsx":
                        fileIU = new FileInfo(Path.Combine(fileI.DirectoryName, Path.GetFileNameWithoutExtension(fileI.FullName) + ".XLSX"));
                        //fileI = fileIU;
                        break;
                }
                if (fileI.Exists || fileIU.Exists)
                {
                    if (fileIU.Exists)
                    {
                        //fileI = fileIU;
                    }
                    string name = fileI.Name;
                    string dir = fileI.DirectoryName;
                    string estensione = string.Empty;
                    string textNameOK = string.Empty;
                    string textNameKO = string.Empty;
                    FileInfo fileTestoOK = null;
                    FileInfo fileTestoKO = null;
                    try
                    {

                        if (fileI.Extension == ".xls")
                        {
                            estensione = ".xls";
                            textNameOK = name.Replace(".xls", "_OK.txt");
                            textNameKO = name.Replace(".xls", "_KO.txt");
                        }
                        if (fileI.Extension == ".xlsx")
                        {
                            estensione = ".xlsx";
                            textNameOK = name.Replace(".xlsx", "_OK.txt");
                            textNameKO = name.Replace(".xlsx", "_KO.txt");
                        }
                        if (fileI.Extension == ".XLS")
                        {
                            estensione = ".XLS";
                            textNameOK = name.Replace(".XLS", "_OK.txt");
                            textNameKO = name.Replace(".XLS", "_KO.txt");
                        }
                        if (fileI.Extension == ".XLSX")
                        {
                            estensione = ".XLSX";
                            textNameOK = name.Replace(".XLS", "_OK.txt");
                            textNameKO = name.Replace(".XLSX", "_KO.txt");
                        }
                        fileTestoOK = new FileInfo(Path.Combine(dir, textNameOK));
                        fileTestoKO = new FileInfo(Path.Combine(dir, textNameKO));
                    }
                    catch (Exception exp)
                    {
                        Logger.PrintLC("Errore 9: " + exp.Message);
                        continue;
                    }
                    if (fileTestoKO.Exists)
                    {
                        if (!ConfigFile.DEST_FOLD_UNIQUE)
                        {
                            try
                            {
                                string dirDestinationKO = Funct.GetFolderDestination(fileI.FullName, estensione);
                                FileInfo fileCopiare = new FileInfo(dirDestinationKO);
                                try
                                {
                                    fileI.IsReadOnly = false;
                                    //fileI.MoveTo(fileCopiare.DirectoryName);
                                    fileI.MoveTo(dirDestinationKO);
                                    //fileI.Delete();
                                }
                                catch(Exception exp)
                                {
                                    try
                                    {
                                        if(exp.HResult == -2147024713)
                                            fileI.Delete();
                                        //if (exp.Message == "Impossibile creare un file, se il file esiste già.\r\n")
                                        //    fileI.Delete();
                                    }
                                    catch { }
                                }


                                try
                                {
                                    fileIU.IsReadOnly = false;
                                    fileIU.MoveTo(fileCopiare.DirectoryName + @"\");
                                    fileIU.Delete();
                                }
                                catch (Exception exp)
                                {
                                    try
                                    {
                                        //fileIU.Delete();
                                    }
                                    catch { }
                                }


                                //if (!fileCopiare.Exists)
                                //{
                                //    fileI.MoveTo(fileCopiare.DirectoryName);
                                //}
                                //else
                                //{
                                //    fileCopiare.Delete();
                                //    fileI.MoveTo(fileCopiare.DirectoryName);
                                //}
                            }
                            catch (Exception exp)
                            {
                                Logger.PrintLC("Errore 10: " + exp.Message);
                            }
                        }
                        else
                        {
                            try
                            {
                                fileI.MoveTo(Path.Combine(ConfigFile.ROOT, ConfigFile.DEST_FOLD_NAME));
                            }
                            catch (Exception exp)
                            {
                                Logger.PrintLC("Errore 11: " + exp.Message);
                            }

                        }
                    }
                    if (fileTestoOK.Exists)
                    {
                        try
                        {
                            fileI.Delete();
                        }
                        catch (Exception exp)
                        {
                            Logger.PrintLC("Errore 12: " + exp.Message);
                        }
                        try
                        {
                            string fileXLSX = Path.GetFileNameWithoutExtension(fileI.FullName);
                            FileInfo fileIxls = new FileInfo( Path.Combine(fileI.DirectoryName, fileXLSX + ".xlsx"));
                            FileInfo fileIXLSX = new FileInfo(Path.Combine(fileI.DirectoryName, fileXLSX + ".XLSX"));
                            try
                            {
                                fileIxls.Delete();
                            }
                            catch { }
                            try
                            {
                                fileIXLSX.Delete();
                            }
                            catch { }
                        }
                        catch
                        {

                        }
                    }
                }
            }
            return true;
        }


        public static void PrintList(List<string> list)
        {
            Logger.PrintLC("List of elements: ",2);
            foreach (string inList in list)
            {
                Logger.PrintLC(inList, 3);
            }
        }

        public static List<string> DetermineElaborated(List<string> completi, List<string> elaborati)
        {
            List<string> restituzione = new List<string>();
            if (completi == null)
                return null;
            if (elaborati == null)
            {
                foreach(string elemento in completi)
                {
                    restituzione.Add(elemento + ": NON processato.");
                }
                return restituzione;
            }
            foreach(string elemento in completi)
            {
                if (elaborati.Contains(elemento))
                    restituzione.Add(elemento + ": PROCESSATO.");
                else
                    restituzione.Add(elemento + ": NON processato.");
            }
            return restituzione;
        }

        public static string GetFolderDestination(string FileInElaborazione,string Estensione)
        {
            string mystring = string.Empty;
            mystring = Path.GetFullPath(FileInElaborazione);
            mystring = mystring.Replace(Path.GetFileName(FileInElaborazione), "");
            mystring = Path.Combine(mystring, ConfigFile.DEST_FOLD_NAME, ConfigFile.TIMESTAMPFOLDER, Path.GetFileNameWithoutExtension(FileInElaborazione) + Estensione);
            return mystring;
        }
        public static string GetFolderDestination2(string FileInElaborazione, string FullName)
        {
            string mystring = string.Empty;
            mystring = Path.GetFullPath(FileInElaborazione);
            mystring = mystring.Replace(Path.GetFileName(FileInElaborazione), "");
            mystring = Path.Combine(mystring, ConfigFile.DEST_FOLD_NAME, ConfigFile.TIMESTAMPFOLDER, FullName);
            return mystring;
        }

        public static string RemoveWhitespace(string input)
        {
            return new string(input.ToCharArray()
                .Where(c => !Char.IsWhiteSpace(c))
                .ToArray());
        }

        public static bool ParseDataType(string value, string databaseType, bool OnlyFormal = false)
        {
            if (value.ToUpper().Contains("NUMBER"))
            {
                //Logger.PrintLC("TROVATO");
            }
            string[] actualDB = null;
            if (!ConfigFile.DBS.Contains(databaseType))
                return false;
            else
            {
                if (OnlyFormal)
                {
                    databaseType = databaseType + "_FOR";
                }
                switch (databaseType)
                {
                    case ConfigFile.DB2_NAME:
                        actualDB = ConfigFile.DATATYPE_DB2;
                        break;
                    case ConfigFile.ORACLE:
                        actualDB = ConfigFile.DATATYPE_ORACLE;
                        break;
                    case ConfigFile.SQLSERVER:
                        actualDB = ConfigFile.DATATYPE_SQLSERVER;
                        break;
                    case ConfigFile.DB2_NAME + "_FOR":
                        actualDB = ConfigFile.DATATYPE_DB2_FOR;
                        break;
                    case ConfigFile.ORACLE + "_FOR":
                        actualDB = ConfigFile.DATATYPE_ORACLE_FOR;
                        break;
                    case ConfigFile.SQLSERVER + "_FOR":
                        actualDB = ConfigFile.DATATYPE_SQLSERVER_FOR;
                        break;
                    default:
                        break;
                }
            }
            int oUt1;
            int oUt2;
            if (value.Contains(","))
            {
                try
                {
                    string[] a = value.Split('(');
                    string primo = a[0];
                    string[] b = a[1].Split(',');
                    string secondo = b[0];
                    string[] c = (b[1]).Split(')');
                    string terzo = c[0];
                    if (int.TryParse(secondo, out oUt1) && int.TryParse(terzo, out oUt2) && actualDB.Contains(primo.ToLower()))
                        return true;
                    else
                        return false;
                }
                catch(Exception exp)
                {
                    return false;
                }
            }
            if (value.Contains("("))
            {
                try
                {
                    string[] a = value.Split('(');
                    string primo = a[0];
                    string[] b = a[1].Split(')');
                    string secondo = b[0];
                    if (int.TryParse(secondo, out oUt1) && (actualDB.Contains(primo.ToLower()) || actualDB.Contains(primo.ToLower() + "()")))
                        return true;
                    else
                        return false;
                }
                catch(Exception exp)
                {
                    return false;
                }
            }
            else
            {
                if (actualDB.Contains(value.ToLower()))
                    return true;
                else
                    return false;
            }
        }

        public static bool ParseFlag(string value, string FlagType)
        {
            string[] actualDB = null;
            switch (FlagType)
            {
                case "YES":
                    actualDB = ConfigFile.Yes;
                    break;
                case "NO":
                    actualDB = ConfigFile.No;
                    break;
                default:
                    break;
            }
            
            if (actualDB.Contains(value.ToUpper().Trim()))
                return true;
            else
                return false;
            
        }

        public static bool Stats(decimal current, decimal maximum, string message, string fileCorrect)
        {
            try
            { 
                decimal percent = (current / maximum) * 100;
                message = decimal.Round(percent,3) + "% (" + current + " su " + maximum + ") " + message;
                //Logger.PrintF(fileCorrect, message, true, ConfigFile.INFO);
                Logger.PrintLC(message, 2, ConfigFile.INFO);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static GlobalRelationStrut CreaGlobalRelationStrut(List<RelationT> relazioni)
        {
            // crea struttura
            GlobalRelationStrut GStrut = new GlobalRelationStrut();
            GStrut = CreaGlobalRelationStrutGrezze(relazioni);
            GStrut = CleanGlobalRelationStrut(GStrut);
            return GStrut;
        }

        public static GlobalRelationStrut CreaGlobalRelationStrutGrezze(List<RelationT> relazioni)
        {
            GlobalRelationStrut Gstrut = new GlobalRelationStrut();
            if (relazioni == null)
                return Gstrut = null;
            try
            { 
                foreach (var rel in relazioni)
                {
                    //IEnumerable<RelationStrut> ExistRelationStrut = Gstrut.GlobalRelazioni.Where(x => x.ID == rel.IdentificativoRelazione);
                    bool trovato = false;
                    foreach (var Rstrut in Gstrut.GlobalRelazioni)
                        if (Rstrut.ID == rel.IdentificativoRelazione)
                        {
                            trovato = true;
                            Rstrut.Relazioni.Add(rel);
                            continue;
                        }
                    if (trovato == false)
                    {
                        RelationStrut RStrut = new RelationStrut();
                        RStrut.ID = rel.IdentificativoRelazione;
                        RStrut.Relazioni.Add(rel);
                        Gstrut.GlobalRelazioni.Add(RStrut);
                    }
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Error filtering relations list. Error:" + exp.Message, 3, ConfigFile.ERROR);
                return Gstrut = null;
            }

            return Gstrut;
        }

        public static GlobalRelationStrut CleanGlobalRelationStrut(GlobalRelationStrut GStrut)
        {
            try
            {
                List<RelationStrut> errorRelationStrut = new List<RelationStrut>();
                //verifica tutte le strutture
                foreach (RelationStrut RStrut in GStrut.GlobalRelazioni)
                {
                    if (RStrut.Relazioni.Count != 1)
                    {
                        //verifica singola struttura
                        string tabellapadreverifica = null;
                        string tabellafigliaverifica = null;
                        int? cardinalitaverifica = null;
                        int? identificativaverifica = null;
                        bool? tiporelazioneverifica = null;
                        List<string> campopadreverifica = new List<string>();
                        List<string> campofiglioverifica = new List<string>();

                        int contatore = 0;
                        bool errore = false;


                        foreach (RelationT R in RStrut.Relazioni)
                        {

                            if (contatore == 0)
                            {
                                tabellapadreverifica = R.TabellaPadre;
                                tabellafigliaverifica = R.TabellaFiglia;
                                cardinalitaverifica = R.Cardinalita;
                                identificativaverifica = R.Identificativa;
                                tiporelazioneverifica = R.TipoRelazione;
                                campopadreverifica.Add(R.CampoPadre);
                                campofiglioverifica.Add(R.CampoFiglio);

                            }
                            else
                            {
                                if (tabellapadreverifica != R.TabellaPadre
                                    || tabellafigliaverifica != R.TabellaFiglia
                                    || cardinalitaverifica != R.Cardinalita
                                    || identificativaverifica != R.Identificativa
                                    || tiporelazioneverifica != R.TipoRelazione)
                                {
                                    errore = true;
                                    //PUNTO IN CUI ANDARE A SCRIVERE SULL'EXCEL ALLA RIGA APPROPRIATA
                                    R.History = "Relazione ignorata: ID " + RStrut.ID + " presenta valori diversi per uno o più dei seguenti campi: tabella padre, tabella figlia, cardinalità, identificativa e tipo relazione";
                                    Logger.PrintLC("Relazione ignorata: ID " + RStrut.ID + " presenta valori diversi per uno o più dei seguenti campi: tabella padre, tabella figlia, cardinalità, identificativa e tipo relazione", 3, ConfigFile.ERROR);
                                    continue;
                                }


                                if (campopadreverifica.Contains(R.CampoPadre) || campofiglioverifica.Contains(R.CampoFiglio))
                                {
                                    errore = true;
                                    R.History = "Relazione ignorata: ID " + RStrut.ID + " campo padre e/o campo figlio duplicati all'interno della relazione";
                                    Logger.PrintLC("Relazione ignorata: ID " + RStrut.ID + " campo padre e/o campo figlio duplicati all'interno della relazione", 3, ConfigFile.ERROR);
                                    continue;
                                }
                                else
                                {
                                    campopadreverifica.Add(R.CampoPadre);
                                    campofiglioverifica.Add(R.CampoFiglio);
                                }

                            }
                            contatore += 1;

                        }
                        if (errore == true)
                            errorRelationStrut.Add(RStrut);
                    }
                }

                foreach (var errore in errorRelationStrut)
                {
                    RelationStrut mystrut = GStrut.GlobalRelazioni.Find(x => x == errore);
                    if (mystrut != null)
                    {
                        foreach (RelationT R in mystrut.Relazioni)
                        {
                            RelationT comodo = mystrut.Relazioni.Find(x => !string.IsNullOrEmpty(x.History));
                            if (comodo != null)
                            {
                                if (string.IsNullOrEmpty(R.History))
                                {
                                    R.History = comodo.History;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                //
            }
            return GStrut;
        }

        // OLD VALIDATION PROCESS
        // CODE 66
        //public static bool ValidateDatabaseName(string val)
        //{
        //    if (!string.IsNullOrWhiteSpace(val))
        //    {
        //        char[] charVal = val.ToCharArray();
        //        int len = 0;
        //        foreach(char single in charVal)
        //        {
        //            if (char.IsLetterOrDigit(single))
        //            {
        //                len++;
        //                continue;
        //            }
        //            else
        //            {
        //                if (single == '-' || single == ' ' || single == '_')
        //                {
        //                    len++;
        //                    continue;
        //                }
        //            }
        //        }
        //        if (val.Length == len)
        //        {
        //            return true;
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}
    }
}
