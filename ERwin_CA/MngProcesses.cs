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
                if (ConfigFile.RefreshAll() == true)
                    Logger.PrintLC("!! Some error occured while parsing the config file. Standard values will be used instead.",1, ConfigFile.WARNING);
                List<string> FileElaborati = new List<string>();
                List<ElaboratiT> Elaborati = new List<ElaboratiT>();
                List<string> FileElaboratiSQL = new List<string>();
                string[] ElencoExcel = DirOps.GetFilesToProcess(ConfigFile.ROOT, "*.xls|.xlsx");
                List<string> FileDaElaborare = FileOps.GetTrueFilesToProcess(ElencoExcel);
                decimal current = 0;
                decimal maximum = 0;
                string message = string.Empty;
                                
                //####################################
                //Ciclo MAIN
                foreach (var file in FileDaElaborare)
                {
                    #region ProcessingFileExcel
                    Logger.PrintLC("** START PROCESSING EXCEL FILE: " + file, 2);
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
                                case ConfigFile.SQLSERVER:
                                    TemplateFile = ConfigFile.ERWIN_TEMPLATE_SQLSERVER;
                                    break;
                                default:
                                    TemplateFile = ConfigFile.ERWIN_TEMPLATE_DB2;
                                    break;
                            }
                            FileInfo origin = new FileInfo(file);
                            string fileName = Path.GetFileNameWithoutExtension(file);
                            //string[] fileComponents;
                            //fileComponents = fileName.Split(ConfigFile.DELIMITER_NAME_FILE);
                            //fileName = fileComponents[2];
                            if (ConfigFile.DEST_FOLD_UNIQUE)
                            {
                                destERFile = Path.Combine(ConfigFile.FOLDERDESTINATION, fileName + Path.GetExtension(TemplateFile));
                            }
                            else
                            {
                                destERFile = Funct.GetFolderDestination(file, Path.GetExtension(TemplateFile));
                            } 
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

                        #region EsameTabelleExcel
                        Logger.PrintLC("** START PROCESSING - TABLES parsing from Excel", 2);
                        FileInfo fInfo = new FileInfo(file);
                        List<EntityT> DatiFile = ExcelOps.ReadXFileEntity(fInfo, fileT.TipoDBMS);
                        Logger.PrintLC("** FINISH PROCESSING - TABLES parsing from Excel", 2);
                        #endregion

                        #region EsameTabelleErwin
                        Logger.PrintLC("** START PROCESSING - TABLES to ERwin Model", 2);

                        connessione.CreateModel(fileT.NomeModello);

                        int EntitaCreate = 0;
                        foreach (var dati in DatiFile)
                        { 
                            SCAPI.ModelObject Entita = connessione.CreateEntity(dati, TemplateFile);
                            if (Entita != null)
                            {
                                EntitaCreate += 1;
                            }

                            //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                            if (!string.IsNullOrEmpty(dati.History))
                            { 
                                int col = ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1;
                                Logger.PrintLC("Updating excel file for error on entity creation for the table '" + dati.TableName + "' in erwin. Error: " + dati.History, 3);
                                //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                if (ConfigFile.DEST_FOLD_UNIQUE)
                                {
                                    fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                }
                                else
                                {
                                    fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                }

                                if (File.Exists(fInfo.FullName))
                                {
                                    ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Row, col, 1, dati.History, ConfigFile.TABELLE);
                                }
                            }
                        }

                        Logger.PrintLC("** FINISH PROCESSING - TABLES to ERwin Model", 2);
                        #endregion

                        #region StatsTabelleCreate
                        //Al termine dell'elaborazione delle entità scrivo nel file di log la statistica delle tabelle create 
                        string fileError = Path.Combine(new FileInfo(file).DirectoryName, Path.GetFileNameWithoutExtension(file) + "_KO.txt");
                        string fileCorrect = Path.Combine(new FileInfo(file).DirectoryName, Path.GetFileNameWithoutExtension(file) + "_OK.txt");
                        if (EntitaCreate != 0)
                        {
                            //statistica tabelle create
                            current = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            maximum = DatiFile.Count;
                            message = "tabelle correttamente create";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            //statistica tabelle senza descrizione
                            current = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History) && string.IsNullOrEmpty(x.TableDescr)).Count();
                            maximum = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "tabelle senza descrizione sul totale delle tabelle create";
                            Funct.Stats(current, maximum, message, fileCorrect);
                        }
                        else
                        {
                            // nel caso non abbia creato alcuna tabella
                            if (File.Exists(fileError))
                            {
                                //rimuovo un eventuale file di errore
                                FileOps.RemoveAttributes(fileError);
                                File.Delete(fileError);

                            }
                            if (File.Exists(fileCorrect))
                            {
                                //rinomino il file corretto in file di errore
                                FileOps.CopyFile(fileCorrect, fileError);
                                FileOps.RemoveAttributes(fileCorrect);
                                File.Delete(fileCorrect);
                                //scrivo la statistica
                                Logger.PrintF(fileError, EntitaCreate + " tabelle create", true, ConfigFile.INFO);
                                Logger.PrintLC(EntitaCreate + " entity created", 2, ConfigFile.INFO);
                            }
                        }
                        #endregion

                        #region EsameAttributiExcel
                        Logger.PrintLC("** START PROCESSING - ATTRIBUTES parsing from Excel", 2);
                        //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                        if (ConfigFile.DEST_FOLD_UNIQUE)
                        {
                            fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                        }
                        else
                        {
                            fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                        }
                        List<AttributeT> AttrFile = null;
                        if (File.Exists(fInfo.FullName))
                        {
                            AttrFile = ExcelOps.ReadXFileAttribute(fInfo, fileT.TipoDBMS);
                        }

                        Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES parsing from Excel", 2);
                        #endregion

                        //se non è stata creata nessuna entità salto questo step
                        if (EntitaCreate != 0)
                        {
                            #region EsameAttributiErwin1
                            //ATTRIBUTI - PASSAGGIO UNO
                            //Aggiornamento dati struttura
                            Logger.PrintLC("** START PROCESSING - ATTRIBUTES to ERwin model (pass one)", 2);
                            if (!connessione.SetRootObject())
                                continue;
                            if (!connessione.SetRootCollection())
                                continue;
                            //############################
                            foreach (var dati in AttrFile)
                            {
                                connessione.CreateAttributePassOne(dati, TemplateFile);

                                //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                                if (!string.IsNullOrEmpty(dati.History))
                                {
                                    int col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                                    Logger.PrintLC("Updating excel file for error on attributes creation (pass one) for the field '" + dati.NomeCampoLegacy + "' in erwin. Error: " + dati.History, 3);
                                    //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                    if (ConfigFile.DEST_FOLD_UNIQUE)
                                    {
                                        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                    }
                                    else
                                    {
                                        fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                    }
                                    if (File.Exists(fInfo.FullName))
                                    {
                                        ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Row, col, 1, dati.History, ConfigFile.ATTRIBUTI);
                                    }
                                }
                            }


                            if (ConfigFile.CREACOPIEERWIN == "true")
                            {
                                try
                                {
                                    string ORIGIN = connessione.fileInfoERwin.FullName;
                                    if (File.Exists(ORIGIN))
                                    {
                                        string PercorsoCopieErwin = ConfigFile.PERCORSOCOPIEERWINDESTINATION;
                                        string DESTINATION = Path.Combine(PercorsoCopieErwin, Path.GetFileNameWithoutExtension(connessione.fileInfoERwin.FullName) + "_attrib" + connessione.fileInfoERwin.Extension);
                                        FileOps.CopyFile(ORIGIN, DESTINATION);
                                        Logger.PrintLC("Created copy of erwin file after ATTRIBUTES to ERwin model (pass one)", 4, ConfigFile.INFO);
                                    }
                                    else
                                    {
                                        Logger.PrintLC("Impossibile to create a copy of erwin file after ATTRIBUTES to ERwin model (pass one)", 4, ConfigFile.ERROR);
                                    }
                                }
                                catch (Exception exc)
                                {
                                    Logger.PrintLC("Impossibile to create a copy of erwin file after ATTRIBUTES to ERwin model (pass one)", 4, ConfigFile.ERROR);
                                }
                            }

                            Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES to ERwin model (pass one)", 2);
                            #endregion
                        }

                        #region EsameRelazioniExcel
                        Logger.PrintLC("** START PROCESSING - RELATIONS parsing from Excel", 2);
                        List<RelationT> DatiFileRelation = ExcelOps.ReadXFileRelation(fInfo, fileT.TipoDBMS);
                        GlobalRelationStrut globalRelationStrut = Funct.CreaGlobalRelationStrut(DatiFileRelation);
                        Logger.PrintLC("** FINISH PROCESSING - RELATIONS parsing from Excel", 2);
                        #endregion

                        //se non è stata creata nessuna entità salto questo step
                        if (EntitaCreate != 0)
                        {
                            #region EsameRelazioniErwin
                            Logger.PrintLC("** START PROCESSING - RELATIONS to ERwin Model", 2);
                            //object temp = connessione.trID;
                            //connessione.CommitAndSave(temp);
                            foreach (var dati in globalRelationStrut.GlobalRelazioni)
                            {
                                connessione.CreateRelation(dati, TemplateFile);

                                //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                                foreach (var dato in dati.Relazioni)
                                {
                                    if (!string.IsNullOrEmpty(dato.History))
                                    {
                                        int col = ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1;
                                        Logger.PrintLC("Updating excel file for error on relation creation for the field '" + dato.IdentificativoRelazione + "' in erwin. Error: " + dato.History, 3);
                                        //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                        if (ConfigFile.DEST_FOLD_UNIQUE)
                                        {
                                            fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                        }
                                        else
                                        {
                                            fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                        }
                                        if (File.Exists(fInfo.FullName))
                                        {
                                            ExcelOps.XLSXWriteErrorInCell(fInfo, dato.Row, col, 1, dato.History, ConfigFile.RELAZIONI);
                                        }
                                    }
                                }
                            }
                            if (ConfigFile.CREACOPIEERWIN == "true")
                            {
                                try
                                {
                                    string ORIGIN = connessione.fileInfoERwin.FullName;
                                    if (File.Exists(ORIGIN))
                                    {
                                        string PercorsoCopieErwin = ConfigFile.PERCORSOCOPIEERWINDESTINATION;
                                        string DESTINATION = Path.Combine(PercorsoCopieErwin, Path.GetFileNameWithoutExtension(connessione.fileInfoERwin.FullName) + "_rel" + connessione.fileInfoERwin.Extension);
                                        FileOps.CopyFile(ORIGIN, DESTINATION);
                                        Logger.PrintLC("Created copy of erwin file after RELATIONS to ERwin Model", 4, ConfigFile.INFO);
                                    }
                                    else
                                    {
                                        Logger.PrintLC("Impossibile to create a copy of erwin file after RELATIONS to ERwin Model", 4, ConfigFile.ERROR);
                                    }
                                }
                                catch (Exception exc)
                                {
                                    Logger.PrintLC("Impossibile to create a copy of erwin file after RELATIONS to ERwin Model", 4, ConfigFile.ERROR);
                                }
                            }


                            Logger.PrintLC("** FINISH PROCESSING - RELATIONS to ERwin Model", 2);
                            #endregion
                        }

                        //se non è stata creata nessuna entità salto questo step
                        if (EntitaCreate != 0)
                        {
                            #region EsameAttributiErwin2
                            //ATTRIBUTI - PASSAGGIO DUE
                            //Aggiornamento dati struttura
                            Logger.PrintLC("** START PROCESSING - ATTRIBUTES to ERwin model (pass two)", 2);
                            if (!connessione.SetRootObject())
                                continue;
                            if (!connessione.SetRootCollection())
                                continue;
                            //############################
                            foreach (var dati in AttrFile)
                            {
                                connessione.CreateAttributePassTwo(dati, TemplateFile);

                                //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                                if (!string.IsNullOrEmpty(dati.History) && (dati.Step == 2))
                                {
                                    int col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                                    Logger.PrintLC("Updating excel file for error on attributes creation (pass two) for the field '" + dati.NomeCampoLegacy + "' in erwin. Error: " + dati.History, 3);
                                    //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                    if (ConfigFile.DEST_FOLD_UNIQUE)
                                    {
                                        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                    }
                                    else
                                    {
                                        fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                    }
                                    if (File.Exists(fInfo.FullName))
                                    {
                                        ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Row, col, 1, dati.History, ConfigFile.ATTRIBUTI);
                                    }
                                }

                            }
                            Logger.PrintLC("** FINISH PROCESSING - ATTRIBUTES to ERwin model (pass two)", 2);
                            #endregion
                        }

                        if (EntitaCreate != 0)
                        {
                            #region StatsAttributi
                            #region TabelleConPrimaryKey
                            //statistica tabelle senza primary key
                            List<string> SenzaChiave = new List<string>();
                            List<string> ConChiave = new List<string>();
                            foreach (AttributeT attributo in AttrFile)
                            {
                                //se è chiave
                                if (attributo.Chiave == 0 && string.IsNullOrEmpty(attributo.History))
                                {
                                    //se già non esistente
                                    if (!(ConChiave.Exists(x => x == attributo.NomeTabellaLegacy)))
                                    {
                                        //rimuovo dai senza chiave
                                        if (SenzaChiave.Exists(x => x == attributo.NomeTabellaLegacy))
                                        {
                                            SenzaChiave.Remove(attributo.NomeTabellaLegacy);
                                        }
                                        //aggiungo
                                        ConChiave.Add(attributo.NomeTabellaLegacy);
                                    }
                                }
                                else
                                {
                                    //se non esiste fra quelli con le chiavi e non esiste nemmeno fra quelli senza
                                    if (!(ConChiave.Exists(x => x == attributo.NomeTabellaLegacy)) && !(SenzaChiave.Exists(x => x == attributo.NomeTabellaLegacy)))
                                    {
                                        //aggiungo
                                        SenzaChiave.Add(attributo.NomeTabellaLegacy);
                                    }
                                }
                            }
                            current = SenzaChiave.Count;
                            maximum = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "tabelle prive di chiave primaria rispetto al totale delle tabelle create";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            #region AttributiConalmenoUnErrore
                            //statistica tabelle senza descrizione
                            current = AttrFile.FindAll(x => !(string.IsNullOrEmpty(x.History))).Count;
                            maximum = AttrFile.Count;
                            message = "attributi errati rispetto al totale degli attributi elaborati";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            #region AttributiSenzaDescrizione
                            //statistica tabelle senza descrizione
                            current = AttrFile.FindAll(x => string.IsNullOrEmpty(x.History) && string.IsNullOrEmpty(x.DefinizioneCampo)).Count();
                            maximum = AttrFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "attributi privi di descrizione rispetto al totale degli attributi creati";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            #endregion
                            #region StatsRelazioni
                            List<string> TabelleRelazionate = new List<string>();
                            List<string> RelazioniOK = new List<string>();
                            foreach (RelationStrut rel in globalRelationStrut.GlobalRelazioni)
                            {
                                #region TabelleIsola
                                RelationT myrel = rel.Relazioni.Find(x => string.IsNullOrEmpty(x.History));
                                if (myrel != null)
                                {
                                    if (!TabelleRelazionate.Exists(x => x == myrel.TabellaPadre))
                                    {
                                        TabelleRelazionate.Add(myrel.TabellaPadre);
                                    }
                                    if (!TabelleRelazionate.Exists(x => x == myrel.TabellaFiglia))
                                    {
                                        TabelleRelazionate.Add(myrel.TabellaFiglia);
                                    }
                                }
                                #endregion
                                #region RelazioniOK/KO
                                //myrel = rel.Relazioni.Find(x => string.IsNullOrEmpty(x.History));
                                if (myrel != null)
                                {
                                    if (!RelazioniOK.Exists(x => x == myrel.IdentificativoRelazione))
                                    {
                                        RelazioniOK.Add(myrel.IdentificativoRelazione);
                                    }
                                }
                                #endregion
                            }
                            
                            current = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count - TabelleRelazionate.Count;
                            maximum = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "tabelle prive di relazioni sul totale delle relazioni create";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            current = (globalRelationStrut.GlobalRelazioni.Count) - TabelleRelazionate.Count;
                            maximum = (globalRelationStrut.GlobalRelazioni.Count);
                            message = "relazioni errate sul totale delle elaborazioni elaborate";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            
                        }

                        //Chiusura connessione per il file attuale.
                        connessione.CloseModelConnection();
                        //Eliminazione file originale
                        bool OriginalXLS = false;
                        string FileElaborato = null;
                        if (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xlsx")))
                        {
                            FileElaborato = Path.Combine(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                            if ((EntitaCreate != 0) || (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls")))) 
                                File.Delete(FileElaborato);
                        }
                        if (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls")))
                        {
                            OriginalXLS = true;
                            FileElaborato = Path.Combine(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls"));
                            //if (EntitaCreate != 0) //se non ha creato entyty non lo cancello perche KO
                            //{
                            File.Delete(FileElaborato);
                            //}
                        }
                        //Conversione file di destinazione nel formato XLS
                        //if (EntitaCreate != 0)
                        //{
                        if (OriginalXLS == true)
                        {
                            if (File.Exists(fInfo.FullName))
                            {
                                if (ExcelOps.ConvertXLSXtoXLS(fInfo.FullName))
                                {
                                    File.Delete(fInfo.FullName);
                                }
                            }
                        }
                        //}
                        if (EntitaCreate == 0)
                        {
                            //if (File.Exists(fInfo.FullName))
                            //{
                            //    File.Delete(fInfo.FullName);
                            //}
                            string erw = Path.GetFileNameWithoutExtension(fInfo.FullName) + ".erwin";
                            erw = fInfo.FullName.Replace(fInfo.Name, erw);
                            if (File.Exists(erw))
                            {
                                File.Delete(erw);
                            }
                        }

                        
                        FileElaborati.Add(FileElaborato);
                        ElaboratiT Elaborato = new ElaboratiT(fileElaborato: "", entityElaborate: new List<EntityT>(), attributiElaborati: new List<AttributeT>(), relazioniElaborate: new GlobalRelationStrut());
                        Elaborato.FileElaborato = FileElaborato;
                        Elaborato.EntityElaborate = DatiFile;
                        Elaborato.AttributiElaborati = AttrFile;
                        Elaborato.RelazioniElaborate = globalRelationStrut;
                        Elaborati.Add(Elaborato);
                        
                    }
                    //Fine processi
                    Logger.PrintLC("** FINISH PROCESSING EXCEL FILE: " + file, 2);
                    #endregion
                }
                #region SummaryFiles
                //Stampa elenco completo file presi in considerazione
                Logger.PrintLC("\n## SUMMARY EXCEL FILES:");
                List<string> ListaCompleta = Funct.DetermineElaborated(FileDaElaborare, FileElaborati);
                foreach (string elemento in ListaCompleta)
                {
                    Logger.PrintLC(elemento, 2);
                }
                #endregion

                foreach (var Elaborato in Elaborati)
                {
                    string FileElaborato = Elaborato.FileElaborato;
                    List<EntityT> EntityElaborate = Elaborato.EntityElaborate;
                    List<AttributeT> AttributiElaborati = Elaborato.AttributiElaborati;
                    GlobalRelationStrut RelazioniElaborate = Elaborato.RelazioniElaborate;

                    #region DocExcelControlli
                    List<string> DocExcelControlli = new List<string>();
                    List<EntityT> EntityBFD = Elaborato.EntityElaborate.FindAll(x => x.FlagBFD == "S" && string.IsNullOrEmpty(x.History));
                    List<RelationStrut> LRelazioniBFD = RelazioniElaborate.GlobalRelazioni;
                    foreach (EntityT E in EntityBFD)
                    {
                        List<AttributeT> AttributiBFD = AttributiElaborati.FindAll(x => x.NomeTabellaLegacy.ToUpper() == E.TableName.ToUpper() && string.IsNullOrEmpty(x.History));
                        foreach (AttributeT A in AttributiBFD)
                        {
                            int? type = A.Chiave;
                            int? null_option_type = A.MandatoryFlag;
                            string phisical_data_type = A.DataType;
                            string NomeStrutturaInformativa = string.Empty;
                            string NomeCampo = string.Empty;
                            string CodLocaleControllo = string.Empty;
                            string RuoloCampo = string.Empty;

                            if (type == 0)
                            {
                                NomeStrutturaInformativa = E.TableName.ToUpper();
                                NomeCampo = A.NomeCampoLegacy.ToUpper();
                                CodLocaleControllo = "DUP";
                                RuoloCampo = "OggettoControllo";
                                DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                            }
                            if (A.MandatoryFlag == 1)
                            {
                                NomeStrutturaInformativa = E.TableName.ToUpper();
                                NomeCampo = A.NomeCampoLegacy.ToUpper();
                                CodLocaleControllo = "NUL";
                                RuoloCampo = "OggettoControllo";
                                DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                            }
                            if (Funct.ParseDataType(phisical_data_type, "DB2", true))
                            {
                                NomeStrutturaInformativa = E.TableName.ToUpper();
                                NomeCampo = A.NomeCampoLegacy.ToUpper();
                                CodLocaleControllo = "FOR";
                                RuoloCampo = "OggettoControllo";
                                DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                            }
                            foreach (RelationStrut SRelazioniBFD in LRelazioniBFD)
                            {
                                List<RelationT> Relazioni = SRelazioniBFD.Relazioni.FindAll(x => x.CampoFiglio == A.NomeCampoLegacy && x.TabellaFiglia == A.NomeTabellaLegacy && string.IsNullOrEmpty(x.History));
                                foreach (RelationT Relazione in Relazioni)
                                {
                                    NomeStrutturaInformativa = E.TableName.ToUpper();
                                    NomeCampo = A.NomeCampoLegacy.ToUpper();
                                    CodLocaleControllo = "DRI";
                                    RuoloCampo = "OggettoControllo";
                                    DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);

                                    NomeStrutturaInformativa = Relazione.TabellaPadre.ToUpper();
                                    NomeCampo = Relazione.CampoPadre.ToUpper();
                                    CodLocaleControllo = "DRI";
                                    RuoloCampo = "CampoConfronto";
                                    DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                }
                            }
                        }
                    }
                    string FileDocControlli = Path.GetFileNameWithoutExtension(FileElaborato) + "_ControlliCampi.xlsx";
                    //FileDocControlli = Path.Combine(ConfigFile.FOLDERDESTINATION, FileDocControlli);
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        FileDocControlli = Path.Combine(ConfigFile.FOLDERDESTINATION, FileDocControlli);
                    }
                    else
                    {
                        FileDocControlli = Funct.GetFolderDestination2(FileElaborato, new FileInfo(FileDocControlli).Name);
                    }
                    ExcelOps.WriteDocExcelControlli(new FileInfo(FileDocControlli), DocExcelControlli);

                    #endregion

                    #region ProcessingFileSQL
                    Logger.PrintLC("** INIZIO ELABORAZIONE DDL: " + FileElaborato, 2);

                    
                    #region ElaborazioneCoppieXlsDdl
                    string fullNameSQL = Path.GetFileNameWithoutExtension(FileElaborato) + ".sql";
                    string FileDaElaborareSQL = Path.GetFullPath(FileElaborato);
                    FileDaElaborareSQL = FileDaElaborareSQL.Replace(Path.GetFileName(FileElaborato), fullNameSQL);
                    string FileDifferenze = Path.GetFileNameWithoutExtension(FileElaborato) + "_diffvsddl.xlsx";
                    FileDifferenze = Path.Combine(ConfigFile.FOLDERDESTINATION,FileDifferenze);
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        FileDifferenze = Path.Combine(ConfigFile.FOLDERDESTINATION, FileDifferenze);
                    }
                    else
                    {
                        FileDifferenze = Funct.GetFolderDestination2(FileDaElaborareSQL, new FileInfo(FileDifferenze).Name);
                    }

                    //per i file correttamente elaborati nel modulo precedente cerchiamo se ci sono i corrispettivi file ddl
                    if (File.Exists(FileDaElaborareSQL))
                    {
                        Logger.PrintLC("Un corrispondente file DDL esiste per il file " + FileElaborato, 3,ConfigFile.INFO);

                        //se il file esiste inizio a leggere il contenuto e a collezionarne le informazioni
                        #region EsameTabelleSQL
                        Dictionary<string, List<String>> CompareResults = new Dictionary<string, List<string>>();

                        Logger.PrintLC("** INIZIO ELABORAZIONE - TABELLE parsing da DDL", 2);
                        List<string> ListaRigheFileSQL = new List<string>();

                        Logger.PrintLC("Lettura file " + FileDaElaborareSQL, 3, ConfigFile.INFO);
                        if (FileOps.LeggiFile(FileDaElaborareSQL, ref ListaRigheFileSQL))
                        {
                            //info lette correttamente
                            Logger.PrintLC(ListaRigheFileSQL.Count + " righe lette nel file " + FileDaElaborareSQL, 3, ConfigFile.INFO);

                            //estrazione elenco entity dalle righe del file sql
                            Logger.PrintLC("Estrazione entity da " + FileDaElaborareSQL, 3, ConfigFile.INFO);
                            List<string> CollezioneEntity = SqlOps.CollezionaEntity(ListaRigheFileSQL);
                            Logger.PrintLC("Entity trovate in " + FileDaElaborareSQL, 3, ConfigFile.INFO);

                            if (SqlOps.CompareEntity(CollezioneEntity, EntityElaborate, ref CompareResults))
                            {
                                Logger.PrintLC("Comparazione riuscita tra " + FileElaborato + " e " + FileDaElaborareSQL, 3, ConfigFile.INFO);
                            }
                            else
                            {
                                Logger.PrintLC("Comparazione non riuscita tra " + FileElaborato + " e " + FileDaElaborareSQL, 3, ConfigFile.ERROR);
                            }
                        }
                        else
                        {
                            //info non lette correttamente
                            Logger.PrintLC("Lettura non riuscita del file " + FileDaElaborareSQL, 3, ConfigFile.ERROR);
                        }
                        
                        FileElaboratiSQL.Add(FileDaElaborareSQL);
                        Logger.PrintLC("** FINE ELABORAZIONE - TABELLE parsing da DDL", 2);
                        #endregion

                        #region ScritturaFileXLS
                        Logger.PrintLC("** INIZIO ELABORAZIONE - TABELLE scrittura risultati differenze DDL <-> XLS", 2);

                        if (ExcelOps.WriteExcelStatsForEntity(new FileInfo(FileDifferenze), CompareResults))
                        {
                            //scrittura excel OK
                            Logger.PrintLC("Scrittura dei risultati dell'elaborazione del file DDL riuscita", 3, ConfigFile.INFO);
                        }
                        else
                        {
                            //scrittura excel KO
                            Logger.PrintLC("Scrittura dei risultati dell'elaborazione del file DDL non riuscita", 3, ConfigFile.ERROR);
                        }

                        Logger.PrintLC("** FINE ELABORAZIONE - TABELLE scrittura risultati differenze DDL <-> XLS", 2);
                        #endregion

                        #region EsameAttributiSQL
                        Logger.PrintLC("** INIZIO ELABORAZIONE - ATTRIBUTI parsing da DDL", 2);
                        
                        if (ListaRigheFileSQL.Count > 0)
                        {
                            //estrazione elenco attributi dalle righe del file sql
                            Logger.PrintLC("Estrazione attributi da " + FileDaElaborareSQL, 3, ConfigFile.INFO);
                            List<string> CollezioneAttributi = SqlOps.CollezionaAttributi(ListaRigheFileSQL);
                            Logger.PrintLC("Attributi trovati in " + FileDaElaborareSQL, 3, ConfigFile.INFO);

                            if (SqlOps.CompareAttribute(CollezioneAttributi, AttributiElaborati, ref CompareResults))
                            {
                                Logger.PrintLC("Comparazione attributi riuscita tra " + FileElaborato + " e " + FileDaElaborareSQL, 3, ConfigFile.INFO);
                            }
                            else
                            {
                                Logger.PrintLC("Comparazione attributi non riuscita tra " + FileElaborato + " e " + FileDaElaborareSQL, 3, ConfigFile.ERROR);
                            }
                        }
                        else
                        {
                            //info non lette correttamente
                            Logger.PrintLC("Lettura non riuscita del file " + FileDaElaborareSQL, 3, ConfigFile.ERROR);
                        }

                        //FileElaboratiSQL.Add(FileDaElaborareSQL);
                        Logger.PrintLC("** FINE ELABORAZIONE - ATTRIBUTI parsing da DDL", 2);
                        #endregion

                        #region ScritturaFileXLS
                        Logger.PrintLC("** INIZIO ELABORAZIONE - ATTRIBUTI scrittura risultati differenze DDL <-> XLS", 2);

                        if (ExcelOps.WriteExcelStatsForAttribute(new FileInfo(FileDifferenze), CompareResults, AttributiElaborati))
                        {
                            //scrittura excel OK
                            Logger.PrintLC("Scrittura dei risultati dell'elaborazione del file DDL riuscita", 3, ConfigFile.INFO);
                        }
                        else
                        {
                            //scrittura excel KO
                            Logger.PrintLC("Scrittura dei risultati dell'elaborazione del file DDL non riuscita", 3, ConfigFile.ERROR);
                        }

                        Logger.PrintLC("** FINE ELABORAZIONE - ATTRIBUTI scrittura risultati differenze DDL <-> XLS", 2);
                        #endregion

                        FileElaboratiSQL.Add(FileElaborato);

                        
                        if (ConfigFile.DEST_FOLD_UNIQUE)
                        {
                            File.Copy(FileDaElaborareSQL, Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileName(FileDaElaborareSQL)));
                        }
                        else
                        {
                            File.Copy(FileDaElaborareSQL, Funct.GetFolderDestination2(FileDaElaborareSQL, fullNameSQL), true);
                        }
                        File.Delete(FileDaElaborareSQL);
                    }
                    else
                    {
                        Logger.PrintLC("Un corrispondente file SQL non esiste per il file " + FileElaborato, 3, ConfigFile.WARNING);
                    }
                    #endregion

                    Logger.PrintLC("** FINE ELABORAZIONE DDL: " + FileElaborato, 2);
                    #endregion

                }
                #region SummaryFiles
                //Stampa elenco completo file presi in considerazione
                Logger.PrintLC("\n## SOMMARIO DIFFERENZE FILE XLS -> SQL:");
                ListaCompleta = Funct.DetermineElaborated(FileElaborati, FileElaboratiSQL);
                foreach (string elemento in ListaCompleta)
                {
                    string myelemento = elemento.ToUpper().Replace(".XLSX", ".SQL");
                    myelemento = myelemento.Replace(".XLS", ".SQL");
                    Logger.PrintLC(myelemento, 2);
                }
                #endregion

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
