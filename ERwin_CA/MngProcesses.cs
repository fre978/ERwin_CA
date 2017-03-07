using ERwin_CA.T;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;


namespace ERwin_CA
{
    static class MngProcesses
    {
        public static int StartProcess()
        {
            try
            {
                if (ConfigFile.RefreshAll() == true)
                {
                    Logger.PrintLC("!! Some error occured while parsing the config file. Standard values will be used instead.", 1, ConfigFile.WARNING);
                }
                List<string> FileElaborati = new List<string>();
                List<ElaboratiT> Elaborati = new List<ElaboratiT>();
                List<string> FileElaboratiSQL = new List<string>();
                string[] ElencoExcel = DirOps.GetFilesToProcess(ConfigFile.ROOT, "*.xls|.xlsx");
                Logger.PrintLC("Excel files found under " + ConfigFile.ROOT + ": " + ElencoExcel.Count(), 2, "INFO: ");
                List<string> FileDaElaborareCompleto = FileOps.GetTrueFilesToProcess(ElencoExcel);
                Logger.PrintLC("Files to validate: " + FileDaElaborareCompleto.Count(), 2, "INFO: ");
                List<string> FileDaElaborare = Parser.ParseListOfFileNames(FileDaElaborareCompleto);
                Logger.PrintLC("Files valid to be processed: " + FileDaElaborare.Count(), 2, "INFO: ");
                List<string> FileDaElaborareRemoto = new List<string>();

                decimal current = 0;
                decimal maximum = 0;
                string message = string.Empty;

                Funct.PrintList(FileDaElaborare);

                //**************************************
                //SEZIONE ROMOTO - PROVA
                if (ConfigFile.COPY_LOCAL)
                {
                    FileDaElaborare = Funct.RemoteGet(FileDaElaborare);
                    if(FileDaElaborare == null)
                        return 5;
                    if (FileDaElaborare.Count == 0)
                        return 2;
                }
                //**************************************



                //####################################
                //Ciclo MAIN
                foreach (var fileC in FileDaElaborare)
                {
                    #region ProcessingFileExcel
                    string file = fileC; //assegnazione mantenere inalterato il ciclo e per poter cambiare il valore di file (causa introduzione Copia Locale)
                    Logger.PrintLC("** START PROCESSING EXCEL FILE: " + file, 2);
                    string TemplateFile = null;
                    string remoteDir = string.Empty;
                    string remoteFile = string.Empty;
                    if (ExcelOps.FileValidation(file))
                    {
                        FileT fileT = Parser.ParseFileName(file);
                        string destERFile = null;
                        if (fileT != null)
                        {
                            TemplateFile = Funct.GetTemplate(fileT);
                            if (TemplateFile == null)
                            {
                                continue;
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
                        }
                            ////aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                            //if (!string.IsNullOrEmpty(dati.History))
                            //{ 
                            //    int col = ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1;
                            //    Logger.PrintLC("Updating excel file for error on entity creation for the table '" + dati.TableName + "' in erwin. Error: " + dati.History, 3);
                            //    //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                            //    if (ConfigFile.DEST_FOLD_UNIQUE)
                            //    {
                            //        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                            //    }
                            //    else
                            //    {
                            //        fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                            //    }

                            //    if (File.Exists(fInfo.FullName))
                            //    {
                            //        ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Row, col, 1, dati.History, ConfigFile.TABELLE);
                            //    }
                            //}
                            //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                            int col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                            Logger.PrintLC("Updating excel file for error on entity creation", 3);
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
                                ExcelOps.XLSXWriteErrorInCell(fInfo, DatiFile.FindAll(x => !string.IsNullOrEmpty(x.History)), col, 1, ConfigFile.TABELLE);
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
                            message = "tabelle create";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            //statistica tabelle senza descrizione
                            current = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History) && string.IsNullOrEmpty(x.TableDescr)).Count();
                            maximum = DatiFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "tabelle senza descrizione";
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
                        List<AttributeT> AttrFile = new List<AttributeT>();
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
                                if (string.IsNullOrEmpty(dati.History))
                                {
                                    connessione.CreateAttributePassOne(dati, TemplateFile);
                                }
                            }

                            //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                            col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                            Logger.PrintLC("Updating excel file for error on attributes creation (pass one)", 3);
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
                                ExcelOps.XLSXWriteErrorInCell(fInfo, AttrFile.FindAll(x => !string.IsNullOrEmpty(x.History) && x.Step == 1), col, 1, ConfigFile.ATTRIBUTI);
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

                                //if (string.IsNullOrEmpty(dato.History))
                                //{
                                //    col = ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1;
                                //    Logger.PrintLC("Updating excel file for error on relation creation for the field '" + dato.IdentificativoRelazione + "' in erwin. Error: " + dato.History, 3);
                                //    //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                //    if (ConfigFile.DEST_FOLD_UNIQUE)
                                //    {
                                //        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                //    }
                                //    else
                                //    {
                                //        fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                //    }
                                //    if (File.Exists(fInfo.FullName))
                                //    {
                                //        ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Relazioni, col, 1, dato.History, ConfigFile.RELAZIONI);
                                //    }
                                //}
                                //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                                //col = ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                                col = ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1;
                                Logger.PrintLC("Updating excel file for error on relation creation", 3);
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
                                    ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Relazioni.FindAll(x => !string.IsNullOrEmpty(x.History)), col, 1, ConfigFile.RELAZIONI);
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
                                if (string.IsNullOrEmpty(dati.History))
                                {
                                    connessione.CreateAttributePassTwo(dati, TemplateFile);
                                }

                                ////aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                                //if (!string.IsNullOrEmpty(dati.History) && (dati.Step == 2))
                                //{
                                //    col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                                //    Logger.PrintLC("Updating excel file for error on attributes creation (pass two) for the field '" + dati.NomeCampoLegacy + "' in erwin. Error: " + dati.History, 3);
                                //    //fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                //    if (ConfigFile.DEST_FOLD_UNIQUE)
                                //    {
                                //        fInfo = new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, Path.GetFileNameWithoutExtension(file) + ".xlsx"));
                                //    }
                                //    else
                                //    {
                                //        fInfo = (new FileInfo(Funct.GetFolderDestination(file, ".xlsx")));
                                //    }
                                //    if (File.Exists(fInfo.FullName))
                                //    {
                                //        ExcelOps.XLSXWriteErrorInCell(fInfo, dati.Row, col, 1, dati.History, ConfigFile.ATTRIBUTI);
                                //    }
                                //}

                            }
                            //aggiorna le info sulle celle del file excel se la creazione fisica in erwin rileva qualche errore
                            col = ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1;
                            Logger.PrintLC("Updating excel file for error on attributes creation (pass two)", 3);
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
                                ExcelOps.XLSXWriteErrorInCell(fInfo, AttrFile.FindAll(x => !string.IsNullOrEmpty(x.History) && x.Step == 2), col, 1, ConfigFile.ATTRIBUTI);
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
                            message = "tabelle senza PK";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            #region AttributiConAlmenoUnErrore
                            //statistica tabelle senza descrizione
                            current = AttrFile.FindAll(x => !(string.IsNullOrEmpty(x.History))).Count;
                            maximum = AttrFile.Count;
                            message = "attributi con almeno un errore";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            #endregion
                            #region AttributiSenzaDescrizione
                            //statistica tabelle senza descrizione
                            current = AttrFile.FindAll(x => string.IsNullOrEmpty(x.History) && string.IsNullOrEmpty(x.DefinizioneCampo)).Count();
                            maximum = AttrFile.FindAll(x => string.IsNullOrEmpty(x.History)).Count;
                            message = "attributi senza descrizione";
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
                            message = "tabelle senza relazioni";
                            Funct.Stats(current, maximum, message, fileCorrect);
                            current = (globalRelationStrut.GlobalRelazioni.Count) - TabelleRelazionate.Count;
                            maximum = (globalRelationStrut.GlobalRelazioni.Count);
                            message = "relazioni con almeno un errore";
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
                            //if (File.Exists(Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xls")))
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
                        ElaboratiT Elaborato = new ElaboratiT(fileElaborato: "", 
                                                              entityElaborate: new List<EntityT>(), 
                                                              attributiElaborati: new List<AttributeT>(), 
                                                              relazioniElaborate: new GlobalRelationStrut());
                        Elaborato.FileElaborato = FileElaborato;
                        Elaborato.EntityElaborate = DatiFile;
                        Elaborato.AttributiElaborati = AttrFile;
                        Elaborato.RelazioniElaborate = globalRelationStrut;
                        Elaborati.Add(Elaborato);
                        
                    }
                    else
                    {
                        string destFile = null;
                        FileInfo origin = new FileInfo(file);
                        //string fileName = Path.GetFileNameWithoutExtension(file);
                        string fileName = origin.Name;
                        
                        
                        if (ConfigFile.DEST_FOLD_UNIQUE)
                        {
                            destFile = Path.Combine(ConfigFile.FOLDERDESTINATION, fileName);
                        }
                        else
                        {
                            destFile = Funct.GetFolderDestination(file, origin.Extension);
                        }
                        if (!FileOps.CopyFile(file, destFile))
                            continue;
                        else
                            origin.Delete();
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


                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################

                foreach (var Elaborato in Elaborati)
                {
                    string FileElaborato = Elaborato.FileElaborato;
                    List<EntityT> EntityElaborate = Elaborato.EntityElaborate;
                    List<AttributeT> AttributiElaborati = Elaborato.AttributiElaborati;
                    GlobalRelationStrut RelazioniElaborate = Elaborato.RelazioniElaborate;
                    List<string> ListaControlliTempistiche = new List<string>();
                    List<string> ListaControlliCompleta = new List<string>();

                    #region DocExcelControlli
                    List<string> DocExcelControlli = new List<string>();
                    if (Elaborato.EntityElaborate.Count != 0)
                    {
                        Logger.PrintLC("** INIZIO ELABORAZIONE CONTROLLI: " + FileElaborato, 2);
                        List<EntityT> EntityBFD = Elaborato.EntityElaborate.FindAll(x => x.FlagBFD == "S" && string.IsNullOrEmpty(x.History));
                        List<RelationStrut> LRelazioniBFD = RelazioniElaborate.GlobalRelazioni;
                        int myprogr = 0;
                        foreach (EntityT E in EntityBFD)
                        {
                            myprogr += 1;
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
                                string Ambito = "BFDL1";
                                string CC = "LI";
                                string DD = E.SSA;
                                string mydb = E.DatabaseName;
                                string alfanum = "00000000000000";
                                if (mydb.Length > 10)
                                {
                                    mydb = mydb.Substring(0, 10);
                                    if (mydb.Contains('_'))
                                    {
                                        mydb = mydb.Split('_')[0];
                                    }
                                    //CODE 66
                                    //if(mydb.EndsWith(",") || mydb.EndsWith(" ") || mydb.EndsWith("-"))
                                    //{
                                    //    mydb = mydb.Substring(0, 9);
                                    //}
                                }
                                alfanum = alfanum.Substring(mydb.Length, alfanum.Length - myprogr.ToString().Length - mydb.Length);
                                alfanum = mydb + alfanum + myprogr;
                                if (type == 0)
                                {
                                    NomeStrutturaInformativa = E.TableName.ToUpper();
                                    NomeCampo = A.NomeCampoLegacy.ToUpper();
                                    CodLocaleControllo = "DUP";
                                    CodLocaleControllo = Ambito + "_" + CC + "_" + DD + "_" + CodLocaleControllo + "_" + alfanum.ToUpper();
                                    RuoloCampo = "OggettoControllo";
                                    //if (!(DocExcelControlli.Exists(x=> x == NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo)))
                                        DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                    //Lista funzionale al file "ControlloTempistiche"
                                    if (!ListaControlliTempistiche.Contains(CodLocaleControllo))
                                        ListaControlliTempistiche.Add(CodLocaleControllo);
                                    ListaControlliCompleta.Add(CodLocaleControllo);
                                }
                                //########################################################################
                                //TEST
                                //if (A.MandatoryFlag == 1 || (A.MandatoryFlag == 0 && A.Chiave == 0))
                                if (A.MandatoryFlag == 1 || A.Chiave == 0)
                                    {
                                        NomeStrutturaInformativa = E.TableName.ToUpper();
                                    NomeCampo = A.NomeCampoLegacy.ToUpper();
                                    CodLocaleControllo = "NUL";
                                    CodLocaleControllo = Ambito + "_" + CC + "_" + DD + "_" + CodLocaleControllo + "_" + alfanum.ToUpper();
                                    RuoloCampo = "OggettoControllo";
                                    //if (!(DocExcelControlli.Exists(x => x == NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo)))
                                        DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                    //Lista funzionale al file "ControlloTempistiche"
                                    if (!ListaControlliTempistiche.Contains(CodLocaleControllo))
                                        ListaControlliTempistiche.Add(CodLocaleControllo);
                                    ListaControlliCompleta.Add(CodLocaleControllo);
                                }
                                if (Funct.ParseDataType(phisical_data_type, A.DB, true))
                                {
                                    NomeStrutturaInformativa = E.TableName.ToUpper();
                                    NomeCampo = A.NomeCampoLegacy.ToUpper();
                                    CodLocaleControllo = "FOR";
                                    CodLocaleControllo = Ambito + "_" + CC + "_" + DD + "_" + CodLocaleControllo + "_" + alfanum.ToUpper();
                                    RuoloCampo = "OggettoControllo";
                                    //if (!(DocExcelControlli.Exists(x => x == NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo)))
                                        DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                    //Lista funzionale al file "ControlloTempistiche"
                                    if (!ListaControlliTempistiche.Contains(CodLocaleControllo))
                                        ListaControlliTempistiche.Add(CodLocaleControllo);
                                    ListaControlliCompleta.Add(CodLocaleControllo);
                                }
                                foreach (RelationStrut SRelazioniBFD in LRelazioniBFD)
                                {
                                    List<RelationT> Relazioni = SRelazioniBFD.Relazioni.FindAll(x => x.CampoFiglio == A.NomeCampoLegacy && x.TabellaFiglia == A.NomeTabellaLegacy && string.IsNullOrEmpty(x.History));
                                    foreach (RelationT Relazione in Relazioni)
                                    {
                                        NomeStrutturaInformativa = E.TableName.ToUpper();
                                        NomeCampo = A.NomeCampoLegacy.ToUpper();
                                        CodLocaleControllo = "DRI";
                                        CodLocaleControllo = Ambito + "_" + CC + "_" + DD + "_" + CodLocaleControllo + "_" + alfanum.ToUpper();
                                        RuoloCampo = "OggettoControllo";
                                        //if (!(DocExcelControlli.Exists(x => x == NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo)))
                                            DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                        //Lista funzionale al file "ControlloTempistiche"
                                        if (!ListaControlliTempistiche.Contains(CodLocaleControllo))
                                            ListaControlliTempistiche.Add(CodLocaleControllo);
                                        ListaControlliCompleta.Add(CodLocaleControllo);

                                        NomeStrutturaInformativa = Relazione.TabellaPadre.ToUpper();
                                        NomeCampo = Relazione.CampoPadre.ToUpper();
                                        CodLocaleControllo = "DRI";
                                        CodLocaleControllo = Ambito + "_" + CC + "_" + DD + "_" + CodLocaleControllo + "_" + alfanum.ToUpper();
                                        RuoloCampo = "CampoConfronto";
                                        //Da DE-COMMENTARE
                                        //if (!(DocExcelControlli.Exists(x => x == NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo)))
                                            DocExcelControlli.Add(NomeStrutturaInformativa + "|" + NomeCampo + "|" + CodLocaleControllo + "|" + RuoloCampo);
                                        //Lista funzionale al file "ControlloTempistiche"
                                        if (!ListaControlliTempistiche.Contains(CodLocaleControllo))
                                            ListaControlliTempistiche.Add(CodLocaleControllo);
                                        ListaControlliCompleta.Add(CodLocaleControllo);
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
                        //###################################################################
                        //####### STAMPA FILE (da de-commentare)
                        ExcelOps.WriteDocExcelControlliCampi(new FileInfo(FileDocControlli), DocExcelControlli);
                        //###################################################################


                        string FileDocControlliTempistiche = Path.GetFileNameWithoutExtension(FileElaborato) + "_ControlliTempistiche.xlsx";
                        //FileDocControlli = Path.Combine(ConfigFile.FOLDERDESTINATION, FileDocControlli);
                        if (ConfigFile.DEST_FOLD_UNIQUE)
                        {
                            FileDocControlliTempistiche = Path.Combine(ConfigFile.FOLDERDESTINATION, FileDocControlliTempistiche);
                        }
                        else
                        {
                            FileDocControlliTempistiche = Funct.GetFolderDestination2(FileElaborato, new FileInfo(FileDocControlliTempistiche).Name);
                        }
                        ExcelOps.WriteDocExcelControlliTempistiche(new FileInfo(FileDocControlliTempistiche), ListaControlliTempistiche);
                        //ExcelOps.WriteDocExcelControlliCampiX(new FileInfo(FileDocControlliTempistiche), ListaControlliTempistiche);

                    }

                    Logger.PrintLC("** FINE ELABORAZIONE CONTROLLI: " + FileElaborato, 2);
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
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################
                //###################################################################






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

                //**************************************
                //SEZIONE ROMOTO - PROVA
                if (ConfigFile.COPY_LOCAL)
                {
                    if (!Funct.RemoteSet(FileDaElaborare))
                    {
                        Logger.PrintLC("Some error occured while copying or ", 2, ConfigFile.ERROR);
                    }
                }
                else
                {
                }
                //**************************************


                return 0;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("UNEXPECTED ERROR: " + exp.Message, 1);//return exp.HResult;
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
                        try
                        {
                            proc.Kill();
                            proc.WaitForExit();
                        }
                        catch { }
                    }
                }
            }
            catch (System.NullReferenceException)
            {

            }
        }

    }
}
