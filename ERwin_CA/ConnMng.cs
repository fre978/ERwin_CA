using ERwin_CA.T;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VBClassLibrary;


namespace ERwin_CA
{
    class ConnMng
    {
        public SCAPI.Application scERwin;
        public SCAPI.PersistenceUnit scPersistenceUnit;
        public SCAPI.Session scSession;
        public SCAPI.ModelObject erRootObj { get; set; }
        public SCAPI.ModelObjects erRootObjCol { get; set; }
        public SCAPI.ModelObjects erColumn { get; set; }
        public SCAPI.ModelObjects erTable { get; set; }
        public SCAPI.ModelObjects erObjectCollection { get; set; }
        public SCAPI.ModelObjects model { get; set; }
        public SCAPI.ModelObject erEntityObjectPE;
        public SCAPI.ModelObjects erAttributeObjCol { get; set; } //utilizzato nella creazione degli Attributi.
        public SCAPI.ModelObject erAttributeObjectPE;
        public SCAPI.ModelObject scItem;
        public SCAPI.ModelObject scDB;
        public SCAPI.ModelObject scSchema;
        public string fileERwin = null;
        public FileInfo fileInfoERwin = null;
        public List<string> DatabaseN = null;
        public List<string> SchemaN = null;
        public object trID { get; set; }
        public object lastIdCommitted { get; set; }

        public bool openModelConnection(string ERw)
        {
            if (ERw == null)
                return false;
            if (!File.Exists(ERw))
            {
                Logger.PrintLC("Could not find file: " + ERw, 2,ConfigFile.ERROR);
                return false;
            }
            
            try
            {
                DatabaseN = new List<string>();
                SchemaN = new List<string> { "SYSADM", "SYSFUN", "SYSIBM", "SYSPROC", "SYSTOOLS" };
                fileERwin = ERw;
                FileOps.RemoveAttributes(ERw);
                fileInfoERwin = new FileInfo(fileERwin);
                scERwin = new SCAPI.Application();

                scPersistenceUnit = scERwin.PersistenceUnits.Add(ERw, "RDO=No");

                scSession = scERwin.Sessions.Add();
                scSession.Open(scPersistenceUnit);
                Logger.PrintLC("Connection opened.",2, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Connection opening error: " + exp.Message, 2, ConfigFile.ERROR);
                return false;
            }
        }

        public void CloseModelConnection()
        {
            try
            {
                scERwin.PersistenceUnits.Clear();
                SCAPI.Sessions scSessionCol = scERwin.Sessions;
                foreach (SCAPI.Session scSes in scSessionCol)
                    scSes.Close();

                while (scSessionCol.Count > 0)
                    scSessionCol.Remove(0);

                scPersistenceUnit = null;
                //scSession.Close();
                Logger.PrintLC("Session closed successfully.", 2, ConfigFile.INFO);
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Could not close the Session.", 2, ConfigFile.ERROR);
            }
        }



        public object OpenTransaction()
        {
            if (scSession != null)
                try
                {
                    trID = scSession.BeginTransaction();
                    Logger.PrintLC("Transaction began successfully.", 3, ConfigFile.INFO);
                    return trID;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Starting Transaction error: " + exp.Message, 3, ConfigFile.ERROR);
                    return -1;
                }
            else
                Logger.PrintLC("Starting Transaction error: missing SESSION.", 3, ConfigFile.ERROR);
            return -1;
        }

        public bool SetRootObject()
        {
            if (scSession != null)
                try
                {
                    erRootObj = scSession.ModelObjects.Root;
                    Logger.PrintLC("Root has been successful.", 3, ConfigFile.INFO);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Setting Root's Session error: " + exp.Message, 3, ConfigFile.ERROR);
                    return false;
                }
            else
                Logger.PrintLC("Could not determine Root because Session is missing.", 3, ConfigFile.ERROR);
            return false;
        }

        public bool SetRootCollection()
        {
            if (scSession != null)
                try
                {
                    erRootObjCol = scSession.ModelObjects.Collect(erRootObj);
                    Logger.PrintLC("Root Collection has been successful.", 3, ConfigFile.INFO);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not get the Root Collection: " + exp.Message, 3, ConfigFile.ERROR);
                    return false;
                }
            else
            {
                Logger.PrintLC("Could not get Root Collection because Session is missing.", 3, ConfigFile.ERROR);
                return false;
            }
        }

        public SCAPI.ModelObject CreateModel(string ModelName)
        {
            SCAPI.ModelObject ret = null;
            string errore = string.Empty;
            

            if (erRootObjCol != null)
            {
                OpenTransaction();
                
                VBCon con = new VBCon();

                string nome = string.Empty;
                if (!con.RetrieveFromObjModel(scSession.ModelObjects.Root, "Name", ref nome))
                {
                    errore = "Error while retrieving property Name from Model object";
                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    return ret;
                }
                else
                {
                    SCAPI.ModelObject root = scSession.ModelObjects.Root;
                    if (con.AssignToObjModel(ref root, "Name", ModelName))
                    {
                        Logger.PrintLC("Renamed Model object as " + ModelName, 3, ConfigFile.INFO);
                    }
                    else
                    {
                        errore = "Impossible to rename Model Object" + scItem.ObjectId;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        CommitAndSave(trID);
                        return ret;
                    }
                }

                CommitAndSave(trID);
            }
            return ret;
        }

        public SCAPI.ModelObject CreateEntity (EntityT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            string errore = string.Empty;
            if (string.IsNullOrWhiteSpace(db))
            {
                //errore = "There was no DB associated to " + entity.TableName;
                errore = "Non esiste un DB associato a " + entity.TableName;
                if (entity.History != null)
                    errore = "\n" + errore;
                entity.History += errore;
                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                return ret;
            }
            //if (!((entity.FlagBFD == "S") || (entity.FlagBFD == "N")))
            if (!(Funct.ParseFlag(entity.FlagBFD, "YES") || Funct.ParseFlag(entity.FlagBFD, "NO")))
            {
                //errore = "Property FlagBFD of " + entity.TableName + " is not valid. Table will be skipped.";
                errore = "La proprietà FlagBFD di " + entity.TableName + " non è valida. La tabella verrà saltata.";
                if (entity.History != null)
                    errore = "\n" + errore;
                entity.History += errore;
                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                return ret;
            }

            if (erRootObjCol != null)
            {
                OpenTransaction();
                scItem = erRootObjCol.Add("Entity");
                VBCon con = new VBCon();

                //Controlli proprietà essenziali
                //Nome tabella
                if (!string.IsNullOrWhiteSpace(entity.TableName))
                {
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Tabella"], entity.TableName.ToUpper()))
                        Logger.PrintLC("Added Table Physical Name (" + entity.TableName + ") to " + scItem.ObjectId, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Table Physical Name (" + entity.TableName + ") to " + scItem.ObjectId;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome Tabella"] + " (" + entity.TableName + ") a " + scItem.ObjectId;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        CommitAndSave(trID);
                        return scItem;
                    }
                    if (con.AssignToObjModel(ref scItem, "Name", entity.TableName.ToUpper()))
                        Logger.PrintLC("Added Table Name to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Table Name to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome Tabella"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        CommitAndSave(trID);
                        return scItem;
                    }
                }

                //SSA
                if (!string.IsNullOrWhiteSpace(entity.SSA))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["SSA"], entity.SSA))
                        Logger.PrintLC("Added SSA to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding SSA to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["SSA"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }

                //Table Description
                if (!string.IsNullOrWhiteSpace(entity.TableDescr))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Descrizione Tabella"], entity.TableDescr))
                        Logger.PrintLC("Added Table Description to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Table Description to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Descrizione Tabella"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }
                if (!string.IsNullOrWhiteSpace(entity.TableDescr))
                    if (con.AssignToObjModel(ref scItem, "Definition", entity.TableDescr))
                        Logger.PrintLC("Added Table Definition to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        errore = "Errore riscontrato aggiungendo Table Definition a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }


                //Info Type
                if (!string.IsNullOrWhiteSpace(entity.InfoType))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Tipologia Informazione"], entity.InfoType))
                        Logger.PrintLC("Added Information Type to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Information Type to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo "+ ConfigFile._TAB_NAME["Tipologia Informazione"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }

                //Table Limit
                if (!string.IsNullOrWhiteSpace(entity.TableLimit))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Perimetro Tabella"], entity.TableLimit))
                        Logger.PrintLC("Added Table Limit to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Table Limit to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Perimetro Tabella"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }

                //Table Granularity
                if (!string.IsNullOrWhiteSpace(entity.TableGranularity))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Granularità Tabella"], entity.TableGranularity))
                        Logger.PrintLC("Added Table Granularity to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Table Granularity to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Granularità Tabella"] + " a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }

                //Flag BFD
                if (!string.IsNullOrWhiteSpace(entity.FlagBFD))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Flag BFD"], entity.FlagBFD))
                        Logger.PrintLC("Added Flag BFD to " + scItem.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Flag BFD to " + scItem.Name;
                        errore = "Errore riscontrato aggiungendo "+ ConfigFile._TAB_NAME["Flag BFD"] +" a " + scItem.Name;
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    }



                //##################################################
                //** Assegnazione dei valori NOME_HOST e          **
                //** NOME_DATABASE nei rispettivi campi Erwin     **
                //** in base al DB di origine.                    **
                if (entity.DB == "DB2")
                {
                    //************************************************
                    //TEST assegnazione HostName DB2
                    //TESTUNO
                    if (!string.IsNullOrWhiteSpace(entity.HostName))
                    {
                        if (con.AssignToObjModel(ref scItem, "Entity.Physical.NOME_HOST", entity.HostName))
                            Logger.PrintLC("Added Host Name DB2 to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Name to " + scDB.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome host DB2_SQL"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    {
                        if (con.AssignToObjModel(ref scItem, "Entity.Physical.NOME_DATABASE", entity.DatabaseName))
                            Logger.PrintLC("Added Database Name DB2 to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Name to " + scDB.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome Database DB2_SQL"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                    //************************************************
                }

                if(entity.DB == "SQLSERVER")
                {
                    //************************************************
                    //TEST assegnazione HostName DB2
                    //TESTUNO
                    if (!string.IsNullOrWhiteSpace(entity.HostName))
                    {
                        if (con.AssignToObjModel(ref scItem, "Entity.Physical.NOME_HOST", entity.HostName))
                            Logger.PrintLC("Added Host Name SQLServer to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Name to " + scDB.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome host DB2_SQL"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    {
                        if (con.AssignToObjModel(ref scItem, "Entity.Physical.NOME_DATABASE", entity.DatabaseName))
                            Logger.PrintLC("Added Database Name SQLServer to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Name to " + scDB.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome Database DB2_SQL"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                    //************************************************
                }

                //##################################################
                //** Controllo esistenza DB ed eventuale aggiunta **
                //** Qui vanno aggiunti eventuali altri DB da     **
                //** prendere in considerazione oltre a DB2,      **
                //** ORACLE e SQL SERVER                          **
                if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                {
                    if (entity.DB == "DB2")
                    {
                        if (!DatabaseN.Contains(entity.DatabaseName))
                        {
                            scDB = erRootObjCol.Add("DB2_Database");
                            if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome Database"], entity.DatabaseName))
                                Logger.PrintLC("Added Database Name to " + scDB.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Database Name to " + scDB.Name;
                                errore = "Errore riscontrato aggiungendo "+ ConfigFile._TAB_NAME["Nome Database"]+ " a " + scDB.Name;
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                            }

                            if (!string.IsNullOrWhiteSpace(entity.HostName))
                            {
                                if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome host"], entity.HostName))
                                    Logger.PrintLC("Added Host Name to " + scDB.Name, 3, ConfigFile.INFO);
                                else
                                {
                                    //errore = "Error adding Host Name to " + scDB.Name;
                                    errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome host"] + " a " + scDB.Name;
                                    if (entity.History != null)
                                        errore = "\n" + errore;
                                    entity.History += errore;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                }
                            }
                            DatabaseN.Add(entity.DatabaseName);
                        }
                    }
                }

                if (entity.DB == "SQLSERVER")
                {
                    if (!DatabaseN.Contains(entity.DatabaseName))
                    {
                        scDB = erRootObjCol.Add("SQLServer_Database");
                        if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome Database SQLSERVER"], entity.DatabaseName))
                            Logger.PrintLC("Added Database Name to " + scDB.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Database Name to " + scDB.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome Database SQLSERVER"] + " a " + scDB.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }

                        if (!string.IsNullOrWhiteSpace(entity.HostName))
                            if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome host SQLSERVER"], entity.HostName))
                                Logger.PrintLC("Added Host Name to " + scDB.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Host Name to " + scDB.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome host SQLSERVER"] + " a " + scDB.Name;
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                            }
                        DatabaseN.Add(entity.DatabaseName);
                    }

                }

                if (entity.DB == "ORACLE")
                {
                    if (!string.IsNullOrWhiteSpace(entity.HostName))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome host Oracle"], entity.HostName))
                            Logger.PrintLC("Added Host Oracle Name to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Oracle Name to " + scItem.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Nome host Oracle"] +" a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Database Oracle"], entity.DatabaseName))
                            Logger.PrintLC("Added Oracle Database Name to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Oracle Database Name to " + scItem.Name;
                            errore = "Errore riscontrato aggiungendo "+ ConfigFile._TAB_NAME["Nome Database Oracle"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                }

                //##################################################

                //##################################################
                //## Controllo esistenza SCHEMA ed eventuale aggiunta ##
                //** Aggiungere qui eventuali altri casi DB oltre a DB2/ORACLE
                if ((entity.DB == "DB2") || (entity.DB == "SQLSERVER"))
                {
                    if (!string.IsNullOrWhiteSpace(entity.Schema))
                    {
                        if (!SchemaN.Contains(entity.Schema))
                        {
                            scSchema = erRootObjCol.Add("Schema");
                            if (con.AssignToObjModel(ref scSchema, "Name", entity.Schema))
                                Logger.PrintLC("Created Schema Name to " + scSchema.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error creating Schema Name to " + scSchema.Name;
                                errore = "Errore riscontrato creando Schema Name a " + scItem.Name;
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                            }
                            SchemaN.Add(entity.Schema);
                        }
                        //Schema
                        if (!string.IsNullOrWhiteSpace(entity.Schema))
                            if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Schema"], entity.Schema))
                                Logger.PrintLC("Added Schema to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Schema to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Schema"] + " a " + scItem.Name;
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                            }
                    }
                }
                if(entity.DB == "ORACLE")
                {
                    if (!string.IsNullOrWhiteSpace(entity.Schema))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Schema Oracle"], entity.Schema))
                            Logger.PrintLC("Added Host Oracle Schema to " + scItem.Name, 3, ConfigFile.INFO);
                        else
                        {
                            //errore = "Error adding Host Oracle Schema to " + scItem.Name;
                            errore = "Errore riscontrato aggiungendo " + ConfigFile._TAB_NAME["Schema Oracle"] + " a " + scItem.Name;
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        }
                    }
                }
                //##################################################
                CommitAndSave(trID);
            }
            return scItem;
        }

        public SCAPI.ModelObject CreateRelation(RelationStrut relation, string db, GlobalRelationStrut globalRelation)
        {
            //test esistenza file erwin
            SCAPI.ModelObject ret = null;
            if (string.IsNullOrWhiteSpace(db))
            {
                Logger.PrintLC("There was no DB associated to " + relation.ID, 3, ConfigFile.ERROR);
                return ret;
            }

            //verifico che sia valorizzata la root collection
            if (erRootObjCol != null)
            {
                try
                {
                    OpenTransaction();

                    //collezione completa delle entity
                    erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                    int countRelazioni = relation.Relazioni.Count;
                    string campoPadreDaCercare = string.Empty;
                    string campoPadreTrovato = string.Empty;
                    bool PrimoGiro = true;
                    VBCon con = new VBCon();

                    //campi per il controllo di uguaglianza all'interno dello stesso ID relazione
                    string _TabellaPadre = string.Empty;
                    string _TabellaFiglia = string.Empty;
                    string _CampoPadre = string.Empty;
                    string _CampoFiglio = string.Empty;
                    int? _Identificativa = null;
                    int? _Cardinalita = null;
                    bool? _TipoRelazione = null;

                    string errore = string.Empty;

                    int countKey = 0;

                    List<string> RelazioniOk = new List<string>();

                    SCAPI.ModelObject tabellaPadre = null;
                    SCAPI.ModelObject tabellaFiglio = null;
                    SCAPI.ModelObject campoPadre = null;
                    SCAPI.ModelObject campoFiglio = null;

                    bool isNotIdentificativa = false;
                    bool checkNotIdentificativa = false;

                    foreach (var R in relation.Relazioni)
                    {
                        erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                        string isRelKey = string.Empty;
                        errore = string.Empty;
                        tabellaPadre = null;
                        tabellaFiglio = null;
                        campoPadre = null;
                        campoFiglio = null;
                        RelazioniOk.Remove(R.IdentificativoRelazione);
                        if (string.IsNullOrEmpty(R.History))
                        {
                            #region verificheErwin

                            #region controlliTabellaPadre

                            // DEBUG
                            if (R.TabellaPadre == "Y0PRET")
                                Logger.PrintC("DEBUG");
                            //#########

                            //cerchiamo la tabella padre
                            if (!con.RetriveEntity(ref tabellaPadre, erObjectCollection, R.TabellaPadre.ToUpper()))
                            {
                                //errore = "Relation ignored: Could not find table " + R.TabellaPadre + " inside relation ID " + relation.ID;
                                errore = "Relazione ignorata: impossibile trovare la tabella " + R.TabellaPadre + " nel foglio tabelle";
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return ret = null;
                                continue;
                            }
                            else
                            {
                                //verifica che la tabella padre sia sempre la medesima su tutte le righe della relazione
                                if (PrimoGiro)
                                {
                                    _TabellaPadre = R.TabellaPadre;
                                }
                                else
                                {
                                    if (_TabellaPadre == R.TabellaPadre)
                                    {
                                        //la tabella padre è la stessa delle righe precedenti della stessa relazione  
                                    }
                                    else
                                    {
                                        //Non è possibile tracciare una relazione con tabelle differenti
                                        //errore = "Relation ignored: Cannot trace relationship with a different father table inside the same relation" + relation.ID;
                                        errore = "Relatione ignorata: impossibile tracciare la relazione con una tabella padre differente all'interno della stessa relazione" + relation.ID;
                                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return ret = null;
                                        continue;
                                    }
                                }

                            }
                            #endregion

                            #region controlliTabellaFiglia
                            //cerchiamo la tabella figlia
                            if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia.ToUpper()))
                            {
                                //errore = "Relation Ignored: Could not find table " + R.TabellaFiglia + " inside relation ID " + relation.ID;
                                errore = "Relazione ignorata: impossibile trovare la tabella " + R.TabellaFiglia + " nel foglio tabelle";
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return ret = null;
                                continue;
                            }
                            else
                            {
                                //verifica che la tabella figlia sia sempre la medesima su tutte le righe della relazione
                                if (PrimoGiro)
                                {
                                    _TabellaFiglia = R.TabellaFiglia;
                                }
                                else
                                {
                                    if (_TabellaFiglia == R.TabellaFiglia)
                                    {
                                        //la tabella figlia è la stessa delle righe precedenti della stessa relazione  
                                    }
                                    else
                                    {
                                        //Non è possibile tracciare una relazione con tabelle differenti
                                        //errore = "Relation ignored: Cannot trace relationship with a different child table inside the same relation " + relation.ID;
                                        errore = "Relazione ignorata: impossibile tracciare una relazione con una tabella figlia differente all'interno di una stessa relazione. ID: " + relation.ID;
                                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return ret = null;
                                        continue;
                                    }
                                }
                            }
                            #endregion

                            //Code 80
                            #region controlliCampoPadre
                            //esistenza campo padre
                            SCAPI.ModelObjects erAttributesPadre = scSession.ModelObjects.Collect(tabellaPadre, "Attribute");
                            if (!con.RetriveAttribute(ref campoPadre, erAttributesPadre, R.CampoPadre.ToUpper()))
                            {
                                //errore = "Relation Ignored: Could not find field " + R.CampoPadre + " inside table " + R.TabellaPadre + " with relation ID " + relation.ID;
                                errore = "Relazione ignorata: impossibile trovare il campo " + R.CampoPadre + " all'interno della tabella " + R.TabellaPadre + " per la relazione ID: " + relation.ID;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return ret = null;
                                continue;
                            }
                            else
                            {
                                //verifica che sia key
                                string isKey = null;
                                _CampoPadre = R.CampoPadre;
                                campoPadreTrovato = "N";

                                //scandaglio tutte gli attributi dell'entity padre
                                int contatore = 0;
                                int contaPadre = 0;
                                //Code 78
                                //countKey = 0;
                                #region cicloAttributiTabellaPadre
                                foreach (SCAPI.ModelObject attributo in erAttributesPadre)
                                {
                                    contatore++;
                                    // ogni colonna deve avere un valore chiave
                                    if (!con.RetrieveFromObjModel(attributo, "Type", ref isKey))
                                    {
                                        //errore = "Relation Ignored: Could not find attribute Type of field " + R.CampoPadre + " inside table " + R.TabellaPadre + " with relation ID " + relation.ID;
                                        errore = "Relazione ignorata: Impossibile trovare l'attributo Type del campo " + R.CampoPadre + " all'interno della tabella " + R.TabellaPadre + " per la relazione ID: " + relation.ID;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        continue;
                                    }
                                    else
                                    {
                                        //se è chiave primaria verifico che sia quella della colonna che sto cercando
                                        if (isKey == "0")
                                        {
                                            string testPhysicalOrder = null;
                                            if (con.RetrieveFromObjModel(attributo, "Physical_Order", ref testPhysicalOrder))
                                            {
                                                //se siamo al primo giro contiamo le chiavi, dai giri successivi lo sappiamo.
                                                if (PrimoGiro)
                                                    countKey += 1;
                                            }

                                            if (attributo.Name == _CampoPadre)
                                            {
                                                contaPadre++;
                                                campoPadreTrovato = "S";
                                                //bypasso il ciclo perche ho trovato l'attributo di cui desideravo verificare la chiave ma solo dal secondo giro del ciclo
                                                if (!(PrimoGiro))
                                                    continue;
                                            }
                                        }
                                    }

                                }
                                #endregion
                                if (campoPadreTrovato == "S")
                                {
                                    _CampoPadre = string.Empty;
                                    campoPadreTrovato = string.Empty;
                                }
                                else
                                {
                                    //errore = "Relation Ignored: Unmatching PK fields in table " + R.TabellaPadre + " with relation ID " + relation.ID + ": " + _CampoPadre + " is not a key of the table";
                                    errore = "Relazione ignorata: corrispondenza mancante per il campo PK nella tabella " + R.TabellaPadre + " per la relazione ID " + relation.ID + ": " + _CampoPadre + " non è una chiave primaria";
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    if (trID != lastIdCommitted)
                                        CommitAndSave(trID);
                                    //return ret = null;
                                    continue;
                                }
                                if (!(countKey == countRelazioni))
                                {
                                    //errore = "Relation Ignored: Unmatching PK numbers in table " + R.TabellaPadre + " with relation ID " + relation.ID;
                                    errore = "Relazione ignorata: Corrispondenza mancante nel numero di PK della tabella " + R.TabellaPadre + " per la relazione ID " + relation.ID;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    if (trID != lastIdCommitted)
                                        CommitAndSave(trID);
                                    //return ret = null;
                                    continue;
                                }
                            }
                            #endregion

                            #region controlliIdentificativa
                            if (PrimoGiro)
                            {
                                _Identificativa = R.Identificativa;
                            }
                            else
                            {
                                if (_Identificativa == R.Identificativa)
                                {
                                    //la tabella figlia è la stessa delle righe precedenti della stessa relazione  
                                }
                                else
                                {
                                    //Non è possibile tracciare una relazione con tabelle differenti
                                    errore = "Relazione ignorata: Impossibile tracciare una relazione con Identificativa differente all'interno della stessa relazione. ID: " + relation.ID;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return ret = null;
                                    continue;
                                }
                            }
                            #endregion

                            #region controlliCampoFiglio
                            //esistenza campo figlio
                            SCAPI.ModelObjects erAttributesFiglio = scSession.ModelObjects.Collect(tabellaFiglio, "Attribute");
                            if (!con.RetriveAttribute(ref campoFiglio, erAttributesFiglio, R.CampoFiglio.ToUpper()))
                            {
                                //errore = "Relation Ignored Could not find Child Field " + R.CampoFiglio + " inside Child Table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                errore = "Relazione ignorata: Impossibile trovare il Child Field " + R.CampoFiglio + " all'interno della Child Table " + R.TabellaFiglia + " per la relazione ID: " + relation.ID;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return ret = null;
                                continue;
                            }
                            else
                            {
                                //code 77
                                if (con.RetrieveFromObjModel(campoFiglio, "Type", ref isRelKey))
                                {
                                    if (isRelKey == "0")
                                        R.CampoFiglioKey = true;
                                    else
                                        R.CampoFiglioKey = false;
                                }
                                // Se la relazione è di tipo identificativa
                                if (R.Identificativa == 2)
                                {
                                    //i campi di una relazione identificativa nella tabella figlio devono essere tutti di tipo chiave
                                    string isKey = null;
                                    if (!con.RetrieveFromObjModel(campoFiglio, "Type", ref isKey))
                                    {
                                        //errore = "Relation Ignored: Could not find attribute Type of field " + R.CampoFiglio + " inside child table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                        errore = "Relazione ignorata: impossibile trovare l'attributo Type del campo " + R.CampoFiglio + " all'interno della child table " + R.TabellaFiglia + " per la relazione ID: " + relation.ID;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return ret = null;
                                        continue;
                                    }
                                    else
                                    {
                                        if (isKey != "0")
                                        {
                                            //errore = "Relation Ignored: " + R.CampoFiglio + "expected Key inside child table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                            errore = "Relazione ignorata: " + R.CampoFiglio + " deve essere chiave all'interno della child table " + R.TabellaFiglia + " per la relazione ID: " + relation.ID;
                                            Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                            if (R.History != null)
                                                errore = "\n" + errore;
                                            R.History += errore;
                                            CommitAndSave(trID);
                                            //return ret = null;
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    //i campi di una relazione non identificativa nella tabella figlio non possono essere tutti di tipo chiave
                                    string isKey = null;
                                    checkNotIdentificativa = true;
                                    if (!con.RetrieveFromObjModel(campoFiglio, "Type", ref isKey))
                                    {
                                        //errore = "Relation Ignored: Could not find attribute Type of field " + R.CampoFiglio + " inside child table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                        errore = "Relazione ignorata: Impossibile trovare l'attributo Type del campo " + R.CampoFiglio + " all'interno della child table " + R.TabellaFiglia + " per la relazione ID " + relation.ID;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return ret = null;
                                        continue;
                                    }
                                    else
                                    {
                                        //se almeno uno dei campi della relazione non identificativa non è key la relazione può essere tracciata
                                        if (isKey != "0")
                                        {
                                            string testPhysicalOrder = null;
                                            if (con.RetrieveFromObjModel(campoFiglio, "Physical_Order", ref testPhysicalOrder))
                                            {
                                                //se siamo al primo giro contiamo le chiavi, dai giri successivi lo sappiamo.
                                                isNotIdentificativa = true;
                                            }

                                            /* ORIGINALE (eliminare le 6 righe precedenti)
                                            isNotIdentificativa = true;
                                            */
                                        }
                                        else
                                        {
                                            //R.CampoFiglioKey = false;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region controlliCardinalità
                            if (PrimoGiro)
                                _Cardinalita = R.Cardinalita;
                            else
                            {
                                if (_Cardinalita == R.Cardinalita)
                                {
                                    //la tabella padre è la stessa delle righe precedenti della stessa relazione  
                                }
                                else
                                {
                                    //Non è possibile tracciare una relazione con tabelle differenti
                                    //errore = "Relation ignored: Cannot trace relationship with a different cardinality inside the same relation" + relation.ID;
                                    errore = "Relazione ignorata: Impossibile tracciare una relazione con una differente cardinalità all'interno della stessa relazione. ID: " + relation.ID;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return ret = null;
                                    continue;
                                }
                            }
                            #endregion

                            #region controlliTipoRelazione
                            if (PrimoGiro)
                                _TipoRelazione = R.TipoRelazione;
                            else
                            {
                                if (_TipoRelazione == R.TipoRelazione)
                                {
                                    //la tabella padre è la stessa delle righe precedenti della stessa relazione  
                                }
                                else
                                {
                                    //Non è possibile tracciare una relazione con tabelle differenti
                                    //errore = "Relation ignored: Cannot trace relationship with a different relation type inside the same relation. ID: " + relation.ID;
                                    errore = "Relazione ignorata: impossibile tracciare una relazione con un differente Relation Type all'interno della stessa relazione. ID: " + relation.ID;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return ret = null;
                                    continue;
                                }
                            }
                            #endregion
                            
                            #endregion
                            // Code 77 Input
                            PrimoGiro = false;
                            RelazioniOk.Add(R.IdentificativoRelazione);
                        }
                    }

                    if (relation.Relazioni.Exists(x => x.History != null))
                    {
                        if (checkNotIdentificativa == true)
                        {
                            checkNotIdentificativa = false;
                        }
                    }
                    if (!(isNotIdentificativa) && (checkNotIdentificativa))
                    {
                        RelazioniOk.Remove(relation.ID);
                        //errore = "Relation Ignored: all the attribute used for the relation inside child table " + _TabellaFiglia + " are keys. Relation " + relation.ID + " must be an identity relation";
                        errore = "Relazione ignorata: tutti gli attributi usati per la relazione all'interno della child table " + _TabellaFiglia + " sono chiavi. La relazione " + relation.ID + " DEVE essere di tipo Identificativa";
                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                        foreach (var R in relation.Relazioni)
                        {
                            if (string.IsNullOrEmpty(R.History))
                            {
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                            }
                        }
                        if (trID != lastIdCommitted)
                            CommitAndSave(trID);
                        //return ret = null;
                    }


                    if (trID != lastIdCommitted)
                        CommitAndSave(trID);
                    
                    OpenTransaction();
                    
                    #region creazionerelazioni;

                    //creare relazione su erwin
                    SetRootObject();
                    SetRootCollection();


                    scItem = erRootObjCol.Add("Relationship");
                    foreach (var R in relation.Relazioni)
                    {
                        errore = string.Empty;
                        tabellaPadre = null;
                        tabellaFiglio = null;
                        campoPadre = null;
                        campoFiglio = null;
                        if (RelazioniOk.Exists(x => x == R.IdentificativoRelazione))
                        {
                            //La relazione ha passato i controlli erwin e può essere creata
                            #region assegnaIdentificativoRelazione
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Identificativo relazione"], R.IdentificativoRelazione))
                                Logger.PrintLC("Added Relation's Id (" + R.IdentificativoRelazione + ") to " + scItem.ObjectId, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Id (" + R.IdentificativoRelazione + ") to " + scItem.ObjectId;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Identificativo relazione"]  + " ID (" + R.IdentificativoRelazione + ") a " + scItem.ObjectId;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region assegnaTabellaPadre
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tabella Padre"], R.TabellaPadre.ToUpper()))
                                Logger.PrintLC("Added Relation's Parent Table (" + R.TabellaPadre + ") to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Parent Table (" + R.TabellaPadre + ") to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Tabella Padre"] + " (" + R.TabellaPadre + ") a " + scItem.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region assegnaTabellaFiglia
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tabella Figlia"], R.TabellaFiglia.ToUpper()))
                                Logger.PrintLC("Added Relation's Child Table (" + R.TabellaFiglia + ") to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Child Table (" + R.TabellaFiglia + ") to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Tabella Figlia"] + " (" + R.TabellaFiglia + ") a " + scItem.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region assegnaIdentificativa
                            if (con.AssignToObjModelInt(ref scItem, ConfigFile._REL_NAME["Identificativa"], (int)R.Identificativa))
                                Logger.PrintLC("Added Relation's Identifiable (" + R.Identificativa + ") to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Identifiable (" + R.Identificativa + ") to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Identificativa"] + " (" + R.Identificativa + ") a " + scItem.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region assegnaCardinalita
                            int myInt = (R.Cardinalita == null) ? 0 : (int)R.Cardinalita;
                            if (con.AssignToObjModelInt(ref scItem, ConfigFile._REL_NAME["Cardinalita"], myInt))
                                Logger.PrintLC("Added Relation's Cardinality (" + R.Cardinalita + ") to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Cardinality (" + R.Cardinalita + ") to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Cardinalita"] + " (" + R.Cardinalita + ") a " + scItem.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region NullOptionType
                            if (R.NullOptionType != null)
                            {
                                int myNullOT = (R.NullOptionType == null) ? 0 : (int)R.NullOptionType;
                                if (con.AssignToObjModelInt(ref scItem, ConfigFile._REL_NAME["Null Option Type"], myNullOT))
                                    Logger.PrintLC("Added Relation's Null Option Type (" + myNullOT + ") to " + scItem.Name, 3, ConfigFile.INFO);
                                else
                                {
                                    //errore = "Error adding Relation Cardinality (" + R.Cardinalita + ") to " + scItem.Name;
                                    errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Null Option Type"] + " (" + myNullOT + ") a " + scItem.Name;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                            }
                            #endregion
                            #region assegnaTipoRelazione
                            string mystring = (R.TipoRelazione == true) ? "true" : "false";
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tipo Relazione"], mystring))
                                Logger.PrintLC("Added Relation's Type (" + R.TipoRelazione + ") to " + scItem.Name, 3, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Relation Type (" + R.TipoRelazione + ") to " + scItem.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Tipo Relazione"] + " (" + R.TipoRelazione + ") a " + scItem.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (R.History != null)
                                    errore = "\n" + errore;
                                R.History += errore;
                                CommitAndSave(trID);
                                //return scItem;
                                continue;
                            }
                            #endregion
                            #region assegnaNote
                            if (!string.IsNullOrWhiteSpace(R.Note))
                            {
                                if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Note"], R.Note))
                                    Logger.PrintLC("Added Relation's Note (" + R.Note + ") to " + scItem.Name, 3, ConfigFile.INFO);
                                else
                                {
                                    //errore = "Error adding Relation Note (" + R.Note + ") to " + scItem.Name;
                                    errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Note"] + " (" + R.Note + ") a " + scItem.Name;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                            }
                            #endregion
                            #region assegnaEccezioni
                            if (!string.IsNullOrWhiteSpace(R.Eccezioni))
                            {
                                if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Eccezioni"], R.Eccezioni))
                                    Logger.PrintLC("Added Relation's Exceptions (" + R.Eccezioni + ") to " + scItem.Name, 3, ConfigFile.INFO);
                                else
                                {
                                    //errore = "Error adding Relation Exceptions (" + R.Eccezioni + ") to " + scItem.Name;
                                    errore = "Errore riscontrato aggiungendo " + ConfigFile._REL_NAME["Eccezioni"] + " (" + R.Eccezioni + ") a " + scItem.Name;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                            }
                            #endregion
                            CommitAndSave(trID);
                            OpenTransaction();
                            #region RinominaFisica
                            //***************************************
                            //Rename Campo Padre nella Tabella Figlia #1
                            if (R.CampoFiglio != R.CampoPadre)
                            {
                                //Recuperiamo nuovamente la Tabella Figlio
                                erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                                if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia.ToUpper()))
                                {
                                    //errore = "Relation Ignored: Could not find table " + R.TabellaFiglia + " inside relation ID " + relation.ID;
                                    errore = "Relazione ignorata: Imposibile trovare la tabella " + R.TabellaFiglia + " all'interno della relazione ID " + relation.ID;
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                                //Recuperiamo l'Attributo con il nome Campo Padre (aggiunto con la relazione)
                                SCAPI.ModelObjects erAttributesFigliox = scSession.ModelObjects.Collect(tabellaFiglio, "Attribute");
                                if (!con.RetriveAttribute(ref campoFiglio, erAttributesFigliox, R.CampoPadre.ToUpper()))
                                {
                                    //errore = "Failed Rename: could not find Parent Field " + R.CampoPadre + " inside Child Table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                    errore = "Rinominazione fallita: impossibile trovare Parent Field " + R.CampoPadre + " all'interno della Child Table " + R.TabellaFiglia + " per la relazione ID " + relation.ID;
                                    Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                                else
                                {
                                    //Code 78 - PHYSICAL NAME
                                    if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy"], R.CampoFiglio.ToUpper()))
                                    { 
                                        Logger.PrintLC("Renamed (physical) Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4, ConfigFile.INFO);
                                    }
                                    else
                                    {
                                        //errore = "Failed Rename (phisical): could not find rename with name Child Field(" + R.CampoFiglio + ") to Child Name: " + scItem.ObjectId;
                                        errore = "Rinomina fallita (fisica): impossibile trovare Parent Field " + R.CampoFiglio + " all'interno della Child Table " + scItem.ObjectId + " per la relazione ID " + relation.ID;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return scItem;
                                        continue;
                                    }
                                    //Code 77 #1
                                    // Se il campo era Chiave, proviamo a risettarlo Chiave
                                    if (R.CampoFiglioKey)
                                    {
                                        if (con.AssignToObjModelInt(ref campoFiglio, ConfigFile._ATT_NAME["Chiave"], 0))
                                            Logger.PrintLC("Set child field " + R.CampoFiglio + " as Key.", 4, ConfigFile.INFO);
                                        else
                                        {
                                            Logger.PrintLC("Could not set child field " + R.CampoFiglio + " as Key. Continue.", 4, ConfigFile.INFO);
                                            if (R.History != null)
                                                errore = "\n" + "Impossibile risettare il Campo Figlio " + R.CampoFiglio + " come Chiave";
                                            CommitAndSave(trID);
                                            continue;
                                        }
                                    }
                                }
                                //    CommitAndSave(trID);
                                //    OpenTransaction();
                            }
                            #endregion

                            CommitAndSave(trID);
                            OpenTransaction();
                            
                            #region RinominaLogica
                            //***************************************
                            //Rename Campo Padre nella Tabella Figlia #2
                            if (R.CampoFiglio != R.CampoPadre)
                            {
                                //Recuperiamo nuovamente la Tabella Figlio
                                erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                                if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia.ToUpper()))
                                {
                                    //errore = "Relation Ignored: Could not find table " + R.TabellaFiglia + " inside relation ID " + relation.ID;
                                    errore = "Relazione ignorata: Impossibile trovare la tabella " + R.TabellaFiglia + " nel foglio tabelle";
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                                //Recuperiamo l'Attributo con il nome Campo Padre (aggiunto con la relazione)
                                SCAPI.ModelObjects erAttributesFigliox = scSession.ModelObjects.Collect(tabellaFiglio, "Attribute");
                                if (!con.RetriveAttribute(ref campoFiglio, erAttributesFigliox, R.CampoPadre.ToUpper()))
                                {
                                    //errore = "Failed Rename: could not find Parent Field " + R.CampoPadre + " inside Child Table " + R.TabellaFiglia + " with relation ID " + relation.ID;
                                    errore = "Rinomina (logica) fallita: impossibile trovare Parent Field " + R.CampoPadre + " all'interno della Child Table " + R.TabellaFiglia + " per la relazione ID " + relation.ID;
                                    Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    //return scItem;
                                    continue;
                                }
                                else
                                {
                                    if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy Name"], R.CampoFiglio.ToUpper()))
                                    {
                                        Logger.PrintLC("Renamed (logical) Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4, ConfigFile.INFO);
                                    }
                                    else
                                    {
                                        //errore = "Failed Rename: could not find rename with name Child Field(" + R.CampoPadre + ") to Child Name: " + scItem.ObjectId;
                                        errore = "Rinomina (logica) fallita: impossibile trovare il Child Field(" + R.CampoPadre + ") per la Child Table: " + scItem.ObjectId;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (R.History != null)
                                            errore = "\n" + errore;
                                        R.History += errore;
                                        CommitAndSave(trID);
                                        //return scItem;
                                        continue;
                                    }

                                    //Code 77 #2
                                    // Se il campo era Chiave, proviamo a ri-settarlo Chiave
                                    if (R.CampoFiglioKey)
                                    {
                                        if (con.AssignToObjModelInt(ref campoFiglio, ConfigFile._ATT_NAME["Chiave"], 0))
                                            Logger.PrintLC("Set child field " + R.CampoFiglio + " as Key.", 4, ConfigFile.INFO);
                                        else
                                        {
                                            Logger.PrintLC("Could not set child field " + R.CampoFiglio + " as Key. Continue.", 4, ConfigFile.INFO);
                                            if (R.History != null)
                                                errore = "\n" + "Impossibile resettare il Campo Figlio " + R.CampoFiglio + " come Chiave";
                                            continue;
                                        }
                                    }
                                }

                                //Code 79
                                // Nascondiamo nuovamente gli eventuali Attributi che erano in Hide
                                CommitAndSave(trID);
                                OpenTransaction();
                                erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                                if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia.ToUpper()))
                                {
                                    errore = "Fase controllo campi nascosti: impossibile trovare la tabella " + R.TabellaFiglia + " nel foglio tabelle. Anomalia.";
                                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                    if (R.History != null)
                                        errore = "\n" + errore;
                                    R.History += errore;
                                    CommitAndSave(trID);
                                    continue;
                                }
                                else
                                {
                                    SCAPI.ModelObjects erAttributesFiglioHide = scSession.ModelObjects.Collect(tabellaFiglio, "Attribute");
                                    while (con.RetriveAttribute(ref campoFiglio, erAttributesFiglioHide, R.CampoPadre.ToUpper()))
                                    {
                                        if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy"], R.CampoFiglio.ToUpper()))
                                        {
                                            Logger.PrintLC("Rename Parent-in-child fase: renamed (logical) Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4, ConfigFile.INFO);
                                        }
                                        else
                                        {
                                            //errore = "Failed Rename: could not find rename with name Child Field(" + R.CampoPadre + ") to Child Name: " + scItem.ObjectId;
                                            errore = "Rename Parent-in-child fase: Rinomina (Physical) fallita: impossibile trovare il Child Field(" + R.CampoPadre + ") per la Child Table: " + scItem.ObjectId;
                                            Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                            if (R.History != null)
                                                errore = "\n" + errore;
                                            R.History += errore;
                                            CommitAndSave(trID);
                                            OpenTransaction();
                                            continue;
                                        }
                                        if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy Name"], R.CampoFiglio.ToUpper()))
                                        {
                                            Logger.PrintLC("Rename Parent-in-child fase: renamed (logical) Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4, ConfigFile.INFO);
                                        }
                                        else
                                        {
                                            //errore = "Failed Rename: could not find rename with name Child Field(" + R.CampoPadre + ") to Child Name: " + scItem.ObjectId;
                                            errore = "Rename Parent-in-child fase: Rinomina (logica) fallita: impossibile trovare il Child Field(" + R.CampoPadre + ") per la Child Table: " + scItem.ObjectId;
                                            Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                            if (R.History != null)
                                                errore = "\n" + errore;
                                            R.History += errore;
                                            CommitAndSave(trID);
                                            OpenTransaction();
                                            continue;
                                        }
                                    }
                                }
                            }
                            CommitAndSave(trID);
                            OpenTransaction();
                            #endregion
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(R.History))
                            {
                                //R.History = "\n" + "Relation Ignored: another element of the same relation is wrong";
                                if (R.History != null)
                                    errore = "\n" + "Relazione ignorata: un altro elemento della stessa relazione è errato";
                                R.History += "Relazione ignorata: un altro elemento della stessa relazione è errato";
                                //R.History = "\n" + "Relazione ignorata: un altro elemento della stessa relazione è errato";
                                continue;
                            }
                        }
                    }
                    CommitAndSave(trID);
                    return scItem;
                    #endregion
                    
                }
                catch (Exception exc)
                {
                    Logger.PrintLC("Unexpected error while creating relationship: " + exc.Message, 2, ConfigFile.ERROR);
                    CommitAndSave(trID);
                    return ret;
                }
            }
            return ret;
        }

        public bool RefreshTables(GlobalRelationStrut globalRelation)
        {
            foreach(RelationStrut relationGroup in globalRelation.GlobalRelazioni)
            {
                foreach(var relation in relationGroup.Relazioni)
                {
                    string TabellaPadre = relation.TabellaPadre;
                    string TabellaFiglio = relation.TabellaFiglia;
                    string CampoPadre = relation.CampoPadre;
                    string CampoFiglio = relation.CampoFiglio;


                }
            }
            return true;
        }


        public SCAPI.ModelObject CreateAttributePassOne(AttributeT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            string errore = string.Empty;
            if (string.IsNullOrWhiteSpace(db))
            {
                //errore = "There was no DB associated to " + entity.NomeTabellaLegacy;
                errore = "Non ci sono DB associati a " + entity.NomeTabellaLegacy;
                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                if (entity.History != null)
                    errore = "\n" + errore;
                entity.History += errore;
                entity.Step = 1;
                return ret;
            }
            
            if (erRootObjCol != null)
            {
                OpenTransaction();

                erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");

                VBCon con = new VBCon();
                erEntityObjectPE = null;

                if (string.IsNullOrWhiteSpace(entity.NomeTabellaLegacy))
                {
                    //errore = "'Nome Tabella Legacy' at row " + entity.Row + " not found. Skipping the Attribute.";
                    errore = "'Nome Tabella Legacy' alla riga " + entity.Row + " non trovato. Attributo ignorato.";
                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    if (entity.History != null)
                        errore = "\n" + errore;
                    entity.History += errore;
                    entity.Step = 1;
                    CommitAndSave(trID);
                    return ret = null;
                }

                if (con.RetriveEntity(ref erEntityObjectPE, erObjectCollection, entity.NomeTabellaLegacy.ToUpper()))
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " retrived correctly", 3, ConfigFile.INFO);
                else
                {
                    errore = "Tabella " + entity.NomeTabellaLegacy + " non trovata. Attributo ignorato.";
                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    if (entity.History != null)
                        errore = "\n" + errore;
                    entity.History += errore;
                    entity.Step = 1;
                    CommitAndSave(trID);
                    return ret = null;
                }

                //Area
                if (!string.IsNullOrWhiteSpace(entity.Area))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Area"], entity.Area))
                        Logger.PrintLC("Added Area to " + erEntityObjectPE.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Area to " + erEntityObjectPE.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Area"] + " a " + erEntityObjectPE.Name;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        entity.Step = 1;
                    }
                //Tipologia Tabella
                if (!string.IsNullOrWhiteSpace(entity.TipologiaTabella))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Tipologia Tabella"], entity.TipologiaTabella))
                        Logger.PrintLC("Added Tipologia Tabella to " + erEntityObjectPE.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Tipologia Tabella to " + erEntityObjectPE.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Tipologia Tabella"] + " a " + erEntityObjectPE.Name;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        entity.Step = 1;
                    }
                //Storica
                if (!string.IsNullOrWhiteSpace(entity.Storica))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Storica"], entity.Storica))
                        Logger.PrintLC("Added Storica to " + erEntityObjectPE.Name, 3, ConfigFile.INFO);
                    else
                    {
                        //errore = "Error adding Storica to " + erEntityObjectPE.Name;
                        errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Storica"] + " a " + erEntityObjectPE.Name;
                        Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        entity.Step = 1;
                    }

                erAttributeObjCol = scSession.ModelObjects.Collect(erEntityObjectPE, "Attribute");

                if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                    if (con.RetriveAttribute(ref erAttributeObjectPE, erAttributeObjCol, entity.NomeCampoLegacy.ToUpper()))
                    {
                        //errore = "Attribute entity " + entity.NomeCampoLegacy + " already present.";
                        errore = "Attributo tabella " + entity.NomeCampoLegacy + " già presente.";
                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        entity.Step = 1;
                    }
                    else
                    {
                        erAttributeObjectPE = erAttributeObjCol.Add("Attribute");
                        //Name
                        if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                        {
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Nome Campo Legacy Name"], entity.NomeCampoLegacy.ToUpper()))
                                Logger.PrintLC("Added Nome Campo Legacy to " + erAttributeObjectPE.Name + "'s name.", 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Nome Campo Legacy to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Nome Campo Legacy Name"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }
                            //Physical Name
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Nome Campo Legacy"], entity.NomeCampoLegacy.ToUpper()))
                                Logger.PrintLC("Added Nome Campo Legacy to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Nome Campo Legacy to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Nome Campo Legacy"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }
                        }
                        //Datatype
                        if (!string.IsNullOrWhiteSpace(entity.DataType))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Datatype"], entity.DataType))
                                Logger.PrintLC("Added Datatype to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Datatype to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Datatype"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }
                        //Chiave
                        if (entity.Chiave == 0 || entity.Chiave == 100)
                            if (con.AssignToObjModelInt(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave"], (int)entity.Chiave))
                                Logger.PrintLC("Added Chiave to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Chiave to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Chiave"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }
                        //Mandatory Flag
                        if (entity.MandatoryFlag == 1 || entity.MandatoryFlag == 0)
                            if (con.AssignToObjModelInt(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Mandatory Flag"], (int)entity.MandatoryFlag))
                                Logger.PrintLC("Added Mandatory Flag to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Mandatory Flag to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Mandatory Flag"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }

                        //Dati Sensibili
                        if (!string.IsNullOrWhiteSpace(entity.DatoSensibile))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Dato Sensibile"], entity.DatoSensibile))
                                Logger.PrintLC("Added Dato Sensibile to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Dato Sensibile to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Dato Sensibile"] + " a " + erEntityObjectPE.Name;
                                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 1;
                            }

                    }
                CommitAndSave(trID);
            }
            return erEntityObjectPE;
        }


        public SCAPI.ModelObject CreateAttributePassTwo(AttributeT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            string errore = string.Empty;
            if (string.IsNullOrWhiteSpace(db))
            {
                //errore = "There was no DB associated to " + entity.NomeTabellaLegacy;
                errore = "Non ci sono DB associati a " + entity.NomeTabellaLegacy;
                Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                if (entity.History != null)
                    errore = "\n" + errore;
                entity.History += errore;
                entity.Step = 2;
                return ret;
            }
            if (erRootObjCol != null)
            {
                OpenTransaction();
                erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                VBCon con = new VBCon();
                erEntityObjectPE = null;
                if (string.IsNullOrWhiteSpace(entity.NomeTabellaLegacy))
                {
                    //errore = "'Nome Tabella Legacy' at row " + entity.Row + " not found. Skipping the Attribute.";
                    errore = "'Nome Tabella Legacy' alla riga " + entity.Row + " non trovato. Attributo ignorato.";
                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    if (entity.History != null)
                        errore = "\n" + errore;
                    entity.History += errore;
                    entity.Step = 2;
                    CommitAndSave(trID);
                    return ret = null;
                }
                if (con.RetriveEntity(ref erEntityObjectPE, erObjectCollection, entity.NomeTabellaLegacy.ToUpper()))
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " retrived correctly", 3, ConfigFile.INFO);
                else
                {
                    errore = "Tabella " + entity.NomeTabellaLegacy + " non trovata. Attributo ignorato.";
                    Logger.PrintLC(errore, 3, ConfigFile.ERROR);
                    if (!(entity.History.Contains(errore)))
                    {
                        if (entity.History != null)
                            errore = "\n" + errore;
                        entity.History += errore;
                        entity.Step = 2;
                    }
                    CommitAndSave(trID);
                    return ret = null;
                }
                erAttributeObjCol = scSession.ModelObjects.Collect(erEntityObjectPE, "Attribute");

                //Code 77
                if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                    if (con.RetriveAttribute(ref erAttributeObjectPE, erAttributeObjCol, entity.NomeCampoLegacy.ToUpper()))
                    {
                        //valorizzo Ordine
                        string Ordine = string.Empty;
                        if (con.RetrieveFromObjModel(erAttributeObjectPE, ConfigFile._ATT_NAME["Ordine"], ref Ordine))
                        {
                            entity.Ordine = Ordine;
                        }
                            ////Verifico che la chiave non sia stata alterata dalla relazione, eventualmente la ripristino
                            string isKey = string.Empty;
                        // ogni colonna deve avere un valore chiave
                        if (con.RetrieveFromObjModel(erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave"], ref isKey))
                        {
                            //verifico che il tipo di chiave non sia cambiato rispetto a quella che ho settato dalla collezione
                            if (isKey != entity.Chiave.ToString())
                            {
                                //escludendo i casi in cui non era chiave e lo è diventata a seguito della relazione
                                if (entity.Chiave == 0 || entity.Chiave == 100)
                                    //riassegno il corretto valore chiave all'attributo
                                    if (con.AssignToObjModelInt(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave"], (int)entity.Chiave))
                                    {
                                        Logger.PrintLC("La chiave alterata dalla relazione è stata ripristinata per il campo " + entity.NomeCampoLegacy + " della tabella " + entity.NomeTabellaLegacy, 4, ConfigFile.INFO);
                                    }
                                    else
                                    {
                                        errore = "La chiave alterata dalla relazione NON è stata ripristinata per il campo " + entity.NomeCampoLegacy + " della tabella " + entity.NomeTabellaLegacy;
                                        Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                        if (entity.History != null)
                                            errore = "\n" + errore;
                                        entity.History += errore;
                                    }
                            }
                        }
                        //Definizione Campo
                        if (!string.IsNullOrWhiteSpace(entity.DefinizioneCampo))
                        {
                            //Logger.PrintLC("Attribute entity " + entity.NomeCampoLegacy + " already present.", 3, ConfigFile.WARNING);
                            //Definizione Campo (Comment)
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Definizione Campo"], entity.DefinizioneCampo))
                                Logger.PrintLC("Added Definizione Campo (Comment) to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Definizione Campo (Comment) to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Definizione Campo"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                            //Definizione Campo (Definition)
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Definizione Campo Def"], entity.DefinizioneCampo))
                                Logger.PrintLC("Added Definizione Campo (Definition) to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Definizione Campo (Definition) to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Definizione Campo Def"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                        }
                        //Unique
                        if (!string.IsNullOrWhiteSpace(entity.Unique))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Unique"], entity.Unique))
                                Logger.PrintLC("Added Unique to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Unique to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Unique"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                        //Chiave logica
                        if (!string.IsNullOrWhiteSpace(entity.ChiaveLogica))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave Logica"], entity.ChiaveLogica))
                                Logger.PrintLC("Added Chiave Logica to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Chiave Logica to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Chiave Logica"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                        //Dominio
                        if (!string.IsNullOrWhiteSpace(entity.Dominio))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Dominio"], entity.Dominio))
                                Logger.PrintLC("Added Dominio to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Dominio to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Dominio"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                        //Provenienza Dominio
                        if (!string.IsNullOrWhiteSpace(entity.ProvenienzaDominio))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Provenienza Dominio"], entity.ProvenienzaDominio))
                                Logger.PrintLC("Added Provenienza Dominio to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Provenienza Dominio to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Provenienza Dominio"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                        //Note
                        if (!string.IsNullOrWhiteSpace(entity.Note))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Note"], entity.Note))
                                Logger.PrintLC("Added Note to " + erAttributeObjectPE.Name, 4, ConfigFile.INFO);
                            else
                            {
                                //errore = "Error adding Note to " + erAttributeObjectPE.Name;
                                errore = "Errore riscontrato aggiungendo " + ConfigFile._ATT_NAME["Note"] + " a " + erAttributeObjectPE.Name;
                                Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                                if (entity.History != null)
                                    errore = "\n" + errore;
                                entity.History += errore;
                                entity.Step = 2;
                            }
                    }
                    else
                    {
                        //ExcelOps.XLSXWriteErrorInCell()
                        {
                            //errore = "Unexpected Error: searching for " + entity.NomeCampoLegacy + " finding none.";
                            errore = "Errore inaspettato: cercando " + entity.NomeCampoLegacy + " non è stato trovato nulla.";
                            Logger.PrintLC(errore, 4, ConfigFile.ERROR);
                            if (entity.History != null)
                                errore = "\n" + errore;
                            entity.History += errore;
                            entity.Step = 2;
                        }
                    }
                CommitAndSave(trID);
            }
            return erEntityObjectPE;
        }


        /// <summary>
        /// Commits and saves the state of 'id'
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool CommitAndSave(object id)
        {
            if (!CommitTransaction(id))
                return false;
            if (!SavePersistence())
                return false;
            return true;
        }

        /// <summary>
        /// Commits the Transaction of 'id'
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool CommitTransaction(object id)
        {
            try
            {
                if (!scSession.CommitTransaction(id))
                {
                    Logger.PrintLC("Could not Commit for ID: " + id, 3, ConfigFile.WARNING);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Committed successfully: " + id, 3, ConfigFile.INFO);
                    lastIdCommitted = id;
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Can NOT Commit.", 3, ConfigFile.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Save the Persistence on 'this' Connection
        /// </summary>
        /// <returns></returns>
        public bool SavePersistence()
        {
            try
            {
                FileOps.RemoveAttributes(fileERwin);
                if (!scPersistenceUnit.Save())
                {
                    Logger.PrintLC("Could NOT save Persistence: " + scPersistenceUnit.ObjectId, 3, ConfigFile.ERROR);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Persistence SAVED: " + scPersistenceUnit.ObjectId, 3, ConfigFile.INFO);
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Persistence NOT saved.", 3, ConfigFile.ERROR);
                return false;
            }
        }

        public static bool AssignToObjModel(ref SCAPI.ModelObject model, string property, string value)
        {
            VBCon VBcon = new VBCon();
            if (VBcon.AssignToObjModel(ref model, property, value))
                return true;
            else
                return false;
        }

        public static bool AssignToObjModelInt(ref SCAPI.ModelObject model, string property, int value)
        {
            VBCon VBcon = new VBCon();
            if (VBcon.AssignToObjModelInt(ref model, property, value))
                return true;
            else
                return false;
        }
    }
}
