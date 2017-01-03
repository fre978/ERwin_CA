﻿using ERwin_CA.T;
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

        public bool openModelConnection(string ERw)
        {
            if (ERw == null)
                return false;
            if (!File.Exists(ERw))
            {
                Logger.PrintLC("Could not find file: " + ERw, 2);
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
                Logger.PrintLC("Connection opened.",2);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Connection opening error: " + exp.Message, 2);
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
                Logger.PrintLC("Session closed successfully.", 2);
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Could not close the Session.", 2);
            }
        }



        public object OpenTransaction()
        {
            if (scSession != null)
                try
                {
                    trID = scSession.BeginTransaction();
                    Logger.PrintLC("Transaction began successfully.", 3);
                    return trID;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Starting Transaction error: " + exp.Message, 3);
                    return -1;
                }
            else
                Logger.PrintLC("Starting Transaction error: missing SESSION.", 3);
            return -1;
        }

        public bool SetRootObject()
        {
            if (scSession != null)
                try
                {
                    erRootObj = scSession.ModelObjects.Root;
                    Logger.PrintLC("Root has been successful.", 3);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Setting Root's Session error: " + exp.Message, 3);
                    return false;
                }
            else
                Logger.PrintLC("Could not determine Root because Session is missing.", 3);
            return false;
        }

        public bool SetRootCollection()
        {
            if (scSession != null)
                try
                {
                    erRootObjCol = scSession.ModelObjects.Collect(erRootObj);
                    Logger.PrintLC("Root Collection has been successful.", 3);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not get the Root Collection: " + exp.Message, 3);
                    return false;
                }
            else
            {
                Logger.PrintLC("Could not get Root Collection because Session is missing.", 3);
                return false;
            }
        }

        public SCAPI.ModelObject CreateEntity (EntityT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            if (string.IsNullOrWhiteSpace(db))
            {
                Logger.PrintLC("There was no DB associated to " + entity.TableName, 3);
                return ret;
            }
            if (!((entity.FlagBFD == "S") || (entity.FlagBFD == "N")))
            {
                Logger.PrintLC("Property FlagBFD of " + entity.TableName + " is not valid. Table will be skipped.", 3);
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
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Tabella"], entity.TableName))
                        Logger.PrintLC("Added Table Physical Name (" + entity.TableName + ") to " + scItem.ObjectId, 3);
                    else
                    {
                        Logger.PrintLC("Error adding Table Physical Name (" + entity.TableName + ") to " + scItem.ObjectId, 3);
                        CommitAndSave(trID);
                        return scItem;
                    }
                    if (con.AssignToObjModel(ref scItem, "Name", entity.TableName))
                        Logger.PrintLC("Added Table Name to " + scItem.Name, 3);
                    else
                    {
                        Logger.PrintLC("Error adding Table Name to " + scItem.Name, 3);
                        CommitAndSave(trID);
                        return scItem;
                    }
                }

                //SSA
                if (!string.IsNullOrWhiteSpace(entity.SSA))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["SSA"], entity.SSA))
                        Logger.PrintLC("Added SSA to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding SSA to " + scItem.Name, 3);
                
                //Table Description
                if (!string.IsNullOrWhiteSpace(entity.TableDescr))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Descrizione Tabella"], entity.TableDescr))
                        Logger.PrintLC("Added Table Description to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Table Description to " + scItem.Name, 3);
                if (!string.IsNullOrWhiteSpace(entity.TableDescr))
                    if (con.AssignToObjModel(ref scItem, "Definition", entity.TableDescr))
                        Logger.PrintLC("Added Table Definition to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Table Definition to " + scItem.Name, 3);


                //Info Type
                if (!string.IsNullOrWhiteSpace(entity.InfoType))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Tipologia Informazione"], entity.InfoType))
                        Logger.PrintLC("Added Information Type to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Information Type to " + scItem.Name, 3);

                //Table Limit
                if (!string.IsNullOrWhiteSpace(entity.TableLimit))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Perimetro Tabella"], entity.TableLimit))
                        Logger.PrintLC("Added Table Limit to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Table Limit to " + scItem.Name, 3);

                //Table Granularity
                if (!string.IsNullOrWhiteSpace(entity.TableGranularity))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Granularità Tabella"], entity.TableGranularity))
                        Logger.PrintLC("Added Table Granularity to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Table Granularity to " + scItem.Name, 3);

                //Flag BFD
                if (!string.IsNullOrWhiteSpace(entity.FlagBFD))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Flag BFD"], entity.FlagBFD))
                        Logger.PrintLC("Added Flag BFD to " + scItem.Name, 3);
                    else
                        Logger.PrintLC("Error adding Flag BFD to " + scItem.Name, 3);

                //##################################################
                //## Controllo esistenza DB ed eventuale aggiunta ##
                //** Qui vanno aggiunti eventuali altri DB da     **
                //** prendere in considerazione oltre a DB2 e ORACLE
                if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                {
                    if (entity.DB == "DB2")
                    {
                        if (!DatabaseN.Contains(entity.DatabaseName))
                        {
                            scDB = erRootObjCol.Add("DB2_Database");
                            if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome Database"], entity.DatabaseName))
                                Logger.PrintLC("Added Database Name to " + scDB.Name, 3);
                            else
                                Logger.PrintLC("Error adding Database Name to " + scDB.Name, 3);

                            if (!string.IsNullOrWhiteSpace(entity.HostName))
                                if (con.AssignToObjModel(ref scDB, ConfigFile._TAB_NAME["Nome host"], entity.HostName))
                                    Logger.PrintLC("Added Host Name to " + scDB.Name, 3);
                                else
                                    Logger.PrintLC("Error adding Host Name to " + scDB.Name, 3);
                            DatabaseN.Add(entity.DatabaseName);
                        }
                    }
                }
                if (entity.DB == "ORACLE")
                {
                    if (!string.IsNullOrWhiteSpace(entity.HostName))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome host Oracle"], entity.HostName))
                            Logger.PrintLC("Added Host Oracle Name to " + scItem.Name, 3);
                        else
                            Logger.PrintLC("Error adding Host Oracle Name to " + scItem.Name, 3);
                    }
                    if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Database Oracle"], entity.DatabaseName))
                            Logger.PrintLC("Added Oracle Database Name to " + scItem.Name, 3);
                        else
                            Logger.PrintLC("Error adding Oracle Database Name to " + scItem.Name, 3);
                    }
                }

                //##################################################

                //##################################################
                //## Controllo esistenza SCHEMA ed eventuale aggiunta ##
                //** Aggiungere qui eventuali altri casi DB oltre a DB2/ORACLE
                if (entity.DB == "DB2")
                {
                    if (!string.IsNullOrWhiteSpace(entity.Schema))
                    {
                        if (!SchemaN.Contains(entity.Schema))
                        {
                            scSchema = erRootObjCol.Add("Schema");
                            if (con.AssignToObjModel(ref scSchema, "Name", entity.Schema))
                                Logger.PrintLC("Created Schema Name to " + scSchema.Name, 3);
                            else
                                Logger.PrintLC("Error creating Schema Name to " + scSchema.Name, 3);
                            SchemaN.Add(entity.Schema);
                        }
                        //Schema
                        if (!string.IsNullOrWhiteSpace(entity.Schema))
                            if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Schema"], entity.Schema))
                                Logger.PrintLC("Added Schema to " + scItem.Name, 3);
                            else
                                Logger.PrintLC("Error adding Schema to " + scItem.Name, 3);
                    }
                }
                if(entity.DB == "ORACLE")
                {
                    if (!string.IsNullOrWhiteSpace(entity.Schema))
                    {
                        if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Schema Oracle"], entity.Schema))
                            Logger.PrintLC("Added Host Oracle Schema to " + scItem.Name, 3);
                        else
                            Logger.PrintLC("Error adding Host Oracle Schema to " + scItem.Name, 3);
                    }
                }
                //##################################################
                CommitAndSave(trID);
            }
            return scItem;
        }

        public SCAPI.ModelObject CreateRelation(RelationStrut relation, string db)
        {
            SCAPI.ModelObject ret = null;
            if (string.IsNullOrWhiteSpace(db))
            {
                Logger.PrintLC("There was no DB associated to " + relation.ID, 3);
                return ret;
            }
            if (erRootObjCol != null)
            {
                try
                {
                    OpenTransaction();
                    //collezione completa delle entity
                    erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                    int countRelazioni = relation.Relazioni.Count;
                    VBCon con = new VBCon();
                    foreach (var R in relation.Relazioni)
                    {
                        int countKey = 0;
                        SCAPI.ModelObject tabellaPadre = null;
                        SCAPI.ModelObject tabellaFiglio = null;
                        SCAPI.ModelObject campoPadre = null;
                        SCAPI.ModelObject campoFiglio = null;
                        
                        #region verificheErwin
                        //cerchiamo la tabella padre
                        if (!con.RetriveEntity(ref tabellaPadre, erObjectCollection, R.TabellaPadre))
                        {
                            Logger.PrintLC("Relation ignored: Could not find table " + R.TabellaPadre + " inside relation ID " + relation.ID, 3);
                            CommitAndSave(trID);
                            return ret = null;
                        }
                        //cerchiamo la tabella figlia
                        if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia))
                        {
                            Logger.PrintLC("Relation Ignored: Could not find table " + R.TabellaFiglia + " inside relation ID " + relation.ID, 3);
                            CommitAndSave(trID);
                            return ret = null;
                        }
                        //esistenza campo padre
                        SCAPI.ModelObjects erAttributesPadre = scSession.ModelObjects.Collect(tabellaPadre, "Attribute");
                        if (!con.RetriveAttribute(ref campoPadre, erAttributesPadre, R.CampoPadre))
                        {
                            Logger.PrintLC("Relation Ignored: Could not find field " + R.CampoPadre + " inside table " + R.TabellaPadre + " with relation ID " + relation.ID, 3);
                            CommitAndSave(trID);
                            return ret = null;
                        }
                        else
                        {
                            //key
                            string isKey = null;
                            foreach (SCAPI.ModelObject attributo in erAttributesPadre)
                            {
                                if (!con.RetrieveFromObjModel(attributo, "Type", ref isKey))
                                {
                                    Logger.PrintLC("Relation Ignored: Could not find attribute Type of field " + R.CampoPadre + " inside table " + R.TabellaPadre + " with relation ID " + relation.ID, 4);
                                    CommitAndSave(trID);
                                    return ret = null;
                                }
                                else
                                {
                                    if (isKey == "0")
                                    {
                                        countKey += 1;
                                    }
                                }
                            }
                            if (countKey != countRelazioni)
                            {
                                Logger.PrintLC("Unmatching PK numbers in table " + R.TabellaPadre + " with relation ID " + relation.ID, 3);
                                CommitAndSave(trID);
                                return ret = null;
                            }
                        }
                        
                        //esistenza campo figlio
                        SCAPI.ModelObjects erAttributesFiglio = scSession.ModelObjects.Collect(tabellaFiglio, "Attribute");
                        if (!con.RetriveAttribute(ref campoFiglio, erAttributesFiglio, R.CampoFiglio))
                        {
                            Logger.PrintLC("Could not find Child Field " + R.CampoFiglio + " inside Child Table " + R.TabellaFiglia + " with relation ID " + relation.ID, 3);
                            CommitAndSave(trID);
                            return ret = null;
                        }
                        else
                        {
                            // if rel=identificativa is key
                            if (R.Identificativa == 2)
                            {
                                string isKey = null;
                                if (!con.RetrieveFromObjModel(campoFiglio, "Type", ref isKey))
                                {
                                    Logger.PrintLC("Relation Ignored: Could not find attribute Type of field " + R.CampoFiglio + " inside child table " + R.TabellaFiglia + " with relation ID " + relation.ID, 4);
                                    CommitAndSave(trID);
                                    return ret = null;
                                }
                                else
                                {
                                    if (isKey != "0")
                                    {
                                        Logger.PrintLC("Relation Ignored: " + R.CampoFiglio + "expected Key inside child table " + R.TabellaFiglia + " with relation ID " + relation.ID, 4);
                                        CommitAndSave(trID);
                                        return ret = null;
                                    }
                                }
                            }
                        }
                        #endregion

                        #region creazionerelazioni;
                        //creare relazione su erwin
                        SetRootObject();
                        SetRootCollection();
                        scItem = erRootObjCol.Add("Relationship");
                        //CommitAndSave(trID);
                        
                        if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Identificativo relazione"],R.IdentificativoRelazione))
                            Logger.PrintLC("Added Relation Id (" + R.IdentificativoRelazione + ") to " + scItem.ObjectId, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Id (" + R.IdentificativoRelazione + ") to " + scItem.ObjectId, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }
                        
                        if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tabella Padre"], R.TabellaPadre))
                            Logger.PrintLC("Added Relation Parent Table (" + R.TabellaPadre + ") to " + scItem.Name, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Parent Table (" + R.TabellaPadre + ") to " + scItem.Name, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }
                        
                        if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tabella Figlia"], R.TabellaFiglia))
                            Logger.PrintLC("Added Relation Child Table (" + R.TabellaFiglia + ") to " + scItem.Name, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Child Table (" + R.TabellaFiglia + ") to " + scItem.Name, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }

                        if (con.AssignToObjModelInt(ref scItem, ConfigFile._REL_NAME["Identificativa"], (int)R.Identificativa))
                            Logger.PrintLC("Added Relation Identifiable (" + R.Identificativa + ") to " + scItem.Name, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Identifiable (" + R.Identificativa + ") to " + scItem.Name, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }

                        //CommitAndSave(trID);
                        int myInt = (R.Cardinalita == null) ? 0 : (int)R.Cardinalita;
                        if (con.AssignToObjModelInt(ref scItem, ConfigFile._REL_NAME["Cardinalita"],myInt ))
                            Logger.PrintLC("Added Relation Cardinality (" + R.Cardinalita + ") to " + scItem.Name, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Cardinality (" + R.Cardinalita + ") to " + scItem.Name, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }
                        myInt = (R.TipoRelazione == true) ? 1 : 0;
                        if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Tipo Relazione"], "False"))
                            Logger.PrintLC("Added Relation Type (" + R.TipoRelazione + ") to " + scItem.Name, 3);
                        else
                        {
                            Logger.PrintLC("Error adding Relation Type (" + R.TipoRelazione + ") to " + scItem.Name, 3);
                            CommitAndSave(trID);
                            return scItem;
                        }
                        if (!string.IsNullOrWhiteSpace(R.Note))
                        {
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Note"], R.Note))
                                Logger.PrintLC("Added Relation Note (" + R.Note + ") to " + scItem.Name, 3);
                            else
                            {
                                Logger.PrintLC("Error adding Relation Note (" + R.Note + ") to " + scItem.Name, 3);
                                CommitAndSave(trID);
                                return scItem;
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(R.Eccezioni))
                        {
                            if (con.AssignToObjModel(ref scItem, ConfigFile._REL_NAME["Eccezioni"], R.Eccezioni))
                                Logger.PrintLC("Added Relation Exceptions (" + R.Eccezioni + ") to " + scItem.Name, 3);
                            else
                            {
                                Logger.PrintLC("Error adding Relation Exceptions (" + R.Eccezioni + ") to " + scItem.Name, 3);
                                CommitAndSave(trID);
                                return scItem;
                            }
                        }
                        #endregion
                        CommitAndSave(trID); 

                        //***************************************
                        //Rename Campo Padre nella Tabella Figlia
                        if (R.CampoFiglio != R.CampoPadre)
                        {
                            OpenTransaction();
                            //Recuperiamo nuovamente la Tabella Figlio
                            erObjectCollection = scSession.ModelObjects.Collect(scSession.ModelObjects.Root, "Entity");
                            if (!con.RetriveEntity(ref tabellaFiglio, erObjectCollection, R.TabellaFiglia))
                            {
                                Logger.PrintLC("Relation Ignored: Could not find table " + R.TabellaFiglia + " inside relation ID " + relation.ID, 3);
                                CommitAndSave(trID);
                                return ret = null;
                            } 
                            //Recuperiamo l'Attributo con il nome Campo Padre (aggiunto con la relazione)
                            if (!con.RetriveAttribute(ref campoFiglio, erAttributesFiglio, R.CampoPadre))
                            {
                                Logger.PrintLC("Failed Rename: could not find Parent Field " + R.CampoPadre + " inside Child Table " + R.TabellaFiglia + " with relation ID " + relation.ID, 4);
                                CommitAndSave(trID);
                                return ret = null;
                            }
                            else
                            {
                                //if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy"], R.CampoFiglio))
                                //    Logger.PrintLC("Renamed (Physical) Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4);
                                //else
                                //{
                                //    Logger.PrintLC("Failed Rename (Physical): could not find rename Child Field(" + R.CampoPadre + ") to Child Name: " + scItem.ObjectId, 4);
                                //    CommitAndSave(trID);
                                //    return ret = null;
                                //}
                                if (con.AssignToObjModel(ref campoFiglio, ConfigFile._ATT_NAME["Nome Campo Legacy Name"], R.CampoFiglio))
                                    Logger.PrintLC("Renamed Child Field with name (" + R.CampoPadre + ") to Child Field Name: " + R.CampoFiglio, 4);
                                else
                                {
                                    Logger.PrintLC("Failed Rename: could not find rename Child Field(" + R.CampoPadre + ") to Child Name: " + scItem.ObjectId, 4);
                                    CommitAndSave(trID);
                                    return ret = null;
                                }
                                CommitAndSave(trID);
                            }
                        }
                    }

                    //CommitAndSave(trID);
                    return ret;
                }
                catch (Exception exc)
                {
                    CommitAndSave(trID);
                    return ret;
                }
            }
            return ret;
        }

        public SCAPI.ModelObject CreateAttributePassOne(AttributeT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            if (string.IsNullOrWhiteSpace(db))
            {
                Logger.PrintLC("There was no DB associated to " + entity.NomeTabellaLegacy, 3);
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
                    Logger.PrintLC("'Nome Tabella Legacy' at row " + entity.Row + " not found. Skipping the Attribute.", 3);
                    CommitAndSave(trID);
                    return ret = null;
                }

                if (con.RetriveEntity(ref erEntityObjectPE, erObjectCollection, entity.NomeTabellaLegacy))
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " retrived correctly", 3);
                else
                {
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " not found. Skipping the Attribute.", 3);
                    CommitAndSave(trID);
                    return ret = null;
                }

                //Area
                if (!string.IsNullOrWhiteSpace(entity.Area))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Area"], entity.Area))
                        Logger.PrintLC("Added Area to " + erEntityObjectPE.Name, 3);
                    else
                        Logger.PrintLC("Error adding Area to " + erEntityObjectPE.Name, 3);
                //Tipologia Tabella
                if (!string.IsNullOrWhiteSpace(entity.TipologiaTabella))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Tipologia Tabella"], entity.TipologiaTabella))
                        Logger.PrintLC("Added Tipologia Tabella to " + erEntityObjectPE.Name, 3);
                    else
                        Logger.PrintLC("Error adding Tipologia Tabella to " + erEntityObjectPE.Name, 3);
                //Storica
                if (!string.IsNullOrWhiteSpace(entity.Storica))
                    if (con.AssignToObjModel(ref erEntityObjectPE, ConfigFile._ATT_NAME["Storica"], entity.Storica))
                        Logger.PrintLC("Added Storica to " + erEntityObjectPE.Name, 3);
                    else
                        Logger.PrintLC("Error adding Storica to " + erEntityObjectPE.Name, 3);

                erAttributeObjCol = scSession.ModelObjects.Collect(erEntityObjectPE, "Attribute");

                if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                    if (con.RetriveAttribute(ref erAttributeObjectPE, erAttributeObjCol, entity.NomeCampoLegacy))
                        Logger.PrintLC("Attribute entity " + entity.NomeCampoLegacy + " already present.", 3);
                    else
                    {
                        erAttributeObjectPE = erAttributeObjCol.Add("Attribute");
                        //Name
                        if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                        {
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Nome Campo Legacy Name"], entity.NomeCampoLegacy))
                                Logger.PrintLC("Added Nome Campo Legacy to " + erAttributeObjectPE.Name + "'s name.", 4);
                            else
                                Logger.PrintLC("Error adding Nome Campo Legacy to " + erAttributeObjectPE.Name, 4);
                            //Physical Name
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Nome Campo Legacy"], entity.NomeCampoLegacy))
                                Logger.PrintLC("Added Nome Campo Legacy to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Nome Campo Legacy to " + erAttributeObjectPE.Name, 4);
                        }
                        //Datatype
                        if(!string.IsNullOrWhiteSpace(entity.DataType))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Datatype"], entity.DataType))
                                Logger.PrintLC("Added Datatype to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Datatype to " + erAttributeObjectPE.Name, 4);
                        //Chiave
                        if(entity.Chiave == 0 || entity.Chiave == 100)
                            if (con.AssignToObjModelInt(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave"], (int)entity.Chiave))
                                Logger.PrintLC("Added Chiave to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Chiave to " + erAttributeObjectPE.Name, 4);
                        //Mandatory Flag
                        if (entity.MandatoryFlag == 1 || entity.MandatoryFlag == 0)
                            if (con.AssignToObjModelInt(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Mandatory Flag"], (int)entity.MandatoryFlag))
                                Logger.PrintLC("Added Mandatory Flag to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Mandatory Flag to " + erAttributeObjectPE.Name, 4);

                    }
                CommitAndSave(trID);
            }
            return erEntityObjectPE;
        }


        public SCAPI.ModelObject CreateAttributePassTwo(AttributeT entity, string db)
        {
            SCAPI.ModelObject ret = null;
            if (string.IsNullOrWhiteSpace(db))
            {
                Logger.PrintLC("There was no DB associated to " + entity.NomeTabellaLegacy, 3);
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
                    Logger.PrintLC("'Nome Tabella Legacy' at row " + entity.Row + " not found. Skipping the Attribute.", 3);
                    CommitAndSave(trID);
                    return ret = null;
                }
                if (con.RetriveEntity(ref erEntityObjectPE, erObjectCollection, entity.NomeTabellaLegacy))
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " retrived correctly", 3);
                else
                {
                    Logger.PrintLC("Table entity " + entity.NomeTabellaLegacy + " not found. Skipping the Attribute.", 3);
                    CommitAndSave(trID);
                    return ret = null;
                }
                erAttributeObjCol = scSession.ModelObjects.Collect(erEntityObjectPE, "Attribute");

                if (!string.IsNullOrWhiteSpace(entity.NomeCampoLegacy))
                    if (con.RetriveAttribute(ref erAttributeObjectPE, erAttributeObjCol, entity.NomeCampoLegacy))
                    {
                        //Definizione Campo
                        if (!string.IsNullOrWhiteSpace(entity.DefinizioneCampo))
                        {
                            Logger.PrintLC("Attribute entity " + entity.NomeCampoLegacy + " already present.", 3);
                            //Definizione Campo (Comment)
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Definizione Campo"], entity.DefinizioneCampo))
                                Logger.PrintLC("Added Definizione Campo (Comment) to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Definizione Campo (Comment) to " + erAttributeObjectPE.Name, 4);
                            //Definizione Campo (Definition)
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Definizione Campo Def"], entity.DefinizioneCampo))
                                Logger.PrintLC("Added Definizione Campo (Definition) to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Definizione Campo (Definition) to " + erAttributeObjectPE.Name, 4);
                        }
                        //Unique
                        if (!string.IsNullOrWhiteSpace(entity.Unique))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Unique"], entity.Unique))
                                Logger.PrintLC("Added Unique to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Unique to " + erAttributeObjectPE.Name, 4);
                        //Chiave logica
                        if (!string.IsNullOrWhiteSpace(entity.ChiaveLogica))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Chiave Logica"], entity.ChiaveLogica))
                                Logger.PrintLC("Added Chiave Logica to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Chiave Logica to " + erAttributeObjectPE.Name, 4);
                        //Dominio
                        if (!string.IsNullOrWhiteSpace(entity.Dominio))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Dominio"], entity.Dominio))
                                Logger.PrintLC("Added Dominio to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Dominio to " + erAttributeObjectPE.Name, 4);
                        //Provenienza Dominio
                        if (!string.IsNullOrWhiteSpace(entity.ProvenienzaDominio))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Provenienza Dominio"], entity.ProvenienzaDominio))
                                Logger.PrintLC("Added Provenienza Dominio to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Provenienza Dominio to " + erAttributeObjectPE.Name, 4);
                        //Note
                        if (!string.IsNullOrWhiteSpace(entity.Note))
                            if (con.AssignToObjModel(ref erAttributeObjectPE, ConfigFile._ATT_NAME["Note"], entity.Note))
                                Logger.PrintLC("Added Note to " + erAttributeObjectPE.Name, 4);
                            else
                                Logger.PrintLC("Error adding Note to " + erAttributeObjectPE.Name, 4);
                    }
                    else
                    {
                        //ExcelOps.XLSXWriteErrorInCell()
                        Logger.PrintLC("Unexpected Error: searching for " + entity.NomeCampoLegacy + " finding none." , 4);
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
                    Logger.PrintLC("Could not Commit for ID: " + id, 3);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Committed successfully: " + id, 3);
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Could NOT Commit.", 3);
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
                    Logger.PrintLC("Could NOT save Persistence: " + scPersistenceUnit.ObjectId, 3);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Persistence SAVED: " + scPersistenceUnit.ObjectId, 3);
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Persistence NOT saved.", 3);
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
