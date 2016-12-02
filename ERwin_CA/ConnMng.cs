using ERwin_CA.T;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VBClassLibrary;

namespace ERwin_CA
{
    class ConnMng
    {
        //Dim scPersistanceUnit As SCAPI.PersistenceUnit    'Persistance unit object
        //Dim scERwin As SCAPI.Application    'SCAPI application object
        //Dim scSession As SCAPI.Session            'SCAPI session object
        //Dim erColumn As SCAPI.ModelObjects    'collezione delle colonne
        //Dim erTable As SCAPI.ModelObjects    'collezione delle tabella
        //Dim scItem As SCAPI.ModelObject            'A single SCAPI object

        public SCAPI.Application scERwin;
        public SCAPI.PersistenceUnit scPersistenceUnit = null;
        public SCAPI.Session scSession;
        public SCAPI.ModelObject erRootObj { get; set; }
        public SCAPI.ModelObjects erRootObjCol { get; set; }
        public SCAPI.ModelObjects erColumn { get; set; }
        public SCAPI.ModelObjects erTable { get; set; }
        public SCAPI.ModelObject scItem;
        public object trID { get; set; }

        public bool openModelConnection(string ERw)
        {
            if (ERw == null)
                return false;
            try
            {
                scERwin = new SCAPI.Application();
                scPersistenceUnit = scERwin.PersistenceUnits.Add(ERw, "RDO=No");

                scSession = scERwin.Sessions.Add();
                scSession.Open(scPersistenceUnit);
                Logger.PrintLC("Connection opened.");
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Connection opening error: " + exp.Message);
                return false;
            }
        }

        public object openTransaction()
        {
            if (scSession != null)
                try
                {
                    trID = scSession.BeginTransaction();
                    Logger.PrintLC("Transaction began successfully.");
                    return trID;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Starting Transaction error: " + exp.Message);
                    return -1;
                }
            else
                Logger.PrintLC("Starting Transaction error: missing SESSION.");
            return -1;
        }
        public bool SetRootObject()
        {
            if (scSession != null)
                try
                {
                    erRootObj = scSession.ModelObjects.Root;
                    Logger.PrintLC("Root has been successful.");
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Setting Root's Session error: " + exp.Message);
                    return false;
                }
            else
                Logger.PrintLC("Could not determine Root because Session is missing.");
            return false;
        }
        public bool SetRootCollection()
        {
            if (scSession != null)
                try
                {
                    erRootObjCol = scSession.ModelObjects.Collect(erRootObj);
                    Logger.PrintLC("Root Collection has been successful.");
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not get the Root Collection: " + exp.Message);
                    return false;
                }
            else
            {
                Logger.PrintLC("Could not get Root Collection because Session is missing.");
                return false;
            }
        }

        public SCAPI.ModelObject CreateEntity (EntityT entity)
        {
            if (erRootObjCol != null)
            {
                scItem = erRootObjCol.Add("Entity");
                VBCon con = new VBCon();
                //Nome tabella
                if (!string.IsNullOrWhiteSpace(entity.TableName))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Tabella"], entity.TableName))
                        Logger.PrintLC("Added Physical Name to " + scItem.ObjectId);
                    else
                    {
                        Logger.PrintLC("Error adding Physical Name to " + scItem.ObjectId);
                        return scItem;
                    }
                CommitTransaction(trID);

                scSession.CommitTransaction(trID);
                //scPersistenceUnit.Save();
                //scSession.Close();

                //SSA
                if (!string.IsNullOrWhiteSpace(entity.SSA))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["SSA"], entity.SSA))
                        Logger.PrintLC("Added SSA to " + scItem.Name);
                    else
                        Logger.PrintLC("Error adding SSA to " + scItem.Name);
                CommitTransaction(trID);
                //Nome Host
                if (!string.IsNullOrWhiteSpace(entity.HostName))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Host"], entity.HostName))
                        Logger.PrintLC("Added Host Name to " + scItem.Name);
                    else
                        Logger.PrintLC("Error adding Host Name to " + scItem.Name);
                CommitTransaction(trID);
                //Nome Database
                if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Nome Database"], entity.DatabaseName))
                        Logger.PrintLC("Added Database Name to " + scItem.Name);
                    else
                        Logger.PrintLC("Error adding Database Name to " + scItem.Name);
                CommitTransaction(trID);
                //Schema
                if (!string.IsNullOrWhiteSpace(entity.Schema))
                    if (con.AssignToObjModel(ref scItem, ConfigFile._TAB_NAME["Schema"], entity.Schema))
                        Logger.PrintLC("Added Schema to " + scItem.Name);
                    else
                        Logger.PrintLC("Error adding Schema to " + scItem.Name);
                CommitAndSave(trID);
            }
            return scItem;
        }


        public bool CommitAndSave(object id)
        {
            if (!CommitTransaction(id))
                return false;
            if (!SavePersistence())
                return false;
            return true;
        }

        public bool CommitTransaction(object id)
        {
            try
            {
                if (!scSession.CommitTransaction(id))
                {
                    Logger.PrintLC("Could not Commit for ID: " + id);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Committed successfully: " + id);
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Could NOT Commit.");
                return false;
            }
        }

        public bool SavePersistence()
        {
            try
            {
                if (!scPersistenceUnit.Save())
                {
                    Logger.PrintLC("Could NOT save Persistence: " + scPersistenceUnit.ObjectId);
                    return false;
                }
                else
                {
                    Logger.PrintLC("Persistence SAVED: " + scPersistenceUnit.ObjectId);
                    return true;
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Persistence NOT saved.");
                return false;
            }
        }

        public void CloseSession()
        {
            try
            {
                scSession.Close();
                Logger.PrintLC("Session closed successfully.");
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Could not close the Session.");
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
        public static bool AssignToObjModel(ref SCAPI.ModelObject model, string property, int value)
        {
            VBCon VBcon = new VBCon();
            if (VBcon.AssignToObjModel(ref model, property, value))
                return true;
            else
                return false;
        }
    }
}
