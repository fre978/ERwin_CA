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
        public SCAPI.PersistenceUnit scPersistanceUnit;
        public SCAPI.Session scSession;
        public SCAPI.ModelObject erRootObj { get; set; }
        public SCAPI.ModelObjects erRootObjCol { get; set; }
        public SCAPI.ModelObjects erColumn { get; set; }
        public SCAPI.ModelObjects erTable { get; set; }
        public SCAPI.ModelObject scItem { get; set; }
        public long trID = 0;

        public void openModelConnection(string ERw)
        {
            if (ERw == null)
                return;
            try
            {
                scERwin = new SCAPI.Application();
                scPersistanceUnit = scERwin.PersistenceUnits.Add(ERw, "RDO=No");

                scSession = scERwin.Sessions.Add();
                scSession.Open(scPersistanceUnit);
                Logger.PrintLC("Connection opened.");
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Connection opening error: " + exp.Message);
                return;
            }
        }

        public long openTransaction()
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
        public void SetRootObject()
        {
            if (scSession != null)
                try
                {
                    erRootObj = scSession.ModelObjects.Root;
                    Logger.PrintLC("Root has been successful.");
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Setting Root's Session error: " + exp.Message);
                }
            else
                Logger.PrintLC("Could not determine Root because Session is missing.");
        }
        public void SetRootCollection()
        {
            if (scSession != null)
                try
                {
                    erRootObjCol = scSession.ModelObjects.Collect(erRootObj);
                    Logger.PrintLC("Root Collection has been successful.");
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not get the Root Collection: " + exp.Message);
                }
            else
            {
                Logger.PrintLC("Could not get Root Collection because Session is missing.");
            }
        }

        public SCAPI.ModelObject CreateEntity (EntityT entity)
        {
            SCAPI.ModelObject mObject = new SCAPI.ModelObject();
            if (erRootObjCol != null)
            {
                //scSession.ModelObjects.Collect("Entity");
                mObject = erRootObjCol.Add("Entity");
                VBCon con = new VBCon();
                //Nome tabella
                if (!string.IsNullOrWhiteSpace(entity.TableName))
                    if (con.AssignToObjModel(ref mObject, ConfigFile._TAB_NAME["Nome Tabella"], entity.TableName))
                        Logger.PrintLC("Added Physical Name to " + mObject.Name);
                    else
                        Logger.PrintLC("Error adding Physical Name to " + mObject.Name);
                //SSA
                if (!string.IsNullOrWhiteSpace(entity.SSA))
                    if (con.AssignToObjModel(ref mObject, ConfigFile._TAB_NAME["SSA"], entity.SSA))
                        Logger.PrintLC("Added SSA to " + mObject.Name);
                    else
                        Logger.PrintLC("Error adding SSA to " + mObject.Name);
                //Nome Host
                if (!string.IsNullOrWhiteSpace(entity.HostName))
                    if (con.AssignToObjModel(ref mObject, ConfigFile._TAB_NAME["Nome Host"], entity.HostName))
                        Logger.PrintLC("Added Host Name to " + mObject.Name);
                    else
                        Logger.PrintLC("Error adding Host Name to " + mObject.Name);
                //Nome Database
                if (!string.IsNullOrWhiteSpace(entity.DatabaseName))
                    if (con.AssignToObjModel(ref mObject, ConfigFile._TAB_NAME["Nome Database"], entity.DatabaseName))
                        Logger.PrintLC("Added Database Name to " + mObject.Name);
                    else
                        Logger.PrintLC("Error adding Database Name to " + mObject.Name);
                //Schema
                if (!string.IsNullOrWhiteSpace(entity.Schema))
                    if (con.AssignToObjModel(ref mObject, ConfigFile._TAB_NAME["Schema"], entity.Schema))
                        Logger.PrintLC("Added Schema to " + mObject.Name);
                    else
                        Logger.PrintLC("Error adding Schema to " + mObject.Name);
            }
            return mObject;
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
