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
        public SCAPI.ModelObjects erColumn;
        public SCAPI.ModelObjects erTable;
        public SCAPI.ModelObject scItem;
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
