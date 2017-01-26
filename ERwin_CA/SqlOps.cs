using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ERwin_CA.T;

namespace ERwin_CA
{
    class SqlOps
    {
        public SqlOps()
        {

        }

        public static bool CompareEntity(List<string> CollezioneSQL, List<EntityT> CollezioneXLS, ref Dictionary<string, List<String>> CompareResults)
        {
            List<string> CollezioneTrovati = new List<string>();
            List<string> CollezioneNonTrovatiSQL = new List<string>();
            List<string> CollezioneNonTrovatiXLS = new List<string>();
            Dictionary<string, List<String>> MyDictionary = new Dictionary<string, List<string>>();

            try
            {
                foreach (EntityT entity in CollezioneXLS)
                {
                    string TableName = entity.TableName.ToUpper();
                    if (CollezioneSQL.Exists(x => x == TableName))
                    {
                        //presente nel file sql -> OK
                        CollezioneTrovati.Add(TableName);
                    }
                    else
                    {
                        //non presente nel file sql -> KO
                        CollezioneNonTrovatiSQL.Add(TableName);
                    }
                }

                foreach (string entity in CollezioneSQL)
                {
                    if (CollezioneTrovati.Exists(x => x == entity))
                    {
                        //esiste già nella tabella trovati -> OK
                    }
                    else
                    {
                        //non esiste sull'xls -> KO
                        CollezioneNonTrovatiXLS.Add(entity);
                    }
                }
                MyDictionary.Add("CollezioneTrovati", CollezioneTrovati);
                MyDictionary.Add("CollezioneNonTrovatiSQL", CollezioneNonTrovatiSQL);
                MyDictionary.Add("CollezioneNonTrovatiXLS", CollezioneNonTrovatiXLS);
                Logger.PrintLC(CollezioneSQL.Count() + " entity esistenti nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneXLS.Count() + " entity esistenti processate nel file XLS", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneTrovati.Count() + " entity esistenti sia nel file XLS che nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneNonTrovatiSQL.Count() + " entity non esistono nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneNonTrovatiXLS.Count() + " entity non esistono nel file XLS", 3, ConfigFile.INFO);
                CompareResults = MyDictionary;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static List<string> CollezionaEntity(List<string> ListaRigheFileSQL)
        {
            List<string> ret = new List<string>();
            string Entity;
            string[] RigaSplit;
            try
            {
                foreach (string riga in ListaRigheFileSQL)
                {
                    Entity = string.Empty;
                    string myriga = riga.ToUpper();
                    int i = 0;
                    if (myriga.Contains("TABLE"))
                    {
                        RigaSplit = riga.Split(' ');
                        foreach (string e in RigaSplit)
                        {
                            if (e == "TABLE")
                            {
                                string elemento = string.Empty;
                                //verifico che ci sia l'elemento successivo nell'array
                                if (RigaSplit.Count() >= i)
                                {
                                    elemento = RigaSplit[i + 1];
                                    //separo eventuali notazioni db.dbo.tabella prendendo l'ultimo elemento dell'insieme
                                    string[] arrayelemento = elemento.Split('.');
                                    elemento = arrayelemento[arrayelemento.Count() - 1];
                                }
                                if (!(ret.Exists(x => x == elemento)))
                                    ret.Add(elemento);
                                continue;
                            }
                            i = i + 1;
                        }
                    }

                }
                return ret;
            }
            catch
            {
                return ret;
            }
        }
    }
}
