using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class SqlOps
    {
        public SqlOps()
        {

        }

        public static bool CompareEntity(List<string> CollezioneSQL, List<EntityT> CollezioneXLS, ref Dictionary<string, List<string>> CompareResults)
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

        public static bool CompareAttribute(List<string> CollezioneSQL, List<AttributeT> CollezioneXLS, ref Dictionary<string, List<String>> CompareResults)
        {
            List<string> CollezioneTrovati = new List<string>();
            List<string> CollezioneNonTrovatiSQL = new List<string>();
            List<string> CollezioneNonTrovatiXLS = new List<string>();
            Dictionary<string, List<String>> MyDictionary = new Dictionary<string, List<string>>();

            try
            {
                foreach (AttributeT attribute in CollezioneXLS)
                {
                    string TableName = attribute.NomeTabellaLegacy.ToUpper();
                    string AttributeName = attribute.NomeCampoLegacy.ToUpper();
                    string Attributo = TableName + "." + AttributeName;
                    if (CollezioneSQL.Exists(x => x.Contains(Attributo)))
                    {
                        //presente nel file sql -> OK
                        CollezioneTrovati.Add(Attributo);
                    }
                    else
                    {
                        //non presente nel file sql -> KO
                        CollezioneNonTrovatiSQL.Add(Attributo);
                    }
                }

                foreach (string attribute in CollezioneSQL)
                {
                    string[] elementoCollezioneSQL = attribute.Split('|');
                    if (CollezioneTrovati.Exists(x => x == elementoCollezioneSQL[0]))
                    {
                        //esiste già nella tabella trovati -> OK
                        CollezioneTrovati.Remove(elementoCollezioneSQL[0]);
                        CollezioneTrovati.Add(attribute);
                    }
                    else
                    {
                        //non esiste sull'xls -> KO
                        CollezioneNonTrovatiXLS.Add(elementoCollezioneSQL[0]);
                    }
                }
                MyDictionary.Add("CollezioneAttributiTrovati", CollezioneTrovati);
                MyDictionary.Add("CollezioneAttributiNonTrovatiSQL", CollezioneNonTrovatiSQL);
                MyDictionary.Add("CollezioneAttributiNonTrovatiXLS", CollezioneNonTrovatiXLS);
                Logger.PrintLC(CollezioneSQL.Count() + " attributi esistenti nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneXLS.Count() + " attributi esistenti processate nel file XLS", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneTrovati.Count() + " attributi esistenti sia nel file XLS che nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneNonTrovatiSQL.Count() + " attributi non esistono nel file SQL", 3, ConfigFile.INFO);
                Logger.PrintLC(CollezioneNonTrovatiXLS.Count() + " attributi non esistono nel file XLS", 3, ConfigFile.INFO);
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
                    if ((myriga.ToUpper()).Contains("CREATE TABLE"))
                    {
                        RigaSplit = riga.Split(' ');
                        foreach (string e in RigaSplit)
                        {
                            if (e.ToUpper() == "TABLE")
                            {
                                string elemento = string.Empty;
                                //verifico che ci sia l'elemento successivo nell'array
                                if (RigaSplit.Count() >= i)
                                {
                                    elemento = RigaSplit[i + 1];
                                    //separo eventuali notazioni db.dbo.tabella prendendo l'ultimo elemento dell'insieme
                                    string[] arrayelemento = elemento.Split('.');
                                    elemento = arrayelemento[arrayelemento.Count() - 1];
                                    elemento = elemento.Replace("[", "").Replace("]", "").Replace("(", "").ToUpper();
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

        public static List<string> CollezionaAttributi(List<string> ListaRigheFileSQL)
        {
            List<string> ret = new List<string>();
            string[] RigaSplit;
            bool cerca = false;
            string memTable = string.Empty;
            string memConstraint = string.Empty;
            bool completeContraint = true;
            try
            {
                string entity = string.Empty;
                foreach (string riga in ListaRigheFileSQL)
                {
                    string myriga = riga.ToUpper();
                    int i = 0;
                    if (myriga.Contains("CONSTRAINT"))
                    {
                        if (myriga.Contains(";"))
                        {
                            //la constraint è completa cosi
                            memConstraint = myriga;
                            completeContraint = true;
                        }
                        else
                        {
                            //la constraint è suddivisa su piu righe
                            memConstraint = myriga;
                            completeContraint = false;
                            continue;
                        }
                    }
                    if (!(completeContraint))
                    {
                        memConstraint += myriga;
                        if (myriga.Contains(";"))
                        {
                            //la constraint è completa cosi
                            myriga = memConstraint;
                            completeContraint = true;
                        }
                        else
                        {
                            //la constraint non è ancora completa
                            continue;
                        }
                    }
                    //salto le righe vuote
                    if (string.IsNullOrEmpty(myriga))
                        continue;
                    //all'interno del ciclo cerco le righe contenenti create table
                    if ((myriga.ToUpper()).Contains("CREATE TABLE"))
                    {
                        //se trovo una create table devo controllare le righe successive
                        cerca = true;
                        RigaSplit = riga.Split(' ');
                        foreach (string e in RigaSplit)
                        {
                            if (e.ToUpper() == "TABLE")
                            {
                                string elemento = string.Empty;
                                //verifico che ci sia l'elemento successivo nell'array
                                if (RigaSplit.Count() >= i)
                                {
                                    elemento = RigaSplit[i + 1];
                                    //separo eventuali notazioni db.dbo.tabella prendendo l'ultimo elemento dell'insieme
                                    string[] arrayelemento = elemento.Split('.');
                                    elemento = arrayelemento[arrayelemento.Count() - 1];
                                    elemento = elemento.Replace("[", "").Replace("]", "").Replace("(", "").ToUpper();
                                    //usero entity come prefisso per gli attributi che troverò nelle righe successive
                                    entity = elemento;
                                    memTable = elemento;
                                    break;
                                }
                            }
                            i = i + 1;
                        }
                        continue;
                    }
                    if (myriga.Trim().Equals("("))
                    {
                        continue;
                    }

                    if (myriga.Trim().Equals(");"))
                    {
                        if (cerca)
                            cerca = false;
                        continue;
                    }

                    if ((myriga.Trim().Contains("CONSTRAINT")))
                    {
                        if (cerca) 
                            cerca = false;

                        if (string.IsNullOrEmpty(memTable))
                        {
                            //non ho una tabella su cui lavorare
                        }
                        else
                        {
                            if (myriga.ToUpper().Contains("PRIMARY"))
                            {
                                string constraint = myriga.ToUpper().Substring(myriga.IndexOf("(") + 1, myriga.IndexOf(")") - myriga.IndexOf("(") - 1);
                                string[] keys = constraint.Split(',');
                                foreach (string key in keys)
                                    if (ret.Exists(x => x.Contains(memTable + "." + key.Trim())))
                                    {
                                        string temp = ret.Find(x => x.Contains(memTable + "." + key.Trim()));
                                        ret.Remove(temp);
                                        ret.Add(temp + "true");
                                    }
                            }
                        }
                        continue;
                    }
                    if ((myriga.Trim().Contains("TABLE")))
                    {
                        if (cerca)
                            cerca = false;
                        //all'interno del ciclo cerco le righe contenenti create table
                        if ((myriga.ToUpper()).Contains("ALTER TABLE"))
                        {
                            //se trovo una alter table devo verificare che ci sia una constraint nelle righe successive
                            RigaSplit = riga.Split(' ');
                            foreach (string e in RigaSplit)
                            {
                                if (e.ToUpper() == "TABLE")
                                {
                                    string elemento = string.Empty;
                                    //verifico che ci sia l'elemento successivo nell'array
                                    if (RigaSplit.Count() >= i)
                                    {
                                        elemento = RigaSplit[i + 1];
                                        //separo eventuali notazioni db.dbo.tabella prendendo l'ultimo elemento dell'insieme
                                        string[] arrayelemento = elemento.Split('.');
                                        elemento = arrayelemento[arrayelemento.Count() - 1];
                                        //usero entity come prefisso per gli attributi che troverò nelle righe successive
                                        memTable = elemento;
                                        break;
                                    }
                                }
                                i = i + 1;
                            }
                        }
                        continue;
                    }
                    if ((myriga.Trim().Contains("LABEL")))
                    {
                        if (cerca)
                            cerca = false;
                        continue;
                    }

                    //se arrivo a questo punto del ciclo è una riga di attributi e provo ad aggiungere alla collezione quello che ritengo essere il mio attributo
                    if (cerca)
                    {
                        try
                        {
                            //splitto le parole della riga
                            RigaSplit = riga.Split(' ');

                            bool cercaDatatype = false;
                            string datatype = string.Empty;
                            string attribute = string.Empty;
                            string mandatory = riga.ToUpper().Contains("NOT NULL") ? "true" : "false";
                            string key = string.Empty;


                            //il nome dell'attributo è il primo elemento di una riga di dichiarazione degli attributi ma può essere nel formato db.dbo.tabella.attributo
                            foreach (string x in RigaSplit)
                            {
                                if (!(string.IsNullOrEmpty(x)))
                                {
                                    if (!(cercaDatatype))
                                    { 
                                        string[] arrayelemento = x.Split('.');
                                        //prendo l'ultimo elemento dell'attributo e lo aggiungo alla lista di attributi se non esiste
                                        string elemento = entity + "." + arrayelemento[arrayelemento.Count() - 1];
                                        if (!(ret.Exists(y => y == elemento)))
                                        {
                                            attribute = elemento;
                                            cercaDatatype = true;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        datatype += x;
                                        if ((datatype.Contains('(')) && (!datatype.Contains(')')))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                            }

                            if (!(ret.Exists(y => y.Contains(attribute))))
                            {
                                datatype = datatype.Trim().EndsWith(",") ? 
                                    datatype.Trim().Substring(0, datatype.Trim().Length - 1) 
                                    : datatype;
                                ret.Add(attribute + "|" + datatype + "|" + mandatory + "|" + key);
                            }

                            continue;


                        }
                        catch
                        {
                            //errore nella ricerca dell'attributo
                        }
                    }
                }
                return ret;
            }
            catch (Exception exp)
            {
                return ret;
            }
        }
    }
}
