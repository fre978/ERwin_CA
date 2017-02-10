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
            
            if (actualDB.Contains(value.ToUpper()))
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
            Logger.PrintF(fileCorrect, message, true, ConfigFile.INFO);
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
            
            // verifica formale dei dati

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
                                    R.History = "Relazione ignorata: ID " + RStrut.ID + " contiene già una relazione col medesimo campo padre e/o col medesimo campo figlio";
                                    Logger.PrintLC("Relazione ignorata: ID " + RStrut.ID + " contiene già una relazione col medesimo campo padre e/o col medesimo campo figlio", 3, ConfigFile.ERROR);
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
    }
}
