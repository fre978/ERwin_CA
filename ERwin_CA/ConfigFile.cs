﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    public static class ConfigFile //: IDisposable
    {
        //void IDisposable.Dispose()
        //{

        //}

        // SEZIONE DATABASE
        public const string ERWIN_TEMPLATE_DB2 = @"D:\TEST\Template_DB2_LF.erwin";
        public const string ERWIN_TEMPLATE_ORACLE = @"D:\TEST\Template_Oracle_LF.erwin";
        public const string DB2_NAME = "DB2";
        public const string ORACLE = "Oracle";

        // SEZIONE FILE
        public static string LOG_FILE = @"D:\TEST\Log.txt";
        public static string ERWIN_FILE = @"D:\ERwin\Template_DB2_LF - Copia.erwin";
        public static string FILETEST = @"D:\TEST\TestFiles\";
        
        // SEZIONE CARTELLE
        public static string FOLDERDESTINATION = @"D:\TEST\Destinazione";

        // SEZIONE GENERALE
        public static char[] DELIMITER_NAME_FILE = { '_', '.' };

        public const string TABELLE =   "Censimento Tabelle";
        public const string ATTRIBUTI = "Censimento Attributi";
        public const string RELAZIONI = "Relazioni-ModelloDatiLegacy";
        public static string COLONNA_01 = "SSA";
        public static int HEADER_RIGA = 3;
        public static int HEADER_COLONNA_MIN = 1;
        public static int HEADER_COLONNA_MAX = 10;
        public static int HEADER_MAX_COLONNE = 10;

        public static int SSA;

        // SEZIONE DICTIONARY TABELLE
        public static Dictionary<string, int> _TABELLE = new Dictionary<string, int>()
        {
            {"SSA", 1 },
            {"Nome host", 2 },
            {"Nome Database", 3 },
            {"Schema", 4 },
            {"Nome Tabella", 5 },
            {"Descrizione Tabella", 6 },
            {"Tipologia Informazione", 7 },
            {"Perimetro Tabella", 8 },
            {"Granularità Tabella", 9 },
            {"Flag BFD", 10 }
        };

        public static Dictionary<string, string> _TAB_NAME = new Dictionary<string, string>()
        {
            {"SSA", "Entity.Physical.SSA" },
            {"Nome host", "DB2_Database.Physical.NOME_HOST" },
            {"Nome Database", "Name" },
            //{"Schema", "Name_Qualifier" },
            {"Schema", "Schema_Ref" },
            {"Nome Tabella", "Physical_Name" },
            {"Descrizione Tabella", "Comment" },
            {"Tipologia Informazione", "Entity.Physical.TIPOLOGIA_INFORMAZIONE" },
            {"Perimetro Tabella", "Entity.Physical.PERIMETRO_TABELLA" },
            {"Granularità Tabella", "Entity.Physical.GRANULARITA_TABELLA" },
            {"Flag BFD", "Entity.Physical.FLAG_BFD" }
        };


        // SEZIONE DICTIONARY ATTRIBUTI
        public static Dictionary<string, int> _ATTRIBUTI = new Dictionary<string, int>()
        {
            {"SSA", 1 },
            {"Area", 2 },
            {"Nome Tabella Legacy", 3 },
            {"Nome Campo Legacy", 4 },
            {"Definizione Campo", 5 },
            {"Tipologia Tabella", 6 },
            {"Datatype", 7 },
            {"Lunghezza", 8 },
            {"Decimali", 9 },
            {"Chiave", 11 },
            {"Unique", 12 },
            {"Chiave Logica", 13 },
            {"Mandatory Flag", 14 },
            {"Dominio", 15 },
            {"Provenienza Dominio", 16 },
            {"Note", 17 },
            {"Storica", 18 },
            {"Dato Sensibile", 19 }
        };

        public static Dictionary<string, string> _ATT_NAME = new Dictionary<string, string>()
        {
            {"SSA", "" },
            {"Area", "Entity.Physical.AREA" },
            {"Nome Tabella Legacy", "Name" },
            {"Nome Campo Legacy", "Physical_Name" },
            {"Nome Campo Legacy Name", "Name" },
            {"Definizione Campo", "Comment" },
            {"Definizione Campo Def", "Definition" },
            {"Tipologia Tabella", "Entity.Physical.TIPOLOGIA_TABELLA" },
            {"Datatype", "Physical_Data_Type" },
            {"Lunghezza", "" },
            {"Decimali", "" },
            {"Chiave", "Type" },
            {"Unique", "Attribute.Physical.UNIQUE" },
            {"Chiave Logica", "Attribute.Physical.CHIAVE_LOGICA" },
            {"Mandatory Flag", "Null_Option_Type" },
            {"Dominio", "Attribute.Physical.DOMINIO" },
            {"Provenienza Dominio", "Attribute.Physical.PROVENIENZA_DOMINIO" },
            {"Note", "Attribute.Physical.NOTE" },
            {"Storica", "Entity.Physical.STORICA" },
            {"Dato Sensibile", "Attribute.Physical.DATO_SENSIBILE" }
        };



    }
}
