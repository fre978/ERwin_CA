using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
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

        // SEZIONE ESECUZIONE
        private static int LOG_LEVEL_DEFAULT = 4;
        public static int LOG_LEVEL = LOG_LEVEL_DEFAULT;
        public static bool RefreshLogLevel()
        {
            try
            {
                int.TryParse(ConfigurationSettings.AppSettings["Log Level"], out LOG_LEVEL);
                return true;
            }
            catch
            {
                LOG_LEVEL = LOG_LEVEL_DEFAULT;
                return false;
            }
        }

        public static string ERROR = "ERR: ";
        public static string WARNING = "WARN: ";
        public static string INFO = "INFO: ";

        public static string BASE_PATH = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
        public static string APP_PATH = System.IO.Path.GetDirectoryName(BASE_PATH).Replace("file:\\", "");
        public static string SEARCH_PATH = ConfigurationSettings.AppSettings["Search Folder"];
        public static string INPUT_FOLDER_NAME = ConfigurationSettings.AppSettings["Input Folder Name"];

        public static string CREACOPIEERWIN = ConfigurationSettings.AppSettings["CREACOPIEERWIN"];
        public static string PERCORSOCOPIEERWIN = APP_PATH + @"\" + ConfigurationSettings.AppSettings["PERCORSOCOPIEERWIN"] + @"\";
        public static string PERCORSOCOPIEERWINDESTINATION;

        // Sezione file remoti
        public static string COPY_TO_LOCAL = ConfigurationSettings.AppSettings["Copy to Local"];
        public static string LOCAL_TEMP_DIR = ConfigurationSettings.AppSettings["Local Folder Name"];
        public static bool COPY_LOCAL = false;
        public static string LOCAL_DIR_FULL = null;
        public static string TEMP_REMOTE_FILE = null;
        public static string TEMP_REMOTE_ROOT = null;
        //public static string 
        
        public static bool RefreshLocal()
        {
            try
            {
                if (COPY_TO_LOCAL.Trim().ToUpper() == "TRUE")
                {
                    COPY_LOCAL = true;
                }
                else
                {
                    COPY_LOCAL = false;
                }

                if (!string.IsNullOrWhiteSpace(LOCAL_TEMP_DIR))
                {
                    string tempDir = Path.Combine(APP_PATH, LOCAL_TEMP_DIR);
                    DirectoryInfo localTemp = new DirectoryInfo(tempDir);
                    if (!localTemp.Exists)
                    {
                        try
                        {
                            Directory.CreateDirectory(tempDir);
                        }
                        catch
                        {
                            COPY_LOCAL = false;
                            LOCAL_DIR_FULL = string.Empty;
                            return false;
                        }
                    }
                    LOCAL_DIR_FULL = localTemp.FullName;
                }
                else
                {
                    COPY_LOCAL = false;
                    LOCAL_DIR_FULL = string.Empty;
                    return true;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        // SEZIONE DATABASE
        public static string ERWIN_TEMPLATE_DB2 = APP_PATH + @"\Template\Template_DB2_LF.erwin";
        public static string ERWIN_TEMPLATE_ORACLE = APP_PATH + @"\Template\Template_Oracle_LF.erwin";
        public static string ERWIN_TEMPLATE_SQLSERVER = APP_PATH + @"\Template\Template_SqlServer_LF.erwin";
        public static string CONTROLLI_CAMPI_TEMPLATE = APP_PATH + @"\Template\Controlli_Campi_v4.xlsx";
        public static string CONTROLLI_TEMPISTICHE_TEMPLATE = APP_PATH + @"\Template\Controlli_Tempistiche_v7.xlsx";

        private static string tempString = ConfigurationSettings.AppSettings["DBS"].ToUpper();
        public static List<string> DBS = tempString.Split(',').ToList(); //new List<string> { "DB2", "ORACLE", "SQLSERVER }; //Sempre Upper case

        public const string DB2_NAME = "DB2";
        public const string ORACLE = "ORACLE";
        public const string SQLSERVER = "SQLSERVER";
        //public static string dd = AppDomain.CurrentDomain.BaseDirectory;

        // SEZIONE FILE
        public static string LOG_FILE = APP_PATH + @"\" + ConfigurationSettings.AppSettings["PERCORSOLOG"] + @"\Log.txt";
        public static string ROOT = SEARCH_PATH;

        // SEZIONE CARTELLE
        public static string DEST_FOLD_NAME = ConfigurationSettings.AppSettings["Destination Folder Name"];
        public static bool DEST_FOLD_UNIQUE = (ConfigurationSettings.AppSettings["Destination Folder Unique"] == "true") ? true : false;
        public static string FOLDERDESTINATION_GENERAL = Path.Combine(ROOT, DEST_FOLD_NAME);
        public static string FOLDERDESTINATION;
        public static string TIMESTAMPFOLDER;

        // SEZIONE GENERALE
        public static char[] DELIMITER_DATABASE_NAME = new char[10];
        public static bool RefreshDatabaseDelimiters()
        {
            try
            {
                List<string> lista = ConfigurationSettings.AppSettings["Database Name Delimiter"].Split('|').ToList();
                if (lista.Count > 10)
                    lista.RemoveRange(10, lista.Count - 10);
                string[] newLista = lista.ToArray();
                int x = 0;
                foreach (string elemento in lista)
                {
                    DELIMITER_DATABASE_NAME[x] = elemento[0];
                    x++;
                }
                return true;
            }
            catch
            {
                DELIMITER_DATABASE_NAME[0] = '-';
                return false;
            }
        }

        public static char[] DELIMITER_NAME_FILE = new char[10];
        public static bool RefreshDelimiters()
        {
            try
            {
                List<string> lista = ConfigurationSettings.AppSettings["File Name Delimiter"].Split('|').ToList();
                if (lista.Count > 10)
                    lista.RemoveRange(10, lista.Count - 10);
                string[] newLista = lista.ToArray();
                int x = 0;
                foreach (string elemento in lista)
                {
                    DELIMITER_NAME_FILE[x] = elemento[0];
                    x++;
                }
                return true;
            }
            catch
            {
                DELIMITER_NAME_FILE[0] = '_';
                DELIMITER_NAME_FILE[1] = '.';
                return false;
            }
        }
        public static string[] Yes = { "S", "SI" };
        public static string[] No = { "N", "NO" };

        public static string[] DATATYPE_DB2 = {"char", "char()", "varchar()", "clob", "clob()",
                                               "date", "time", "timestamp", "timestamp()",
                                               "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint",
                                               "blob", "blob()", "binary", "binary()"};
        public static bool RefreshDatatypeDB2()
        {
            try
            {
                DATATYPE_DB2 = ConfigurationSettings.AppSettings["DB2 Types"].Split('|');
                return true;
            }
            catch
            {
               DATATYPE_DB2 = new string[] { "char", "char()", "varchar()", "clob", "clob()",
                                               "date", "time", "timestamp", "timestamp()",
                                               "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint",
                                               "blob", "blob()", "binary", "binary()"};
                return false;
            }
        }

        public static string[] DATATYPE_SQLSERVER = { "char", "char()", "varchar", "varchar()", "xml", "text",
                                                    "date", "datetime", "time", "time()", "timestamp", "smalldatetime", "datetime2", "datetime2()",
                                                    "decimal", "decimal()", "decimal(,)", "bit", "bigint", "double precision", "float", "float()", "real", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "money", "smallmoney", "tinyint", "uniqueidentifier",
                                                    "binary", "binary()", "image", "sql_variant", "varbinary", "varbinary()" };
        public static bool RefreshDatatypeSQLSERVER()
        {
            try
            {
                DATATYPE_SQLSERVER = ConfigurationSettings.AppSettings["SQLSERVER Types"].Split('|');
                return true;
            }
            catch
            {
                DATATYPE_SQLSERVER = new string[] { "char", "char()", "varchar", "varchar()", "xml", "text",
                                                    "date", "datetime", "time", "time()", "timestamp", "smalldatetime", "datetime2", "datetime2()",
                                                    "decimal", "decimal()", "decimal(,)", "bit", "bigint", "double precision", "float", "float()", "real", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "money", "smallmoney", "tinyint", "uniqueidentifier",
                                                    "binary", "binary()", "image", "sql_variant", "varbinary", "varbinary()" };
                return false;
            }
        }
        public static string[] DATATYPE_ORACLE = {"char", "char()", "varchar()", "clob", "clob()", "varchar2()",
                                                  "date", "timestamp", "timestamp()",
                                                  "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "number", "number()", "number(,)",
                                                  "blob"};
        public static bool RefreshDatatypeOracle()
        {
            try
            {
                DATATYPE_ORACLE = ConfigurationSettings.AppSettings["ORACLE Types"].Split('|');
                return true;
            }
            catch
            {
                DATATYPE_ORACLE = new string[] {  "char", "char()", "varchar()", "clob", "clob()", "varchar2()",
                                                  "date", "timestamp", "timestamp()",
                                                  "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "number", "number()", "number(,)",
                                                  "blob"};
                return false;
            }
        }

        public static string[] DATATYPE_DB2_FOR = {"date", "time", "timestamp", "timestamp()",
                                               "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint" };
        public static bool RefreshDatatypeDB2_FOR()
        {
            try
            {
                DATATYPE_DB2_FOR = ConfigurationSettings.AppSettings["DB2 Types FOR"].Split('|');
                return true;
            }
            catch
            {
                DATATYPE_DB2_FOR = new string[] {"date", "time", "timestamp", "timestamp()",
                                               "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint"};
                return false;
            }
        }

        public static string[] DATATYPE_SQLSERVER_FOR = {"date", "datetime", "time", "time()", "timestamp", "smalldatetime", "datetime2", "datetime2()",
                                                    "decimal", "decimal()", "decimal(,)", "bit", "bigint", "double precision", "float", "float()", "real", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "money", "smallmoney", "tinyint", "uniqueidentifier"};
        public static bool RefreshDatatypeSQLSERVER_FOR()
        {
            try
            {
                DATATYPE_SQLSERVER_FOR = ConfigurationSettings.AppSettings["SQLSERVER Types FOR"].Split('|');
                return true;
            }
            catch
            {
                DATATYPE_SQLSERVER_FOR = new string[] {"date", "datetime", "time", "time()", "timestamp", "smalldatetime", "datetime2", "datetime2()",
                                                    "decimal", "decimal()", "decimal(,)", "bit", "bigint", "double precision", "float", "float()", "real", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "money", "smallmoney", "tinyint", "uniqueidentifier"};
                return false;
            }
        }
        public static string[] DATATYPE_ORACLE_FOR = {"date", "timestamp", "timestamp()",
                                                  "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "number", "number()", "number(,)"};
        public static bool RefreshDatatypeOracle_FOR()
        {
            try
            {
                DATATYPE_ORACLE_FOR = ConfigurationSettings.AppSettings["ORACLE Types FOR"].Split('|');
                return true;
            }
            catch
            {
                DATATYPE_ORACLE_FOR = new string[] {  "date", "timestamp", "timestamp()",
                                                  "decimal", "decimal()", "decimal(,)", "dec", "dec()", "dec(,)", "numeric", "numeric()", "numeric(,)", "integer", "int", "smallint", "number", "number()", "number(,)"};
                return false;
            }
        }

        public static bool RefreshAll()
        {
            bool response = false;
            if (!RefreshLogLevel())
                response = true;
            if (!RefreshDelimiters())
                response = true;
            if (!RefreshDatabaseDelimiters())
                response = true;
            if (!RefreshDatatypeDB2())
                response = true;
            if (!RefreshDatatypeOracle())
                response = true;
            if (!RefreshDatatypeSQLSERVER())
                response = true;
            if (!RefreshDatatypeDB2_FOR())
                response = true;
            if (!RefreshDatatypeOracle_FOR())
                response = true;
            if (!RefreshDatatypeSQLSERVER_FOR())
                response = true;
            if (!RefreshColumns())
                response = true;
            if (!RefreshLocal())
                response = true;
            return response;
        }

        

        public static bool RefreshColumns()
        {
            try
            {
                int tempInt = 0;
                int.TryParse(ConfigurationSettings.AppSettings["Header Row"], out tempInt);
                HEADER_RIGA = tempInt;

                int.TryParse(ConfigurationSettings.AppSettings["First Tables Column"], out tempInt);
                HEADER_COLONNA_MIN_TABELLE = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Last Tables Column"], out tempInt);
                HEADER_COLONNA_MAX_TABELLE = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Tables Columns Number"], out tempInt);
                HEADER_MAX_COLONNE_TABELLE = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Tables Columns Offset 1"], out tempInt);
                TABELLE_EXCEL_COL_OFFSET1 = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Tables Columns Offset 2"], out tempInt);
                TABELLE_EXCEL_COL_OFFSET2 = tempInt;

                int.TryParse(ConfigurationSettings.AppSettings["First Attributes Column"], out tempInt);
                HEADER_COLONNA_MIN_ATTRIBUTI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Last Attributes Column"], out tempInt);
                HEADER_COLONNA_MAX_ATTRIBUTI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Attributes Columns Number"], out tempInt);
                HEADER_MAX_COLONNE_ATTRIBUTI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Attributes Columns Offset 1"], out tempInt);
                ATTRIBUTI_EXCEL_COL_OFFSET1 = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Attributes Columns Offset 2"], out tempInt);
                ATTRIBUTI_EXCEL_COL_OFFSET2 = tempInt;

                int.TryParse(ConfigurationSettings.AppSettings["First Relations Column"], out tempInt);
                HEADER_COLONNA_MIN_RELAZIONI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Last Relations Column"], out tempInt);
                HEADER_COLONNA_MAX_RELAZIONI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Relations Columns Number"], out tempInt);
                HEADER_MAX_COLONNE_RELAZIONI = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Relations Columns Offset 1"], out tempInt);
                RELAZIONI_EXCEL_COL_OFFSET1 = tempInt;
                int.TryParse(ConfigurationSettings.AppSettings["Relations Columns Offset 2"], out tempInt);
                RELAZIONI_EXCEL_COL_OFFSET2 = tempInt;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public const string TABELLE =   "Censimento Tabelle";
        public const string ATTRIBUTI = "Censimento Attributi";
        public const string RELAZIONI = "Relazioni-ModelloDatiLegacy";
        public const string TABELLE_DIFF = "Differenze Tabelle";
        public const string ATTRIBUTI_DIFF = "Differenze Attributi";
        public const string CONTROLLI_CAMPI = "Controlli_campi";
        public const string CONTROLLI_TEMPISTICHE = "Controlli_tempistiche";


        public static string COLONNA_01 = "SSA";
        public static int HEADER_RIGA = 3;

        public static int HEADER_COLONNA_MIN_TABELLE = 1;
        public static int HEADER_COLONNA_MAX_TABELLE = 10;
        public static int HEADER_MAX_COLONNE_TABELLE = 10;

        public static int HEADER_COLONNA_MIN_ATTRIBUTI = 1;
        public static int HEADER_COLONNA_MAX_ATTRIBUTI = 18;
        public static int HEADER_MAX_COLONNE_ATTRIBUTI = 18;

        public static int HEADER_COLONNA_MIN_RELAZIONI = 1;
        public static int HEADER_COLONNA_MAX_RELAZIONI = 10;
        public static int HEADER_MAX_COLONNE_RELAZIONI = 10;

        public static int SSA;

        // SEZIONE DICTIONARY TABELLE
        public static int TABELLE_EXCEL_COL_OFFSET1 = 1;
        public static int TABELLE_EXCEL_COL_OFFSET2 = 2;
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
            {"Nome host DB2_SQL", "Entity.Physical.NOME_HOST" },
            {"Nome Database DB2_SQL", "Entity.Physical.NOME_DATABASE" },
            //Sezione ORACLE
            {"Nome host Oracle", "Entity.Physical.NOME_HOST" },
            {"Nome Database Oracle", "Entity.Physical.NOME_DATABASE" },
            {"Schema Oracle", "Name_Qualifier" },
            //Fine ORACLE
            //Sezione SQLSERVER
            {"Nome host SQLSERVER", "SQLServer_Database.Physical.NOME_HOST" },
            {"Nome Database SQLSERVER", "Name" },
            //Fine SQLSERVER
            {"Nome Tabella", "Physical_Name" },
            {"Descrizione Tabella", "Comment" },
            {"Tipologia Informazione", "Entity.Physical.TIPOLOGIA_INFORMAZIONE" },
            {"Perimetro Tabella", "Entity.Physical.PERIMETRO_TABELLA" },
            {"Granularità Tabella", "Entity.Physical.GRANULARITA_TABELLA" },
            {"Flag BFD", "Entity.Physical.FLAG_BFD" }
        };
        // ##############################

        // SEZIONE DICTIONARY ATTRIBUTI
        public static int ATTRIBUTI_EXCEL_COL_OFFSET1 = 7;
        public static int ATTRIBUTI_EXCEL_COL_OFFSET2 = 8;
        public static Dictionary<string, int> _ATTRIBUTI = new Dictionary<string, int>()
        {
            {"SSA", 1 },
            {"Area", 2 },
            {"Nome Tabella Legacy", 3 },
            {"Nome  Campo Legacy", 4 }, //ATTENZIONE doppio spazio tra 'Nome' e 'Campo'
            {"Definizione Campo", 5 },
            {"Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati", 6 },
            {"Datatype", 7 },
            {"Lunghezza", 8 },
            {"Decimali", 9 },
            {"Chiave", 10 },
            {"Unique", 11 },
            {"Chiave Logica", 12 },
            {"Mandatory Flag", 13 },
            {"Dominio", 14 },
            {"Provenienza dominio ", 15 }, //ATTENZIONE minuscole e spazio finale
            {"Note", 16 },
            {"Storica", 17 },
            {"Dato Sensibile", 18 }
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
            {"Dato Sensibile", "Attribute.Physical.DATO_SENSIBILE" },
            {"Ordine", "Attribute_Order" }
        };
        // ##############################


        // SEZIONE RELAZIONI
        public static int RELAZIONI_EXCEL_COL_OFFSET1 = 1;
        public static int RELAZIONI_EXCEL_COL_OFFSET2 = 2;
        public static Dictionary<string, int> _RELAZIONI = new Dictionary<string, int>()
        {
            {"Identificativo relazione", 1 },
            {"Tabella Padre", 2 },
            {"Tabella Figlia", 3 },
            {"Cardinalità", 4 },
            {"Campo Padre", 5 },
            {"Campo Figlio", 6 },
            {"Identificativa", 7 },
            {"Eccezioni", 8 },
            {"Tipo Relazione", 9 },
            {"Note", 10 }
        };

        public static Dictionary<string, string> _REL_NAME = new Dictionary<string, string>()
        {
            {"Identificativo relazione", "Name" },
            {"Tabella Padre", "Parent_Entity_Ref" },
            {"Tabella Figlia", "Child_Entity_Ref" },
            {"Cardinalita", "Cardinality" },
            {"Null Option Type", "Null_Option_Type" },
            {"Identificativa", "Type" },
            {"Eccezioni", "Relationship.Physical.ECCEZIONI" },
            {"Tipo Relazione", "Do_Not_Generate" },
            {"Note", "Relationship.Physical.NOTE" }
        };

        // ##############################

        public static bool DDL_Show_Right_Rows = (ConfigurationSettings.AppSettings["DDL Show Right Rows"] == "true") ? true : false;
    }
}
