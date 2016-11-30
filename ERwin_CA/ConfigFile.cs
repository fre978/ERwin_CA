using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    public class ConfigFile : IDisposable
    {
        void IDisposable.Dispose()
        {

        }

        public static string LOG_FILE = @"C:\ERWIN\Log.txt";

        public static string TABELLE = "Censimento Tabelle";
        public static string COLONNA_01 = "SSA";
        public static int HEADER_RIGA = 3;
        public static int HEADER_COLONNA_MIN = 1;
        public static int HEADER_COLONNA_MAX = 10;
        public static int HEADER_MAX_COLONNE = 10;

        public static int SSA;

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

        public static string FILETEST = @"C:\ERWIN\CODICE\Extra\Test.xlsx";

    }
}
