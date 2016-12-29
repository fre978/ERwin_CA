using ERwin_CA.T;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class Parser
    {
        public static FileT ParseFileName(string fileName)
        {
            FileT file = new FileT();
            FileInfo fileNameInfo = new FileInfo(fileName);
            string[] fileComponents;
            fileComponents = fileNameInfo.Name.Split(ConfigFile.DELIMITER_NAME_FILE);
            int length = fileComponents.Count();

            if (length != 5)
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension>", 2);
                return file = null;
            }
            if (!ConfigFile.DBS.Contains(fileComponents[3].ToUpper()))
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension>", 2);
                return file = null;
            }

            try
            {
                file.SSA = fileComponents[0];
                file.Acronimo = fileComponents[1];
                file.NomeModello = fileComponents[2];
                file.TipoDBMS = fileComponents[3].ToUpper();
                file.Estensione = fileComponents[4];
            }
            catch (Exception exp)
            {
                Logger.PrintLC(fileName + "produced an error while parsing its name: " + exp.Message, 2);
                return file = null;
            }

            return file;
        }
    }
}
