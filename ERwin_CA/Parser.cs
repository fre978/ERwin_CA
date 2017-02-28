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

        public static List<string> ParseListOfFileNames(List<string> fileList)
        {
            List<string> result = new List<string>();
            foreach(string file in fileList)
            {
                FileT parsedResult = ParseFileName(file);
                if (parsedResult != null)
                {
                    result.Add(file);
                }
            }
            return result;
        }

        public static FileT ParseFileName(string fileName)
        {
            FileT file = new FileT();
            FileInfo fileNameInfo = new FileInfo(fileName);
            string[] fileComponents;
            fileComponents = fileNameInfo.Name.Split(ConfigFile.DELIMITER_NAME_FILE);
            int length = fileComponents.Count();
            string correct = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + "_OK.txt");
            string error = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + "_KO.txt");

            if (length != 5)
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension>.", 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "Formattazione del nome file errata.", true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
                return file = null;
            }
            if (!ConfigFile.DBS.Contains(fileComponents[3].ToUpper()))
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension> . DB specified not present.", 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "DB specificato nel nome file non previsto.", true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
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
                Logger.PrintLC(fileName + "produced an error while parsing its name: " + exp.Message, 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "Errore: " + exp.Message, true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
                return file = null;
            }
            return file;
        }
    }
}
