using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class FileOps
    {
        public static bool isFileOpenable(string file)
        {
            return ExcelOps.isFileOpenable(file);
        }

        public static List<string> GetTrueFilesToProcess(string[] list)
        {
            List<string> nlist = new List<string>();
            List<string> Direct = new List<string>();

            if (list != null)
            {
                bool notFullRecursive = true;
                if (!string.IsNullOrEmpty(ConfigFile.INPUT_FOLDER_NAME))
                {
                    notFullRecursive = true;
                    nlist = (from c in list
                             where c.Contains(ConfigFile.INPUT_FOLDER_NAME)
                             select c).ToList();
                }
                else
                {
                    notFullRecursive = false;
                    nlist = (from c in list
                             where c.Contains(ConfigFile.ROOT)
                             select c).ToList();
                }
                if (notFullRecursive)
                {
                    int pathLenght = ConfigFile.INPUT_FOLDER_NAME.Length;
                    foreach (string file in nlist)
                    {

                        try
                        {

                            FileInfo fileI = new FileInfo(file);
                            DirectoryInfo dir = fileI.Directory;
                            int dirLenght = dir.FullName.Length;
                            string padre = dir.FullName.Substring(dirLenght - pathLenght);
                            if (padre == ConfigFile.INPUT_FOLDER_NAME)
                                Direct.Add(file);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                else
                {
                    Direct = nlist;
                }
                if(Direct != null)
                    Direct = CleanDuplicates(Direct);
            }
            return Direct;
        }

        public static List<string> CleanDuplicates(List<string> list)
        {
            List<string> nlist = new List<string>();
            List<string> trueList = new List<string>();
            if (list != null)
            {
                foreach(var x in list)
                {
                    string XLS = Path.Combine(Path.GetDirectoryName(x), Path.GetFileNameWithoutExtension(x) + ".xls");
                    string XLSX = Path.Combine(Path.GetDirectoryName(x), Path.GetFileNameWithoutExtension(x) + ".xlsx");
                    if (!nlist.Contains(XLS) && !nlist.Contains(XLSX))
                    {
                        nlist.Add(x);
                    }
                }
                List<string> nameList = new List<string>(); //da aggiungere fuori dall'IF
                foreach (var elemento in nlist)
                {
                    if (!(nameList.Contains(Path.GetFileNameWithoutExtension(elemento))))
                    {
                        nameList.Add(Path.GetFileNameWithoutExtension(elemento));
                        trueList.Add(elemento);
                    }
                }
            }
            return trueList;
        }

        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }
        /// <summary>
        /// Removes a specific Attribute from a file.
        /// </summary>
        /// <param name="filePath">Path and file name to be elaborated</param>
        /// <param name="attribute">Attribute to be removed. 'ReadOnly' by default.</param>
        public static void RemoveAttributes(string filePath, FileAttributes attribute = FileAttributes.ReadOnly)
        {
            if (File.Exists(filePath))
            {
                FileAttributes attributes = File.GetAttributes(filePath);

                if ((attributes & attribute) == attribute)
                {
                    // Make the file RW
                    attributes = RemoveAttribute(attributes, attribute);
                    File.SetAttributes(filePath, attributes);
                    Logger.PrintLC(filePath + " is no longer RO.", 6, ConfigFile.INFO);
                }
            }
        }

        public static bool CopyFile(string originFile, string destinationFile)
        {
            if (File.Exists(originFile))
            {
                FileInfo fileOriginInfo = new FileInfo(originFile);
                FileInfo fileDestinationInfo = new FileInfo(destinationFile);
                try
                {
                    if (!Directory.Exists(fileDestinationInfo.DirectoryName))
                    {
                        Directory.CreateDirectory(fileDestinationInfo.DirectoryName);
                    }
                    RemoveAttributes(originFile);
                    if (File.Exists(destinationFile))
                        RemoveAttributes(destinationFile);
                    File.Copy(originFile, destinationFile, true);
                    Logger.PrintLC(originFile + " copied to " + 
                                   fileDestinationInfo.DirectoryName + " with the name: " + 
                                   fileDestinationInfo.Name, 2, ConfigFile.INFO);
                    return true;
                }
                catch(Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
        }


        public static bool CopyFile(string originFile, string destinationFile, bool bloccante)
        {
            if (File.Exists(originFile))
            {
                FileInfo fileOriginInfo = new FileInfo(originFile);
                FileInfo fileDestinationInfo = new FileInfo(destinationFile);
                try
                {
                    if (!Directory.Exists(fileDestinationInfo.DirectoryName))
                    {
                        Directory.CreateDirectory(fileDestinationInfo.DirectoryName);
                    }
                    RemoveAttributes(originFile);
                    if (File.Exists(destinationFile))
                        RemoveAttributes(destinationFile);
                    File.Copy(originFile, destinationFile, true);
                    Logger.PrintLC(originFile + " copied to " +
                                   fileDestinationInfo.DirectoryName + " with the name: " +
                                   fileDestinationInfo.Name, 2, ConfigFile.INFO);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
        }




        /// <summary>
        /// Legge tutte le righe del file specificato e restituisce una collezione di righe
        /// </summary>
        /// <param name="File"></param>
        /// <param name="ListaRigheSqlFile"></param>
        /// <returns></returns>
        public static bool LeggiFile(string File, ref List<string> ListaRigheSqlFile)
        {
            try
            {
                int counter = 0;
                string line;

                // Read the file and display it line by line.  
                System.IO.StreamReader file =
                    new System.IO.StreamReader(File);
                while ((line = file.ReadLine()) != null)
                {
                    ListaRigheSqlFile.Add(line);
                    counter++;
                }

                file.Close();
                
            }
            catch
            {
                return false;
            }
            return true;
        }


    }
}
