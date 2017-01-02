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
        public static List<string> GetTrueFilesToProcess(string[] list)
        {
            List<string> nlist = new List<string>();
            if (list != null)
            {
                //List<string> nlist = list.ToList();
                //list = list.Where(x => !list.Contains(ConfigFile.FOLDERDESTINATION, IEqualityComparer));
                nlist = (from c in list
                         where !c.Contains(ConfigFile.FOLDERDESTINATION_GENERAL)
                         select c).ToList();
            }
            nlist = CleanDuplicates(nlist);
            return nlist;
        }

        public static List<string> CleanDuplicates(List<string> list)
        {
            List<string> nlist = new List<string>();
            List<string> trueList = new List<string>();
            if (list != null)
            {
                //foreach(var x in list)
                //{
                //    nlist.Add(Path.GetDirectoryName(x) + Path.GetFileNameWithoutExtension(x));
                //}
                //nlist = nlist.Distinct().ToList();

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
                    Logger.PrintLC(filePath + " is no longer RO.", 2);
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
                                   fileDestinationInfo.Name, 2);
                    return true;
                }
                catch(Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message, 2);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.", 2);
                return false;
            }
                
        }


    }
}
