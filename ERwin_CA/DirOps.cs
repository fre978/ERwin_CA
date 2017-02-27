using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class DirOps
    {

        public static bool TraverseDirectory(DirectoryInfo dirInfo)
        {
            foreach(var dirChild in dirInfo.GetDirectories())
            {
                try
                {
                    if (!TraverseDirectory(dirChild))
                    {
                        Logger.PrintLC("Could not traverse directory " + dirChild.FullName +
                            ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
                    }
                }
                catch
                {
                    Logger.PrintLC("Some error occured while trying to traverse " + dirChild.FullName + 
                        ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
                    continue;
                }
            }

            try
            {
                if (!CleanAllFilesInDirectory(dirInfo))
                {
                    Logger.PrintLC("Could not delete all files in directory " + dirInfo.FullName + ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
                }
            }
            catch
            {
                Logger.PrintLC("Some error occured while trying to delete all files in " + dirInfo.FullName + 
                    ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
            }

            try
            {
                if (dirInfo.GetFiles().Count() == 0)
                {
                    dirInfo.Delete();
                }
            }
            catch
            {
                Logger.PrintLC("Some error occured while trying to delete directory " + dirInfo.FullName + 
                    ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
            }

            return true;
        }

        public static bool CleanAllFilesInDirectory(DirectoryInfo dirInfo)
        {
            foreach(FileInfo file in dirInfo.GetFiles())
            {
                try
                {
                    file.IsReadOnly = false;
                    file.Delete();
                    System.Threading.Thread.Sleep(50);
                }
                catch
                {
                    Logger.PrintLC("Some error occured while trying to delete file " + file.FullName + 
                        ". Skipping it. Later on try to delete it manually.", 2, ConfigFile.ERROR);
                    continue;
                }
            }
            return true;
        }


        /// <summary>
        /// List all files of a determined type(s) in a directory tree
        /// </summary>
        /// <param name="homeDir">Root directory</param>
        /// <param name="fileType">Type of files to search. Multiple types are searchable with the '|' separator</param>
        /// <returns>String array of files paths</returns>
        public static string[] GetFilesToProcess(string homeDir, string fileType)
        {
            string[] AllFiles = new string[0];
            if (Directory.Exists(homeDir))
            {
                try
                {
                    AllFiles = GetFiles(homeDir, fileType, SearchOption.AllDirectories);
                    foreach (string sFile in AllFiles)
                    {
                        FileOps.RemoveAttributes(sFile);
                        //File.Delete(sFile);
                    }
                }
                catch (UnauthorizedAccessException)
                {

                }
                return AllFiles;
            }
            else
            {
                Logger.PrintLC("Search Folder " + homeDir + " not exists, no files found");
                return AllFiles;
            }
        }

        /// <summary>
        /// Lists all files of searchPattern type in a rootDir tree directory.
        /// </summary>
        /// <param name="rootDir">Root directory for the search</param>
        /// <param name="searchPattern">Pattern or types of files to search</param>
        /// <param name="searchOption">SearchOption type of search (AllDirectories or TopDirectoryOnly)</param>
        /// <returns>String array of files paths</returns>
        public static string[] GetFiles(string rootDir, string searchPattern, 
                                            SearchOption searchOption = SearchOption.AllDirectories)
        {
            string[] Patterns = searchPattern.Split('|');
            List<string> Files = new List<string>();
            foreach(string Patt in Patterns)
            {
                string PattClean = Patt.Replace("*", "");
                Files.AddRange(Directory.EnumerateFiles(rootDir, "*", searchOption)
                        .Where(s => s.EndsWith(PattClean, StringComparison.OrdinalIgnoreCase)));
            }
            return Files.ToArray();
        }

        //#############################################################################
        public static bool Copy(string sourceDirectory, string targetDirectory, List<string> list)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            try
            {
                return CopyAll(diSource, diTarget, list);
            }
            catch
            {
                Logger.PrintLC("Error while copying files and directories to the destination folder: " + targetDirectory, 1, ConfigFile.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Generic methods to copy recursively directories and files 
        /// </summary>
        /// <param name="sourceDirectory">Source directory</param>
        /// <param name="targetDirectory">Destination directory</param>
        /// <returns>Bool result of operation</returns>
        public static bool Copy(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            try
            {
                return CopyAll(diSource, diTarget);
            }
            catch
            {
                Logger.PrintLC("Error while copying files and directories to the destination folder: " + targetDirectory, 1, ConfigFile.ERROR);
                return false;
            }
        }

        public static bool CopyAll(DirectoryInfo source, DirectoryInfo target, List<string> list)
        {
            try
            {
                Directory.CreateDirectory(target.FullName);
                int temp = source.GetFiles().Count();
                List<FileInfo> tempList = source.GetFiles().ToList();
                // Copy each file into the new directory.
                foreach (FileInfo fi in source.GetFiles())
                {
                    if (list.Contains(fi.FullName))
                    {
                        try
                        {
                            Logger.PrintLC("Copying " + Path.Combine(target.FullName, fi.Name), 3);
                            fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                        }
                        catch
                        {
                            Logger.PrintLC("Could not copy file " + fi.FullName + ". Skipping it.", 1, ConfigFile.WARNING);
                            continue;
                        }
                    }
                }

                // Copy each subdirectory using recursion.
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    try
                    {
                        DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                        CopyAll(diSourceSubDir, nextTargetSubDir, list);
                    }
                    catch
                    {
                        Logger.PrintLC("Could not copy directory " + diSourceSubDir.FullName + ". Skipping it.", 2, ConfigFile.WARNING);
                        continue;
                    }
                }
                return true;
            }
            catch
            {
                Logger.PrintLC("Error while copying files and directories to the destination location", 1, ConfigFile.ERROR);
                return false;
            }
        }


        public static bool CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            try
            {
                Directory.CreateDirectory(target.FullName);

                // Copy each file into the new directory.
                foreach (FileInfo fi in source.GetFiles())
                {
                    try
                    {
                        Logger.PrintLC("Copying " + Path.Combine(target.FullName,  fi.Name), 3);
                        fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                    }
                    catch
                    {
                        Logger.PrintLC("Could not copy file " + fi.FullName + ". Skipping it.", 1, ConfigFile.WARNING);
                        continue;
                    }
                }

                // Copy each subdirectory using recursion.
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    try
                    {
                        DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                        CopyAll(diSourceSubDir, nextTargetSubDir);
                    }
                    catch
                    {
                        Logger.PrintLC("Could not copy directory " + diSourceSubDir.FullName + ". Skipping it.", 2, ConfigFile.WARNING);
                        continue;
                    }
                }
                return true;
            }
            catch
            {
                Logger.PrintLC("Error while copying files and directories to the destination location", 1, ConfigFile.ERROR);
                return false;
            }
        }
        //#############################################################################

    }
}
