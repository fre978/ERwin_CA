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
                        File.Delete(sFile);
                    }
                }
                catch (UnauthorizedAccessException)
                {

                }
                return AllFiles;
            }
            else
            {
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
    }
}
