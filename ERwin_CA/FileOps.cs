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
                    Logger.PrintLC(filePath + " is no longer RO.");
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
                    File.Copy(originFile, destinationFile, true);
                    Logger.PrintLC(originFile + " copied to " + fileDestinationInfo.DirectoryName + " with the name: " + fileDestinationInfo.Name);
                    return true;
                }
                catch(Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.");
                return false;
            }
                
        }


    }
}
