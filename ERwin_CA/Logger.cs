using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    static class Logger
    {
        private static string FileName;
        private static FileInfo FileInfos;
        //private static StreamWriter StrWr;
        private static string FileNameStream;
        public static void Initialize(string fileName)
        {
            Timer.SetFirstTime(DateTime.Now);
            FileName = fileName;
            FileInfos = new FileInfo(FileName);
            FileNameStream = FileInfos.DirectoryName + 
                             @"\" +
                             Path.GetFileNameWithoutExtension(FileInfos.FullName) + 
                             "_" +
                             Timer.GetTimestampDay(DateTime.Now) + 
                             ".txt";
            //StrWr = File.AppendText(FileNameStream);
        }
        public static void PrintL(string text)
        {
            string line = Timer.GetTimestampPrecision(DateTime.Now) + "    " + text;
            using ( StreamWriter StrWr = File.AppendText(FileNameStream))
            {
                StrWr.WriteLine(line);
            }
        }
        public static void PrintLC(string text)
        {
            string line = Timer.GetTimestampPrecision(DateTime.Now) + "    " + text;

            using (StreamWriter StrWr = File.AppendText(FileNameStream))
            {
                StrWr.WriteLine(line);
            }
        }
    }
}
