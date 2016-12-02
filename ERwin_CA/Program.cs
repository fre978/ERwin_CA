using ERwin_CA.T;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
//using Interop.ERXMLLib;
//using Interop.SCAPI;



namespace ERwin_CA
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.Initialize(ConfigFile.LOG_FILE);
            Logger.PrintL("AVVIO ESECUZIONE");
            ExcelOps Accesso = new ExcelOps();
            
            //string nomeFile = @"C:\ERWIN\CODICE\Extra\" + fileDaAprire.Name.ToString();
            //bool testBool = Accesso.ConvertXLStoXLSX(nomeFile);
            //testBool = ExcelOps.FileValidation(nomeFile);
            string[] ElencoExcel = DirOps.GetFilesToProcess(@"C:\ERWIN\CODICE\Extra\XLS\", "*.xls|*.xlsx");
            ConnMng connessione = new ConnMng();
            connessione.openModelConnection(ConfigFile.ERWIN_FILE);
            connessione.openTransaction();
            connessione.SetRootObject();
            connessione.SetRootCollection();
            foreach(var file in ElencoExcel)
            {
                if (ExcelOps.FileValidation(file))
                {
                    FileInfo fInfo = new FileInfo(file);
                    List<EntityT> DatiFile = ExcelOps.ReadXFile(fInfo);
                    foreach(var dati in DatiFile)
                        connessione.CreateEntity(dati);
                    Logger.PrintC("File " + file + " not valid for processing.");
                }
            }


            MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));

            
            //nomeFile = "";
            //SCAPI.Application testAPP = new SCAPI.Application();
            //if (fileDaAprire.Exists)
            //{
            //    using (ExcelPackage p = new ExcelPackage(fileDaAprire))
            //    {
            //        {
            //            //ExcelWorkbook WB = p.Workbook;
            //            //p.SaveAs(@"C:\nome.xls", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            //            ////WB.Worksheets
            //            //ExcelWorksheets ws = p.Workbook.Worksheets; //.Add(wsName + wsNumber.ToString());
            //            //foreach (var worksheet in ws)
            //            //{
            //            //    if (worksheet.Name == ConfigFile.FOGLIO01)
            //            //    {

            //            //    }
            //            //}
            //            //ws.Cells[1, 1].Value = wsName;
            //            //ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //            //ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            //            //ws.Cells[1, 1].Style.Font.Bold = true;
            //            //p.Save();
            //        }
            //    }
            //}
            Logger.PrintL("TERMINE ESECUZIONE");
            Timer.SetSecondTime(DateTime.Now);
            Logger.PrintL("Tempo esecuzione: " + Timer.GetTimeLapseFormatted(Timer.GetFirstTime(), Timer.GetSecondTime()) + Environment.NewLine);
        }
    }
}
