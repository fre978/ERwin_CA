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
        object testOBJ = new SCAPI.ModelObjects();
        static void Main(string[] args)
        {
            //string[] testFiles = DirOps.GetFilesToProcess(@"C:\ROOTtest\", "*.mpp|*.txt|*.zip|*.xls|.xlsx");

            FileInfo fileDaAprire = new FileInfo(ConfigFile.FILETEST);
            string nomeFile = @"C:\ERWIN\CODICE\Extra\" + fileDaAprire.Name.ToString();
            ExcelOps Accesso = new ExcelOps();
            bool testBool = Accesso.ConvertXLStoXLSX(nomeFile);
            testBool = ExcelOps.FileValidation(nomeFile = nomeFile.Replace("xls","xlsx"));
            //nomeFile = "";
            //SCAPI.Application testAPP = new SCAPI.Application();
            if (fileDaAprire.Exists)
            {
                using (ExcelPackage p = new ExcelPackage(fileDaAprire))
                {
                    {
                        //ExcelWorkbook WB = p.Workbook;
                        //p.SaveAs(@"C:\nome.xls", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                        ////WB.Worksheets
                        //ExcelWorksheets ws = p.Workbook.Worksheets; //.Add(wsName + wsNumber.ToString());
                        //foreach (var worksheet in ws)
                        //{
                        //    if (worksheet.Name == ConfigFile.FOGLIO01)
                        //    {

                        //    }
                        //}
                        //ws.Cells[1, 1].Value = wsName;
                        //ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                        //ws.Cells[1, 1].Style.Font.Bold = true;
                        //p.Save();
                    }
                }
            }
            
        }
    }
}
