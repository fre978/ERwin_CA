using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ERwin_CA
{
    class ExcelOps
    {
        public Excel.ApplicationClass ExApp;
        public ExcelOps()
        {
            ExApp = new Excel.ApplicationClass();
        }
        /// <summary>
        /// Converts a Open Office (xlsx) file to the proprietary MS old format (xls).
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public bool ConvertXLSXtoXLS(string fileName = null)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                return false;
            }
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists && (fileInfo.Extension == ".xlsx"))
            {
                //Excel.ApplicationClass ExApp = new Excel.ApplicationClass();
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();

                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    fileName = Path.ChangeExtension(fileName, ".xls"); //.Replace(".xlsx", ".xls");
                    ExApp.DisplayAlerts = false;
                    FileInfo FileToSaveInfo = new FileInfo(fileName);
                    if (FileToSaveInfo.Exists)
                    {
                        FileToSaveInfo.Delete();
                    }
                    ExWB.SaveAs(fileName, Excel.XlFileFormat.xlExcel8,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    Marshal.FinalReleaseComObject(ExApp);
                }
                catch (Exception exp)
                {
                    return false;
                }
                finally
                {

                }
            }
            return true;
        }
        /// <summary>
        /// Converts a proprietary MS old format (xls) to the Open Office (xlsx).
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public bool ConvertXLStoXLSX(string fileName = null)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists && (fileInfo.Extension == ".xls"))
            {
                //Excel.ApplicationClass ExApp = new Excel.ApplicationClass();
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();

                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fileName = fileName.Replace(".xls", ".xlsx");
                    ExApp.DisplayAlerts = false;
                    FileInfo FileToSaveInfo = new FileInfo(fileName);
                    if (FileToSaveInfo.Exists)
                    {
                        FileToSaveInfo.Delete();
                    }
                    ExWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLStrictWorkbook,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    Marshal.FinalReleaseComObject(ExApp);
                }
                catch (Exception exp)
                {
                    return false;
                }
                finally
                {

                }
            }
            return true;
        }

        public static bool FileValidation(string file)
        {
            //SCAPI.Application testAPP = new SCAPI.Application();
            FileInfo fileDaAprire = new FileInfo(file);
            ExcelPackage p = new ExcelPackage(fileDaAprire);
            //using (ExcelPackage p = new ExcelPackage(fileDaAprire))
            //{
            ExcelWorkbook WB = p.Workbook;
            //p.SaveAs(@"C:\nome.xls");
            //WB.Worksheets
            ExcelWorksheets ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            bool sheetFound = false;
            bool columnsFound = false;
            int columns = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == ConfigFile.FOGLIO01)
                {
                    sheetFound = true;
                    List<string> dd = new List<string>();
                    for (int x = ConfigFile.HEADER_COLONNA_MIN; x < ConfigFile.HEADER_COLONNA_MAX; x++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, x].Text;
                        if (ConfigFile._TABELLE.ContainsKey(value))
                        {
                            columns += 1;
                            //int position = ConfigFile._TABELLE[value];
                            if (ConfigFile._TABELLE[value] != x)
                            {
                                columnsFound = false;
                                return false;
                            }
                            dd.Add(worksheet.Cells[ConfigFile.HEADER_RIGA, x].Text);
                        }
                        else
                            return false;
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE)
                        columnsFound = true;
                    else
                        return false;

                    p.Dispose();
                    //worksheet.Cell[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    //worksheet.Cells[1, 1].Style.Font.Bold = true;
                    //p.Save();
                }
            }
            Marshal.FinalReleaseComObject(p);
            Marshal.FinalReleaseComObject(WB);
            Marshal.FinalReleaseComObject(ws);
            if (sheetFound != true)
            {
                return false;
            }
            return true;
            //}
        }

        //public ExcelWorksheet
    }
}
