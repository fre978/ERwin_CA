﻿using System;
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
using ERwin_CA.T;

namespace ERwin_CA
{
    class ExcelOps
    {
        public static Excel.ApplicationClass ExApp = null;
        public ExcelOps()
        {
            ExApp = new Excel.ApplicationClass();
        }

        /// <summary>
        /// Converts a Open Office (xlsx) file to the proprietary MS old format (xls).
        /// -A.Amato, 2016 11
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public static bool ConvertXLSXtoXLS(string fileName = null)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;
            if (ExApp == null)
                return false;
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
                    Logger.PrintC("Error: " + exp.Message);
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
        /// -A.Amato, 2016 11
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public static bool ConvertXLStoXLSX(string fileName = null)
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
                    if (ExApp == null)
                        ExApp = new Excel.ApplicationClass();

                    FileOps.RemoveAttributes(fileName);
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fileName = Path.ChangeExtension(fileName, ".xlsx");
                    ExApp.DisplayAlerts = false;
                    FileInfo FileToSaveInfo = new FileInfo(fileName);
                    if (FileToSaveInfo.Exists)
                    {
                        FileToSaveInfo.Delete();
                    }
                    ExWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,//.xlOpenXMLStrictWorkbook,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Logger.PrintLC("File " + fileInfo.Name + " converted successfully to XLSX", 3);
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    //Marshal.FinalReleaseComObject(ExApp);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("File " + fileInfo.Name + " could not be converted to XLSX. Error: " + exp.Message, 3);
                    return false;
                }
            }
            else
                return false;
            return true;
        }

        public static bool FileValidation(string file)
        {
            //SCAPI.Application testAPP = new SCAPI.Application();
            string testoLog = string.Empty;
            FileInfo fileDaAprire = new FileInfo(file);
            if (fileDaAprire.Extension == ".xls")
            {
                if (!ConvertXLStoXLSX(file))
                    return false;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
                
           // if (file.EndsWith(".xls"))
            ExcelPackage p = new ExcelPackage(fileDaAprire);
            //using (ExcelPackage p = new ExcelPackage(fileDaAprire))
            //{
            //p.SaveAs(@"C:\nome.xls");
            //WB.Worksheets
            ExcelWorkbook WB = p.Workbook;
            
            ExcelWorksheets ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            bool sheetFound = false;
            bool columnsFound = false;
            int columns = 0;
            foreach (var worksheet in ws)
            {
                // SEZIONE TABELLE
                if (worksheet.Name == ConfigFile.TABELLE)
                {
                    columns = 0;
                    sheetFound = true;
                    columnsFound = false;
                    //List<string> dd = new List<string>();
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_TABELLE; 
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_TABELLE; 
                            columnsPosition++)
                    {   
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._TABELLE.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._TABELLE[value] != columnsPosition)
                                goto ERROR;
                            //dd.Add(worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text);
                        }
                        else
                        {
                            //worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Value = "";
                            testoLog = fileDaAprire.Name + ": file could not be elaborated.";
                            Logger.PrintLC(testoLog, 2);
                            goto ERROR;
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_TABELLE)
                        columnsFound = true;
                    else
                        goto ERROR;
                    //worksheet.Cell[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    //worksheet.Cells[1, 1].Style.Font.Bold = true;
                    //p.Save();
                }

                // SEZIONE ATTRIBUTI
                if (worksheet.Name == ConfigFile.ATTRIBUTI)
                {
                    columns = 0;
                    columnsFound = false;
                    sheetFound = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_ATTRIBUTI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._ATTRIBUTI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._ATTRIBUTI[value] != columnsPosition)
                                goto ERROR;
                        }
                        else
                        {
                            testoLog = fileDaAprire.Name + ": file could not be elaborated.";
                            Logger.PrintLC(testoLog, 2);
                            goto ERROR;
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_ATTRIBUTI)
                        columnsFound = true;
                    else
                        goto ERROR;
                }

                // SEZIONE RELAZIONI
                if (worksheet.Name == ConfigFile.RELAZIONI)
                {
                    columns = 0;
                    columnsFound = false;
                    sheetFound = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_RELAZIONI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_RELAZIONI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._RELAZIONI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._RELAZIONI[value] != columnsPosition)
                                goto ERROR;
                        }
                        else
                        {
                            testoLog = fileDaAprire.Name + ": file could not be elaborated.";
                            Logger.PrintLC(testoLog, 2);
                            goto ERROR;
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_RELAZIONI)
                        columnsFound = true;
                    else
                        goto ERROR;
                }
            }

            ERROR:
            WB.Dispose();
            p.Dispose();
            //MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));
            string fileError = Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(file) + "_KO.txt");
            string fileCorrect = Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(file) + "_OK.txt");
            if (File.Exists(fileError))
            {
                FileOps.RemoveAttributes(fileError);
                File.Delete(fileError);
            }
            if (File.Exists(fileCorrect))
            {
                FileOps.RemoveAttributes(fileCorrect);
                File.Delete(fileCorrect);
            }
            if (sheetFound != true || columnsFound != true)
            {
                Logger.PrintLC(fileDaAprire.Name + ": file NON idoneo all'elaborazione.", 2);
                Logger.PrintF(fileError, "File columns not formatted correctly.", true);
                return false;
            }
            Logger.PrintLC(fileDaAprire.Name + ": file IDONEO all'elaborazione.", 2);
            Logger.PrintF(fileCorrect, "File columns formatted correctly.", true);
            return true;
        }

        /// <summary>
        /// Reads and processes Table data from excel's 'TABELLE' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<EntityT> ReadXFileEntity(FileInfo fileDaAprire, string sheet = ConfigFile.TABELLE)
        {
            string file = fileDaAprire.FullName;
            List<EntityT> listaFile = new List<EntityT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Tables. File " + fileDaAprire.Name + " doesn't exist.", 3);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension == ".xls")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }

            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                p = new ExcelPackage(fileDaAprire);
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Reading Tables. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3);
                return listaFile = null;
            }
            
            bool FilesEnd = false;
            int EmptyRow = 0;
            int columns = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        string nome = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text;
                        string flag = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                        if (!string.IsNullOrWhiteSpace(nome) && (string.Equals(flag, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(flag, "N", StringComparison.OrdinalIgnoreCase))) 
                        {
                            EmptyRow = 0;
                            EntityT ValRiga = new EntityT(tName: nome);
                            ValRiga.TableName = nome;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text))
                                ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text))
                                ValRiga.HostName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text))
                                ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text))
                                ValRiga.Schema = worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text))
                                ValRiga.TableDescr = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text))
                                ValRiga.InfoType = worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text))
                                ValRiga.TableLimit = worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text))
                                ValRiga.TableGranularity = worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text))
                                ValRiga.FlagBFD = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                            listaFile.Add(ValRiga);
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Value = "OK";

                        }
                        else
                        {
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 2].Value = "Valore di NOME TABELLA e/o FLAG BFD non conformi.";

                            EmptyRow += 1;
                            if (EmptyRow >= 10)
                            {
                                FilesEnd = true;
                            }
                        }
                    }
                    p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    return listaFile;
                }
            }
            return listaFile = null;
        }

        /// <summary>
        /// Reads and processes Attributes data from excel's 'ATTRIBUTI' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<AttributeT> ReadXFileAttribute(FileInfo fileDaAprire, string sheet = ConfigFile.ATTRIBUTI)
        {
            string file = fileDaAprire.FullName;
            List<AttributeT> listaFile = new List<AttributeT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Attributes. File " + fileDaAprire.Name + " doesn't exist.", 2);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension == ".xls")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }

            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                p = new ExcelPackage(fileDaAprire);
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Reading Attributes. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 2);
                return listaFile = null;
            }

            bool FilesEnd = false;
            int EmptyRow = 0;
            int columns = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        string value = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text;
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            EmptyRow = 0;
                            AttributeT ValRiga = new AttributeT(nomeTabellaLegacy: value);
                            //ValRiga.TableName = value;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text))
                            //    ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text))
                            //    ValRiga.HostName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text))
                            //    ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text))
                            //    ValRiga.Schema = worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text))
                            //    ValRiga.TableDescr = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text))
                            //    ValRiga.InfoType = worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text))
                            //    ValRiga.TableLimit = worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text))
                            //    ValRiga.TableGranularity = worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text))
                            //    ValRiga.FlagBFD = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                            listaFile.Add(ValRiga);
                        }
                        else
                        {
                            EmptyRow += 1;
                            if (EmptyRow >= 10)
                                FilesEnd = true;
                        }
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }



    }
}
