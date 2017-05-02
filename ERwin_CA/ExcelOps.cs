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
        public static Excel.ApplicationClass ExApp = null;
        public ExcelOps()
        {
            ExApp = new Excel.ApplicationClass();
        }

        public static bool isFileOpenable (string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists)
            {
                ExApp = new Excel.ApplicationClass();
                Excel.Worksheet ExWS = new Excel.Worksheet();
                Excel.Workbook ExWB = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    ExWB = null;
                    FileOps.RemoveAttributes(fileName);
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    if (fileInfo.Extension.ToUpper() == ".XLS")
                    {
                        ExWB.SaveAs(fileName, Excel.XlFileFormat.xlExcel8,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                    }
                    else
                    {
                        ExWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,//.xlOpenXMLStrictWorkbook,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                    }

                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    Marshal.FinalReleaseComObject(ExApp);
                }
                catch (Exception ex)
                {
                    try
                    {
                        if (ExWB != null)
                        {
                            ExWB.Close();
                            Marshal.FinalReleaseComObject(ExWB);
                        }
                        if (ExWS != null)
                        {
                            Marshal.FinalReleaseComObject(ExWS);
                        }
                        if (ExApp != null)
                        {
                            Marshal.FinalReleaseComObject(ExApp);
                        }
                        Logger.PrintLC("Errore: " + ex.Message, 2, ConfigFile.ERROR);
                        return false;
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Converts an Open Office (.xlsx) file to the proprietary MS old format (.xls).
        /// -A.Amato, 2016 11
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public static bool ConvertXLSXtoXLS(string fileName = null)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;

            //if (ExApp == null)
            //    return false;

            ExApp = new Excel.ApplicationClass();
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists && (fileInfo.Extension.ToUpper() == ".XLSX"))
            {
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();
                    ExApp.DisplayAlerts = false;
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    fileName = Path.ChangeExtension(fileName, ".xls"); //.Replace(".xlsx", ".xls");
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
                    Logger.PrintLC("Successfully converted " + fileInfo.FullName + " to " + fileName, 2, ConfigFile.INFO);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Error converting XLSX to XLS: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Converts a proprietary MS old format (.xls) to the Open Office (.xlsx).
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
            if (fileInfo.Exists && (fileInfo.Extension.ToUpper() == ".XLS"))
            {
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();
                    ExApp = new Excel.ApplicationClass();
                    ExApp.DisplayAlerts = false;
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fileName = Path.ChangeExtension(fileName, ".xlsx");
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
                    Logger.PrintLC("File " + fileInfo.Name + " converted successfully to XLSX", 3, ConfigFile.INFO);
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
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
            int genericError = 0;
            string testoLog = string.Empty;
            string TxtControlloNonPassato = string.Empty;
            bool sheetFoundTabelle = false;
            bool sheetFoundAttributi = false;
            bool sheetFoundRelazioni = false;
            bool columnsFoundTabelle = false;
            bool columnsFoundAttributi = false;
            bool columnsFoundRelazioni = false;
            int columns = 0;
            int[] check_sheet = new int[3] { 0, 0, 0 };
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            FileInfo fileDaAprire = new FileInfo(file);
            bool isXLS = false;

            //SEZIONE TEST D'APERTURA
            try
            {
                //test se il file è un temporaneo
                char[] opened = fileDaAprire.Name.ToCharArray();
                if (opened[0] == '~')
                {
                    Logger.PrintLC(fileDaAprire.Name + " è un file temporaneo (probabilmente è già aperto altrove). Non elaboro.", 2, ConfigFile.ERROR);
                    genericError = 1;
                    goto ERROR;
                }

                //test se il file apribile
                if (!FileOps.isFileOpenable(file))
                {
                    Logger.PrintLC("Non è possibile aprire il file " + fileDaAprire.Name + ". Non elaboro.", 2, ConfigFile.ERROR);
                    genericError = 2;
                    goto ERROR;
                }
                //string extension = fileDaAprire.Extension.ToUpper();
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                    {
                        if (!ConvertXLStoXLSX(file))
                        {
                            genericError = 3;
                            goto ERROR;
                        }
                    }
                    isXLS = true;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
            }
            catch
            {
                genericError = 4;
                goto ERROR;
            }

            try
            {
                ExApp = new Excel.ApplicationClass();
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
            }
            catch
            {
                try
                {
                    ExApp.DisplayAlerts = false;
                    Logger.PrintLC(fileDaAprire.Name + " già aperto da un'altra applicazione. Chiudere e riprovare.", 2, ConfigFile.ERROR);
                }
                catch { }
                return false;
            }
            WB = p.Workbook;
            ws = WB.Worksheets;

            foreach (var worksheet in ws)
            {
                // SEZIONE TABELLE
                if (worksheet.Name == ConfigFile.TABELLE)
                {
                    columns = 0;
                    check_sheet[0] += 1;
                    sheetFoundTabelle = true;
                    columnsFoundTabelle = false;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_TABELLE; 
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_TABELLE; 
                            columnsPosition++)
                    {   
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._TABELLE.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._TABELLE[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non trovato alla colonna " + columnsPosition + " del Foglio " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                                value = "[Campo Senza Nome, posizione: " + columnsPosition + "]";
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non è una colonna valida del Foglio " + worksheet.Name;
                            Logger.PrintLC(fileDaAprire.Name + ": " + value + " non è una colonna valida del Foglio " + worksheet.Name + ". Il file non può essere elaborato.", 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_TABELLE)
                        columnsFoundTabelle = true;
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tNumero colonne nel Foglio " + worksheet.Name + " non corretto." + worksheet.Name;
                    }
                }

                // SEZIONE ATTRIBUTI
                if (worksheet.Name == ConfigFile.ATTRIBUTI)
                {
                    check_sheet[1] += 1;
                    columns = 0;
                    columnsFoundAttributi = false;
                    sheetFoundAttributi = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_ATTRIBUTI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._ATTRIBUTI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._ATTRIBUTI[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non trovato alla colonna " + columnsPosition + " del Foglio " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                            {
                                value = "[Campo Senza Nome, posizione: " + columnsPosition + "]";
                            }
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non è una colonna valida del Foglio " + worksheet.Name;
                            Logger.PrintLC(fileDaAprire.Name + ": " + value + " non è una colonna valida del Foglio " + worksheet.Name + ". Il file non può essere elaborato.", 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_ATTRIBUTI)
                    {
                        columnsFoundAttributi = true;
                    }
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tNumero colonne nel Foglio " + worksheet.Name + " non corretto." + worksheet.Name;
                    }
                }

                // SEZIONE RELAZIONI
                if (worksheet.Name == ConfigFile.RELAZIONI)
                {
                    check_sheet[2] += 1;
                    columns = 0;
                    columnsFoundRelazioni = false;
                    sheetFoundRelazioni = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_RELAZIONI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_RELAZIONI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._RELAZIONI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._RELAZIONI[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non trovato alla colonna " + columnsPosition + " del Foglio " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                            {
                                value = "[Campo Senza Nome, posizione: " + columnsPosition + "]";
                            }
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t" + value + " non è una colonna valida del Foglio " + worksheet.Name;
                            testoLog = fileDaAprire.Name + ": " + value + " non è una colonna valida del Foglio " + worksheet.Name + ". Il file non può essere elaborato.";
                            Logger.PrintLC(testoLog, 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_RELAZIONI)
                    {
                        columnsFoundRelazioni = true;
                    }
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tNumero colonne nel Foglio " + worksheet.Name + " non corretto." + worksheet.Name;
                    }
                }
            }

            ERROR:
            try
            {
                WB.Dispose();
                p.Dispose();
            }
            catch
            {
            }

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
            string fileStampa = String.Empty;


            if (genericError != 0)
            {
                Logger.PrintF(fileError, "er_driveup – Caricamento Excel su ERwin", true);
                switch (genericError)
                {
                    case 1:
                        Logger.PrintF(fileError, "Il file è temporaneo. Non può essere elaborato.", true);
                        break;
                    case 2:
                        Logger.PrintF(fileError, "Il file non è apribile. Potrebbe essere già aperto altrove. Non può essere elaborato.", true);
                        break;
                    case 3:
                        Logger.PrintF(fileError, "Non è stato possibile convertire il file con estensione '.XLSX'. Non può essere elaborato.", true);
                        break;
                    case 4:
                        Logger.PrintF(fileError, "Si è verificato un errore non previsto durante l'apertura del file. Non può essere elaborato", true);
                        break;
                }
                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
                return false;
            }

            if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1 ||
                sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                fileStampa = fileError;
            }
            else
            {
                fileStampa = fileCorrect;
            }

            Logger.PrintF(fileStampa, "er_driveup – Caricamento Excel su ERwin", true);

            if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1)
            {
                Logger.PrintLC(fileDaAprire.Name + ": non può essere elaborato: uno dei Fogli non è presente o una delle colonne non è conforme", 2, ConfigFile.ERROR);
                Logger.PrintF(fileStampa, fileDaAprire.Name + ": non può essere elaborato: uno o più Fogli non sono presenti:", true);
                if (check_sheet[0] != 1)
                {
                    Logger.PrintLC("\t\tFoglio Censimento Tabelle non presente.");
                    Logger.PrintF(fileStampa, "Foglio Censimento Tabelle non presente.", true);
                }
                if (check_sheet[1] != 1)
                {
                    Logger.PrintLC("\t\tFoglio Censimento Attributi non presente.");
                    Logger.PrintF(fileStampa, "Foglio Censimento Attributi non presente.", true);
                }
                if (check_sheet[2] != 1)
                {
                    Logger.PrintLC("\t\tFoglio Relazioni-ModelloDatiLegacy non presente.");
                    Logger.PrintF(fileStampa, "Foglio Relazioni-ModelloDatiLegacy non presente.", true);
                }

                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
            }
            if (sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                Logger.PrintLC(fileDaAprire.Name + ": file could not be processed: Columns or Sheets are not in the expected format.", 2, ConfigFile.ERROR);
                Logger.PrintF(fileStampa, "Colonne o Fogli non formattati correttamente:", true);
                string[] val = TxtControlloNonPassato.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                if (val.Count() > 0)
                {
                    foreach (string valC in val)
                    {
                        if (!string.IsNullOrWhiteSpace(valC))
                            Logger.PrintF(fileStampa, valC, true);
                    }
                }
                else
                {
                    Logger.PrintF(fileStampa, TxtControlloNonPassato, true);
                }
            }

            if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1 ||
                sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
                return false;
            }

            Logger.PrintLC(fileDaAprire.Name + ": file valid to be processed.", 2, ConfigFile.INFO);
            Logger.PrintF(fileStampa, "Colonne e Fogli formattati corretamente.", true);
            return true;
        }


        public static bool TestAttributesExist(ExcelWorksheets ws, string nome)
        {
            bool attrExist = false;
            foreach (var sheetAtt in ws)
            {
                if (sheetAtt.Name == ConfigFile.ATTRIBUTI)
                {
                    bool FilesEndAtt = false;
                    for (int RowPosAtt = ConfigFile.HEADER_RIGA + 1;
                            FilesEndAtt != true;
                            RowPosAtt++)
                    {
                        string nomeAtt = sheetAtt.Cells[RowPosAtt, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text.Trim();
                        if (!string.IsNullOrWhiteSpace(nomeAtt))
                        {
                            if (nome == nomeAtt)
                            {
                                attrExist = true;
                                break;
                            }
                        }
                        else
                        {

                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossimeAtt = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(sheetAtt.Cells[RowPosAtt + i, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text))
                                prossimeAtt++;
                        }
                        if (prossimeAtt == 10)
                            FilesEndAtt = true;
                        //******************************************
                    }
                    break;
                }
            }
            return attrExist;
        }


        /// <summary>
        /// Reads and processes Table data from excel's 'TABELLE' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<EntityT> ReadXFileEntity(FileInfo fileDaAprire, string db, string sheet = ConfigFile.TABELLE)
        {
            string file = fileDaAprire.FullName;
            List<EntityT> listaFile = new List<EntityT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Tables. File " + fileDaAprire.Name + " doesn't exist.", 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                {
                    return listaFile = null;
                }
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Reading Tables. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            
            bool FilesEnd = false;
            int EmptyRow = 0;
            //int columns = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string error = string.Empty;
                        string nome = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text;
                        if (listaFile.Exists(x => x.TableName == nome))
                        {
                            incorrect = true;
                            error += "Una tabella con lo stesso NOME TABELLA è già presente. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        // #################################
                        // TEST ESISTENZA ATTRIBUTI PER LA TABELLA
                        bool attrExist = TestAttributesExist(ws, nome);     // 'attributes exist for table' flag
                        if(attrExist == false)
                        {
                            incorrect = true;
                            error += "La Tabella non possiede Attributi; non verrà inserita. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        }
                        // #################################

                        string flag = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                        if (string.IsNullOrWhiteSpace(nome))
                        {
                            incorrect = true;
                            error += "Valore di NOME TABELLA mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        if (!(Funct.ParseFlag(flag, "YES") || Funct.ParseFlag(flag, "NO")))
                        {
                            incorrect = true;
                            error += "Valore di FLAG BFD non conforme. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        // CODE 66
                        if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim()))
                        {
                            string databaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim(); //ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim();
                            if (!Funct.ValidateDatabaseName(databaseName))
                            {
                                incorrect = true;
                                error += "Valore di NOME DATABASE non conforme. ";
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                                worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            }
                        }
                        else
                        {
                            incorrect = true;
                            error += "Valore di NOME DATABASE mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        }

                        if (incorrect == false)
                        { 
                            EmptyRow = 0;
                            EntityT ValRiga = new EntityT(row: RowPos, db: db, tName: nome);
                            ValRiga.TableName = nome;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text.Trim()))
                                ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text.Trim()))
                                ValRiga.HostName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text.Trim()))
                                ValRiga.Schema = worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text.Trim()))
                                ValRiga.TableDescr = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text.Trim()))
                                ValRiga.InfoType = worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text.Trim()))
                                ValRiga.TableLimit = worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text.Trim()))
                                ValRiga.TableGranularity = worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text.Trim()))
                            {
                                if (Funct.ParseFlag(flag, "YES"))
                                    ValRiga.FlagBFD = "S";
                                if (Funct.ParseFlag(flag, "NO"))
                                    ValRiga.FlagBFD = "N";
                            }
                            else
                                ValRiga.FlagBFD = "N";

                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim()))
                                ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim();

                            listaFile.Add(ValRiga); // aggiunta alle tabelle da inserire nel modello ERWIN

                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Nome Tabella"]].Text) && string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Flag BFD"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Table '" + worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Table '" + worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text + "'. Validation OK", 3, ConfigFile.INFO);
                        }
                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(new FileInfo(Funct.GetFolderDestination2(fileDaAprire.FullName, fileDaAprire.Name)));
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }

        /// <summary>
        /// Reads and processes Table data from excel's 'TABELLE' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<RelationT> ReadXFileRelation(FileInfo fileDaAprire, string db, string sheet = ConfigFile.RELAZIONI)
        {
            string file = fileDaAprire.FullName;
            List<RelationT> listaFile = new List<RelationT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Tables. File " + fileDaAprire.Name + " doesn't exist.", 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Reading Relation. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            
            bool FilesEnd = false;
            int EmptyRow = 0;
            
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string error = null;
                        string identificativoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativo relazione"]].Text.ToUpper().Trim();
                        string tabellaPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Padre"]].Text.ToUpper().Trim();
                        string tabellaFiglia = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Figlia"]].Text.ToUpper().Trim();
                        string cardinalita = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Cardinalità"]].Text.ToUpper().Trim();
                        string campoPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Padre"]].Text.ToUpper().Trim();
                        string campoFiglio = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Figlio"]].Text.ToUpper().Trim();
                        string identificativa = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativa"]].Text.ToUpper().Trim();
                        string eccezione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text.ToUpper().Trim();
                        string tipoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tipo Relazione"]].Text.ToUpper().Trim();
                        string note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text.ToUpper().Trim();

                        if (listaFile.Exists(x => x.IdentificativoRelazione == identificativoRelazione &&
                                                  x.TabellaPadre == tabellaPadre &&
                                                  x.TabellaFiglia == tabellaFiglia &&
                                                  x.CampoPadre == campoPadre &&
                                                  x.CampoFiglio == campoFiglio)
                                                  )
                        {
                            incorrect = true;
                            error += "Relazione già presente con ID: " + identificativoRelazione + " Tabella Padre: " + tabellaPadre + " Tabella Figlia: " + tabellaFiglia + " Campo Padre: " + campoPadre + " Campo Figlia: " + campoFiglio;
                        }
                            if (string.IsNullOrWhiteSpace(identificativoRelazione))
                        {
                            incorrect = true;
                            error += "IDENTIFICATIVO RELAZIONE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(tabellaPadre))
                        {
                            incorrect = true;
                            error += "TABELLA PADRE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(tabellaFiglia))
                        {
                            incorrect = true;
                            error += "TABELLA FIGLIA mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(cardinalita))
                        {
                            incorrect = true;
                            error += "CARDINALITA mancante. ";
                        }
                        else
                        {
                            switch (cardinalita.ToUpper())
                            { 
                                case "1:1":
                                    break;
                                case "1:N":
                                    break;
                                case "N:N":
                                    break;
                                case "(0,1) A (0,1)":
                                    break;
                                case "(0,1) A (1,M)":
                                    break;
                                case "(0,1) A (0,1,M)":
                                    break;
                                case "1 A (0,1)":
                                    break;
                                case "1 A (1,M)":
                                    break;
                                case "1 A (0,1,M)":
                                    break;
                                default:
                                    incorrect = true;
                                    error += "CARDINALITA non conforme. ";
                                    break;
                            }
                        }
                        if (string.IsNullOrWhiteSpace(campoPadre))
                        {
                            incorrect = true;
                            error += "CAMPO PADRE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(campoFiglio))
                        {
                            incorrect = true;
                            error += "CAMPO FIGLIO mancante. ";
                        }
                        if (!string.IsNullOrWhiteSpace(identificativa))
                        {
                            if(!(Funct.ParseFlag(identificativa.ToUpper(),"YES") || Funct.ParseFlag(identificativa.ToUpper(),"NO")))
                            {
                                incorrect = true;
                                error += "IDENTIFICATIVA non conforme. ";
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(tipoRelazione))
                        {
                            string upperTipoRelazione = tipoRelazione.ToUpper();
                            if (!(upperTipoRelazione.Equals("L") || upperTipoRelazione.Equals("LOGICA") ||
                                upperTipoRelazione.Equals("F") || upperTipoRelazione.Equals("FISICA")))
                            {
                                incorrect = true;
                                error += "TIPO RELAZIONE non conforme";
                            } 
                        }

                        if (incorrect == false)
                        { 
                            EmptyRow = 0;
                            RelationT ValRiga = new RelationT(row: RowPos, db: db);
                            ValRiga.IdentificativoRelazione = identificativoRelazione;
                            ValRiga.TabellaPadre = tabellaPadre;
                            ValRiga.TabellaFiglia = tabellaFiglia;
                            switch (cardinalita.ToUpper())
                            {
                                case "1:1":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1:N":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "N:N":
                                    ValRiga.History = "CARDINALITA non gestita dall'applicazione";
                                    ValRiga.NullOptionType = null;
                                    break;
                                case "(0,1) A (0,1)":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "(0,1) A (1,M)":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "(0,1) A (0,1,M)":
                                    ValRiga.Cardinalita = -3;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "1 A (0,1)":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1 A (1,M)":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1 A (0,1,M)":
                                    ValRiga.Cardinalita = -3;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                default:
                                    ValRiga.History = "CARDINALITA non conforme";
                                    ValRiga.NullOptionType = null;
                                    break;
                            }
                            ValRiga.CampoPadre = campoPadre;
                            ValRiga.CampoFiglio = campoFiglio;
                            if (Funct.ParseFlag(identificativa.ToUpper(),"YES"))
                                ValRiga.Identificativa = 2;
                            else
                                ValRiga.Identificativa = 7;

                            if (string.IsNullOrEmpty(tipoRelazione))
                            {
                                ValRiga.TipoRelazione = true;
                            }
                            else
                            {
                                switch (tipoRelazione.ToUpper())
                                {
                                    case "L":
                                        ValRiga.TipoRelazione = true;
                                        break;
                                    case "LOGICA":
                                        ValRiga.TipoRelazione = true;
                                        break;
                                    case "F":
                                        ValRiga.TipoRelazione = false;
                                        break;
                                    case "FISICA":
                                        ValRiga.TipoRelazione = false;
                                        break;
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text))
                                ValRiga.Eccezioni = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text))
                                ValRiga.Note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text;
                            listaFile.Add(ValRiga);
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Nome Tabella"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Relation '" + identificativoRelazione + "' between Table '" + tabellaPadre + "' and Table '"+ tabellaFiglia + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Relation '" + identificativoRelazione + "' between Table '" + tabellaPadre + "' and Table '" + tabellaFiglia + "'. Validation OK", 3, ConfigFile.INFO);
                        }

                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(fileDaAprire);
                    }
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
        public static List<AttributeT> ReadXFileAttribute(FileInfo fileDaAprire, string db, string sheet = ConfigFile.ATTRIBUTI)
        {
            string file = fileDaAprire.FullName;
            List<AttributeT> listaFile = new List<AttributeT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Attributes. File " + fileDaAprire.Name + " doesn't exist.", 2, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Reading Attributes. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 2, ConfigFile.ERROR);
                return listaFile = null;
            }

            bool FilesEnd = false;
            int EmptyRow = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string nomeTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text.ToUpper().Trim();
                        string nomeCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text.ToUpper().Trim();
                        if (nomeCampo.Contains("-"))
                        {
                            nomeCampo = nomeCampo.Replace("-", "_");
                            Logger.PrintLC("Field '" + worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text + "' of Table '" + nomeTabella + "' has been renamed as " + nomeCampo + ". This value will be used to produce the Erwin file", 3, ConfigFile.WARNING);
                        }
                        string dataType = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Datatype"]].Text.Trim();
                        dataType = Funct.RemoveWhitespace(dataType);
                        string chiave = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave"]].Text.ToUpper().Trim();
                        string unique = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Unique"]].Text.ToUpper().Trim();
                        string chiaveLogica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave Logica"]].Text.ToUpper().Trim();
                        string mandatoryFlag = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Mandatory Flag"]].Text.ToUpper().Trim();
                        string dominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dominio"]].Text.ToUpper().Trim();
                        string storica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Storica"]].Text.ToUpper().Trim();
                        string datoSensibile = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dato Sensibile"]].Text.ToUpper().Trim();

                        worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "";
                        worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = "";

                        string error = "";
                        //Check Nome Tabella Legacy
                        if (string.IsNullOrWhiteSpace(nomeTabella))
                        {
                            incorrect = true;
                            error += "NOME TABELLA LEGACY mancante.";

                        }
                        //Check Nome Campo Legacy
                        if (string.IsNullOrWhiteSpace(nomeCampo))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "NOME CAMPO LEGACY mancante.";
                        }
                        //Check DataType
                        if (string.IsNullOrWhiteSpace(dataType))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DATATYPE mancante.";
                        }
                        else
                        {
                            if (!Funct.ParseDataType(dataType, db))
                            {
                                incorrect = true;
                                if (!string.IsNullOrWhiteSpace(error))
                                    error += " ";
                                error += "DATATYPE non conforme.";
                            }
                        }
                        //Check Chiave
                        if (!(string.IsNullOrWhiteSpace(chiave)) && (!(Funct.ParseFlag(chiave, "YES") || Funct.ParseFlag(chiave, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "CHIAVE non conforme.";
                        }
                        //Check Unique
                        if (!(string.IsNullOrWhiteSpace(unique)) && (!(Funct.ParseFlag(unique, "YES") || Funct.ParseFlag(unique, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "UNIQUE non conforme.";
                        }
                        //Check Chiave Logica
                        if (!(string.IsNullOrWhiteSpace(chiaveLogica)) && (!(Funct.ParseFlag(chiaveLogica, "YES") || Funct.ParseFlag(chiaveLogica, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "CHIAVE LOGICA non conforme.";
                        }
                        //Check Mandatory Flag
                        if (!(string.IsNullOrWhiteSpace(mandatoryFlag)) && (!(Funct.ParseFlag(mandatoryFlag, "YES") || Funct.ParseFlag(mandatoryFlag, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "MANDATORY FLAG non conforme.";
                        }
                        //Check Dominio
                        if (!(string.IsNullOrWhiteSpace(dominio)) && (!(Funct.ParseFlag(dominio, "YES") || Funct.ParseFlag(dominio, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DOMINIO non conforme.";
                        }
                        if (!(string.IsNullOrWhiteSpace(datoSensibile)) && (!(Funct.ParseFlag(datoSensibile, "YES") || Funct.ParseFlag(datoSensibile, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DATO SENSIBILE non conforme.";
                        }

                        if (incorrect == false)
                        {
                            EmptyRow = 0;
                            AttributeT ValRiga = new AttributeT(row: RowPos, db: db, nomeTabellaLegacy: nomeTabella);
                            // Assegnazione valori checkati
                            ValRiga.NomeTabellaLegacy = nomeTabella;
                            ValRiga.NomeCampoLegacy = nomeCampo;
                            ValRiga.DataType = dataType;

                            if (Funct.ParseFlag(chiave, "YES"))
                                ValRiga.Chiave = 0;
                            else
                                ValRiga.Chiave = 100;

                            if (Funct.ParseFlag(unique, "YES"))
                                ValRiga.Unique = unique;
                            else
                                ValRiga.Unique = "N";

                            if (Funct.ParseFlag(chiaveLogica, "YES"))
                                ValRiga.ChiaveLogica = chiaveLogica;
                            else
                                ValRiga.ChiaveLogica = "N";

                            if (Funct.ParseFlag(mandatoryFlag, "YES"))
                                ValRiga.MandatoryFlag = 1;
                            else
                                ValRiga.MandatoryFlag = 0;

                            if (Funct.ParseFlag(dominio, "YES"))
                                ValRiga.Dominio = dominio;
                            else
                                ValRiga.Dominio = "N";

                            ValRiga.Storica = storica;

                            if (Funct.ParseFlag(datoSensibile, "YES"))
                                ValRiga.DatoSensibile = datoSensibile;
                            else
                                ValRiga.DatoSensibile = "N";
                            
                            //Assegnazione valori opzionali
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["SSA"]].Text))
                                ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["SSA"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Area"]].Text))
                                ValRiga.Area = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Area"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text))
                                ValRiga.DefinizioneCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati"]].Text))
                                ValRiga.TipologiaTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati"]].Text;
                            int t;  //Funzionale all'assegnazione di 'Lunghezza' e 'Decimali'
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text))
                                if (int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text, out t))
                                    ValRiga.Lunghezza = t;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text))
                                if(int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text, out t))
                                    ValRiga.Decimali = t;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Provenienza dominio "]].Text))
                                ValRiga.ProvenienzaDominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Provenienza dominio "]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Note"]].Text))
                                ValRiga.Note = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Note"]].Text;
                            listaFile.Add(ValRiga);
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                            AttributeT ValRiga = new AttributeT(row: RowPos, db: db, nomeTabellaLegacy: nomeTabella);
                            // Assegnazione valori checkati
                            ValRiga.NomeTabellaLegacy = nomeTabella;
                            ValRiga.NomeCampoLegacy = nomeCampo;
                            ValRiga.DataType = dataType;
                            ValRiga.History = error;
                            ValRiga.Step = 0;
                            listaFile.Add(ValRiga);
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Field '" + nomeCampo  + "' of Table '" + nomeTabella + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Field '" + nomeCampo + "' of Table '" + nomeTabella + "'. Validation OK", 3, ConfigFile.INFO);
                        }
                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(fileDaAprire);
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }

        internal static bool XLSXWriteErrorInCell(FileInfo fInfo, List<EntityT> list, int col, int v, string tABELLE)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, tABELLE);
            }
            return true;
        }

        internal static bool XLSXWriteErrorInCell(FileInfo fInfo, List<RelationT> list, int col, int v, string rELAZIONI)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, rELAZIONI);
            }
            return true;
        }

        internal static bool XLSXWriteErrorInCell(FileInfo fInfo, List<AttributeT> list, int col, int v, string aTTRIBUTI)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, aTTRIBUTI);
            }
            return true;
        }



        public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<GenericTypeT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        {
            try
            {
                string file = fileDaAprire.FullName;
                if (!File.Exists(file))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
                    return false;
                }
                FileOps.RemoveAttributes(file);
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                        return false;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
                    return false;
                }

                foreach (var worksheet in ws)
                {
                    if (worksheet.Name == sheet)
                    {
                        try
                        {
                            foreach (var dati in Rows)
                            {
                                worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                                worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
                                worksheet.Cells[dati.Row, column].Value = "KO";
                                string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
                                if (mystring == null)
                                    mystring = "";
                                if (!(mystring.Contains(dati.History)))
                                {
                                    worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
                                }
                                worksheet.Column(column + 1).Width = 100;
                                worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                worksheet.Column(column + 1).Style.WrapText = true;
                                Logger.PrintLC("Updating excel file for error " + dati.History, 3);
                            }
                            p.SaveAs(fileDaAprire);
                            return true;
                        }
                        catch (Exception exp)
                        {
                            Logger.PrintLC("Error while writing on file " +
                                            fileDaAprire.Name +
                                            ". Description: " +
                                            exp.Message, 1, ConfigFile.ERROR);
                            return false;
                        }
                    }
                }
                Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
                return false;
            }
            catch
            {
                return false;
            }
        }

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<RelationT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (mystring == null)
        //                            mystring = "";
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message, 1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<AttributeT> Rows,int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (mystring == null)
        //                            mystring = "";
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message,1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<EntityT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message, 1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        public static bool WriteExcelStatsForEntity(FileInfo fileDaAprire, Dictionary<string, List<String>> CompareResults)
        {
            try
            {
                string file = fileDaAprire.FullName;
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage();
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets; 
                    ws.Add(ConfigFile.TABELLE_DIFF);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Errore durante la scrittura di: " + fileDaAprire.Name + ": impossibile aprire il file " + fileDaAprire.DirectoryName, 1, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.TABELLE_DIFF];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Column(1).Width = 50;
                worksheet.Column(2).Width = 50;
                worksheet.Cells[1, 1].Value = "Tabelle Documento Di Ricognizione Caricate In Erwin";
                worksheet.Cells[1, 2].Value = "Tabelle Documento DDL";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(1).Style.WrapText = true;
                worksheet.Column(2).Style.WrapText = true;
                worksheet.View.FreezePanes(2, 1);
                //ExcelRange firstRow = (ExcelRange)worksheet.Row(1);
                //firstRow.f
                //firstRow.Select();
                //firstRow.Application.ActiveWindow.FreezePanes = true;

                int row = 2;
                bool pair = true;
                bool ExcelVuoto = true;
                foreach (var result in CompareResults)
                {
                    foreach (var element in result.Value)
                    {
                        worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        if ((result.Key == "CollezioneTrovati") && ConfigFile.DDL_Show_Right_Rows)
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = element;
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneNonTrovatiSQL")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = "KO: Entity non trovata sul DDL";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneNonTrovatiXLS")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 2].Value = element;
                            worksheet.Cells[row, 1].Value = "KO: Entity non caricata su Erwin";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        
                    }
                    
                }
                if (ExcelVuoto)
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Differenza Riscontrata";
                    worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);
                }

                Logger.PrintLC("Fine compilazione file excel", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file. Errore: " + exp.Message , 4, ConfigFile.ERROR);
                return false;
            }
        }

        public static bool WriteExcelStatsForAttribute(FileInfo fileDaAprire, Dictionary<string, List<String>> CompareResults, List<AttributeT> Attributi)
        {
            try
            {
                string file = fileDaAprire.FullName;

                if (!File.Exists(file))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", 1, ConfigFile.ERROR);
                    return false;
                }
                FileOps.RemoveAttributes(file);
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                        return false;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
                    ws.Add(ConfigFile.ATTRIBUTI_DIFF);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Lettura file: " + fileDaAprire.Name + ": impossibile aprire il percorso " + fileDaAprire.DirectoryName, 1, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.ATTRIBUTI_DIFF];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                worksheet.Column(1).Width = 45;
                worksheet.Column(2).Width = 45;
                worksheet.Column(3).Width = 25;
                worksheet.Column(4).Width = 25;
                worksheet.Column(5).Width = 25;
                worksheet.Cells[1, 1].Value = "Attributi Documento Di Ricognizione Caricati In Erwin";
                worksheet.Cells[1, 2].Value = "Attributi Documento DDL";
                worksheet.Cells[1, 3].Value = "Differenze Campo Datatype";
                worksheet.Cells[1, 4].Value = "Differenze Campo Chiave";
                worksheet.Cells[1, 5].Value = "Differenze Campo Mandatory";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                //worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 3].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 4].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 5].Style.Font.Color.SetColor(Color.Red);
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(1).Style.WrapText = true;
                worksheet.Column(2).Style.WrapText = true;
                worksheet.Column(3).Style.WrapText = true;
                worksheet.Column(4).Style.WrapText = true;
                worksheet.Column(5).Style.WrapText = true;
                //Excel.Range firstRow = (Excel.Range)worksheet.Row(1);
                //firstRow.Activate();
                //firstRow.Select();
                //firstRow.Application.ActiveWindow.FreezePanes = true;
                worksheet.View.FreezePanes(2, 1);

                bool ExcelVuoto = true;

                int row = 2;
                bool pair = true;
                foreach (var result in CompareResults)
                {
                    foreach (var element in result.Value)
                    {
                        worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        if (result.Key == "CollezioneAttributiTrovati")
                        {
                            string[] elementi = element.Split('|');
                            if (elementi.Count() != 4)
                            {
                                worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                                ExcelVuoto = false;
                                continue;
                            }
                            
                            
                            AttributeT AttributoRif = Attributi.Find(x => elementi[0] == x.NomeTabellaLegacy + "." + x.NomeCampoLegacy);
                            bool datatypeOK = true;
                            bool mandatoryOK = true;
                            bool keyOK = true;
                            string mandatoryXLS = string.Empty;
                            string mandatoryDDL = string.Empty;
                            string keyXLS = string.Empty;
                            string keyDDL = string.Empty;

                            mandatoryDDL = elementi[2] == "true" ? "NOT NULL" : "NULL";
                            mandatoryXLS = AttributoRif.MandatoryFlag == 1 ? "NOT NULL" : "NULL";
                            keyXLS = elementi[3] == "true" ? "CHIAVE PRIMARIA" : "";
                            keyDDL = AttributoRif.Chiave == 0 ? "CHIAVE PRIMARIA" : "";

                            if (AttributoRif.DataType != elementi[1])
                                datatypeOK = false;
                            if (mandatoryDDL != mandatoryXLS)
                                mandatoryOK = false;
                            if (keyDDL != keyXLS)
                                keyOK = false;

                            if ((!ConfigFile.DDL_Show_Right_Rows) && datatypeOK && mandatoryOK && keyOK) 
                            {
                              // se tutte e 4 le condizioni sono vere non scrive. Se anche solo una è falsa scrive.  
                            }
                            else
                            { 
                                ExcelVuoto = false;
                                worksheet.Cells[row, 1].Value = elementi[0];
                                worksheet.Cells[row, 2].Value = elementi[0];
                                worksheet.Cells[row, 3].Value = "XLS: " + AttributoRif.DataType + "\n" + "DDL: " + elementi[1];
                                worksheet.Cells[row, 4].Value = "XLS: " + keyXLS + "\n" + "DDL: " + keyDDL;
                                worksheet.Cells[row, 5].Value = "XLS: " + mandatoryXLS + "\n" + "DDL: " + mandatoryDDL;


                                if (pair)
                                {
                                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    if (datatypeOK)
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                    if (mandatoryOK)
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                    if (keyOK)
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);

                                }
                                else
                                {
                                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    if (datatypeOK)
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    if (mandatoryOK)
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    if (keyOK)
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                }
                                row += 1;
                                pair = !pair;
                            }

                        }
                        if (result.Key == "CollezioneAttributiNonTrovatiSQL")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = "KO: Attributo non trovato sul DDL";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneAttributiNonTrovatiXLS")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 2].Value = element;
                            worksheet.Cells[row, 1].Value = "KO: Attributo non caricato su Erwin";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
                        
                    }

                }

                if (ExcelVuoto)
                {
                    worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Differenza Riscontrata";
                    worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 3].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 4].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 4].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 5].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 5].Style.Font.Color.SetColor(Color.White);

                }

                Logger.PrintLC("Fine compilazione file excel", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file. Errore: " + exp.Message, 4, ConfigFile.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Write on specific 'ControlliTempistiche' file a
        /// list of value.
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="ListCodLocaleControllo"></param>
        /// <returns></returns>
        public static bool WriteDocExcelControlliTempistiche(FileInfo fileDaAprire, List<string> ListCodLocaleControllo)
        {
            string TemplateFile = null;
            if (!string.IsNullOrWhiteSpace(ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE))
            {
                TemplateFile = ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE;
            }
            else
            {
                Logger.PrintLC("Value of 'ControlliTempistiche Template' is not valid. Will not valorize its content.", 2, ConfigFile.ERROR);
                return false;
            }
            string file = fileDaAprire.FullName;
            try
            {
                if (!File.Exists(TemplateFile))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", 1, ConfigFile.ERROR);
                    return false;
                }
                else
                {
                    File.Copy(TemplateFile, file);
                }
                FileOps.RemoveAttributes(file);
            }
            catch
            {
            }

            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets;
                //ws.Add(ConfigFile.CONTROLLI);
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Error while opening: " + fileDaAprire.FullName + ". Will not valorize its content.", 2, ConfigFile.ERROR);
                return false;
            }



            /*
            ExcelWorksheet worksheet = null;
            try
            {
                worksheet = ws[ConfigFile.CONTROLLI_TEMPISTICHE];
            }
            catch
            {
                Logger.PrintLC("Could not find sheet \"" + ConfigFile.CONTROLLI_TEMPISTICHE + "\" in " + file +
                    ". Will not valorize its content.", 4, ConfigFile.INFO);
            }
            Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);
            */


            //bool ExcelVuoto = true;
            foreach (var worksheet in ws)
            {
                if(worksheet.Name == ConfigFile.CONTROLLI_TEMPISTICHE) { 
                int row = 2;
                bool pair = true;
                    foreach (var element in ListCodLocaleControllo)
                    {
                        //worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        string[] elementi = element.Split('|');
                        if (elementi.Count() == 4)
                        {
                            worksheet.Cells[row, 12].Value = "errore nella comparazione dell'elemento: " + element;
                            //ExcelVuoto = false;
                            continue;
                        }
                        //ExcelVuoto = false;
                        worksheet.Cells[row, 12].Value = element.ToString();

                        //if (pair)
                        //{
                        //    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                        //}
                        //else
                        //{
                        //    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                        //}
                        Logger.PrintLC("Riga " + row + " scritta nel file excel", 6, ConfigFile.INFO);
                        row += 1;
                        pair = !pair;
                    }
                }
            }
            fileDaAprire.IsReadOnly = false;

            try
            {
                p.Save();
                //p.SaveAs(fileDaAprire);

                if (ws != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(ws);
                    }
                    catch { }
                }
                if (WB != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(WB);
                    }
                    catch { }
                }
                if (ExApp != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(ExApp);
                    }
                    catch { }
                }
                if (p != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(p);
                    }
                    catch { }
                }
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Error while closing Excel application. Will continue without notice.", 5, ConfigFile.ERROR);
            }
            return true;
        }







        public static bool WriteDocExcelControlliCampiX(FileInfo fileDaAprire, List<string> ExcelControlli)
        {
            string TemplateFile = ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE;

            string file = fileDaAprire.FullName;
            if (!File.Exists(TemplateFile))
            {
                Logger.PrintLC("Trying to find File " + fileDaAprire.Name + ": doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
            else
            {
                File.Copy(TemplateFile, file);
            }
            fileDaAprire = new FileInfo(file);
            //FileOps.RemoveAttributes(file);

            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = new ExcelPackage(fileDaAprire);
            ExcelWorkbook WB = p.Workbook;
            ExcelWorksheets ws = WB.Worksheets;

            //ExcelWorkbook WB = null;
            //ExcelWorksheets ws = null;
            //try
            //{
            //ExApp.DisplayAlerts = false;
            //p = new ExcelPackage(fileDaAprire);
            //ExApp.DisplayAlerts = true;
            //ws.Add(ConfigFile.CONTROLLI);
            //}

            var worksheet = ws[ConfigFile.CONTROLLI_TEMPISTICHE];
            bool isProtect = worksheet.Protection.IsProtected;
            Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

            worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Column(4).Width = 45;
            worksheet.Column(5).Width = 45;
            worksheet.Column(6).Width = 45;
            worksheet.Column(7).Width = 45;
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(4).Style.WrapText = true;
            worksheet.Column(5).Style.WrapText = true;
            worksheet.Column(6).Style.WrapText = true;
            worksheet.Column(7).Style.WrapText = true;
            worksheet.View.FreezePanes(2, 1);

            bool ExcelVuoto = true;

            int row = 2;
            bool pair = true;
            foreach (var element in ExcelControlli)
            {
                worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                string[] elementi = element.Split('|');
                if (elementi.Count() == 2)
                {
                    worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                    ExcelVuoto = false;
                    continue;
                }
                ExcelVuoto = false;
                worksheet.Cells[row, 12].Value = elementi[0];

                if (pair)
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                }
                else
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                }
                row += 1;
                pair = !pair;

                Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
            }

            if (ExcelVuoto)
            {
                worksheet.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Cells[2, 1].Value = "Nessuna Controllo Riscontrato";

            }

            Logger.PrintLC("Fine compilazione file excel controlli", 4, ConfigFile.INFO);
            ExcelProtectedRangeCollection range = worksheet.ProtectedRanges;
            p.Save();
            p.Dispose();
            Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
            return true;
        }







        public static bool WriteDocExcelControlliCampi(FileInfo fileDaAprire, List<string> ExcelControlli)
        {
            string TemplateFile = ConfigFile.CONTROLLI_CAMPI_TEMPLATE;
            
            try
            {
                string file = fileDaAprire.FullName;
                if (!File.Exists(TemplateFile))
                {
                    Logger.PrintLC("Trying to find File " + fileDaAprire.Name + ": doesn't exist.", 2, ConfigFile.ERROR);
                    return false;
                }
                else
                {
                    File.Copy(TemplateFile, file);
                }
                FileOps.RemoveAttributes(file);

                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets;
                    //ws.Add(ConfigFile.CONTROLLI);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Errore durante la scrittura di: " + fileDaAprire.Name + ": impossibile aprire il file in " + fileDaAprire.DirectoryName, 2, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.CONTROLLI_CAMPI];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                worksheet.Column(4).Width = 45;
                worksheet.Column(5).Width = 45;
                worksheet.Column(6).Width = 45;
                worksheet.Column(7).Width = 45;
                //worksheet.Cells[1, 1].Value = "Nome Struttura Informativa";
                //worksheet.Cells[1, 2].Value = "Nome Campo";
                //worksheet.Cells[1, 3].Value = "Cod Locale Controllo";
                //worksheet.Cells[1, 4].Value = "Ruolo Campo";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(4).Style.WrapText = true;
                worksheet.Column(5).Style.WrapText = true;
                worksheet.Column(6).Style.WrapText = true;
                worksheet.Column(7).Style.WrapText = true;
                worksheet.View.FreezePanes(2, 1);

                bool ExcelVuoto = true;

                int row = 2;
                bool pair = true;
                foreach (var element in ExcelControlli)
                {
                    worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                    string[] elementi = element.Split('|');
                    if (elementi.Count() != 4)
                    {
                        worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                        ExcelVuoto = false;
                        continue;
                    }
                    ExcelVuoto = false;
                    worksheet.Cells[row, 4].Value = elementi[0];
                    worksheet.Cells[row, 5].Value = elementi[1];
                    worksheet.Cells[row, 6].Value = elementi[2];
                    worksheet.Cells[row, 7].Value = elementi[3];
 
                    if (pair)
                    {
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                    else
                    {
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                    }
                    row += 1;
                    pair = !pair;

                    Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
                }

                if (ExcelVuoto)
                {
                    worksheet.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Controllo Riscontrato";
                    
                }

                Logger.PrintLC("Fine compilazione file excel controlli", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file excel ControlliCampi. Errore: " + exp.Message, 4, ConfigFile.ERROR);
                return false;
            }
        }

    }
}
