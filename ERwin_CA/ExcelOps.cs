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

            //if (ExApp == null)
            //    return false;

            ExApp = new Excel.ApplicationClass();
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
                    Logger.PrintLC("Successfully converted " + fileInfo.FullName + " to " + fileName, 2, ConfigFile.INFO);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Error: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
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
                    ExApp = new Excel.ApplicationClass();

                    //FileOps.RemoveAttributes(fileName);
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
            //SCAPI.Application testAPP = new SCAPI.Application();
            string testoLog = string.Empty;
            FileInfo fileDaAprire = new FileInfo(file);
            bool isXLS = false;
            if (fileDaAprire.Extension == ".xls")
            {
                if (!ConvertXLStoXLSX(file))
                    if (!ConvertXLStoXLSX(file))
                        return false;
                isXLS = true;
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
            int[] check_sheet = new int[3] { 0, 0, 0 };
            foreach (var worksheet in ws)
            {
                // SEZIONE TABELLE
                if (worksheet.Name == ConfigFile.TABELLE)
                {
                    columns = 0;
                    check_sheet[0] += 1;
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
                            Logger.PrintLC(testoLog, 2, ConfigFile.ERROR);
                            goto ERROR;
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_TABELLE)
                        columnsFound = true;
                    else
                        goto ERROR;
                }

                // SEZIONE ATTRIBUTI
                if (worksheet.Name == ConfigFile.ATTRIBUTI)
                {
                    check_sheet[1] += 1;
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
                            Logger.PrintLC(testoLog, 2, ConfigFile.ERROR);
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
                    check_sheet[2] += 1;
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
                            Logger.PrintLC(testoLog, 2, ConfigFile.ERROR);
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
            if(check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1)
            {
                Logger.PrintLC(fileDaAprire.Name + ": file could not be processed: at least one Sheet is missing from the file.", 2, ConfigFile.ERROR);
                Logger.PrintF(fileError, "er_driveup – Caricamento Excel su ERwin", true);
                Logger.PrintF(fileError, "Foglio/i mancante/i.", true);
                if(isXLS == true)
                    if (File.Exists(fileDaAprire.FullName))
                        File.Delete(fileDaAprire.FullName);
                return false;
            }
            if (sheetFound != true || columnsFound != true)
            {
                Logger.PrintLC(fileDaAprire.Name + ": file could not be processed: Columns or Sheets are not in the expected format.", 2, ConfigFile.ERROR);
                Logger.PrintF(fileError, "er_driveup – Caricamento Excel su ERwin", true);
                Logger.PrintF(fileError, "Colonne o Fogli non formattati correttamente.", true);
                if (isXLS == true)
                    if (File.Exists(fileDaAprire.FullName))
                        File.Delete(fileDaAprire.FullName);
                return false;
            }
            Logger.PrintLC(fileDaAprire.Name + ": file valid to be processed.", 2, ConfigFile.INFO);
            Logger.PrintF(fileCorrect, "er_driveup – Caricamento Excel su ERwin", true);
            Logger.PrintF(fileCorrect, "Colonne e Fogli formattati corretamente.", true);
            return true;
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
                Logger.PrintLC("Reading Tables. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3, ConfigFile.ERROR);
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
                        bool incorrect = false;
                        string nome = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text;
                        string flag = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                        if (string.IsNullOrWhiteSpace(nome))
                        {
                            incorrect = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "Valore di NOME TABELLA mancante.";
                        }
                        if (!(string.IsNullOrWhiteSpace(flag)) && (!(string.Equals(flag, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(flag, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            string error = worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + 2].Text;
                            if (!string.IsNullOrWhiteSpace(error))
                                error = error + " ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error + "Valore di FLAG BFD non conforme.";
                        }
                        
                        if (incorrect == false)
                        { 
                            EmptyRow = 0;
                            EntityT ValRiga = new EntityT(row: RowPos, db: db, tName: nome);
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
                            else
                                ValRiga.FlagBFD = "N";
                            listaFile.Add(ValRiga);
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                        }
                        else
                        {
                            //EmptyRow += 1;
                            //if (EmptyRow >= 10)
                            //{
                            //    FilesEnd = true;
                            //}
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
                    }
                    p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
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
                        string identificativoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativo relazione"]].Text;
                        string tabellaPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Padre"]].Text;
                        string tabellaFiglia = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Figlia"]].Text;
                        string cardinalita = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Cardinalità"]].Text;
                        string campoPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Padre"]].Text;
                        string campoFiglio = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Figlio"]].Text;
                        string identificativa = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativa"]].Text;
                        string eccezione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text;
                        string tipoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tipo Relazione"]].Text;
                        string note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text;


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
                            if (!(cardinalita.Equals("1:1") || cardinalita.Equals("1:n")))
                            {
                                incorrect = true;
                                error += "CARDINALITA non conforme. ";
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
                            if (!(identificativa.ToUpper().Equals("S") || identificativa.ToUpper().Equals("N")))
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
                            if (cardinalita.Equals("1:1"))
                                ValRiga.Cardinalita = -1;
                            else
                                ValRiga.Cardinalita = -3;
                            ValRiga.CampoPadre = campoPadre;
                            ValRiga.CampoFiglio = campoFiglio;
                            if (identificativa.ToUpper().Equals("S"))
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
                            //EmptyRow += 1;
                            //if (EmptyRow >= 10)
                            //    FilesEnd = true;
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
                        string nomeTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text;
                        string nomeCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text;
                        string dataType = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Datatype"]].Text;
                        dataType = Funct.RemoveWhitespace(dataType);
                        string chiave = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave"]].Text;
                        string unique = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Unique"]].Text;
                        string chiaveLogica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave Logica"]].Text;
                        string mandatoryFlag = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Mandatory Flag"]].Text;
                        string dominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dominio"]].Text;
                        string storica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Storica"]].Text;
                        string datoSensibile = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dato Sensibile"]].Text;

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
                        //if (!(string.Equals(chiave, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(chiave, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(chiave)) && (!(string.Equals(chiave, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(chiave, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                        incorrect = true;
                        if (!string.IsNullOrWhiteSpace(error))
                            error += " ";
                        error += "CHIAVE non conforme.";
                        }
                        //Check Unique
                        //if (!(string.Equals(unique, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(unique, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(unique)) && (!(string.Equals(unique, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(unique, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "UNIQUE non conforme.";
                        }
                        //Check Chiave Logica
                        //if (!(string.Equals(chiaveLogica, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(chiaveLogica, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(chiaveLogica)) && (!(string.Equals(chiaveLogica, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(chiaveLogica, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "CHIAVE LOGICA non conforme.";
                        }
                        //Check Mandatory Flag
                        //if (!(string.Equals(mandatoryFlag, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(mandatoryFlag, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(mandatoryFlag)) && (!(string.Equals(mandatoryFlag, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(mandatoryFlag, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "MANDATORY FLAG non conforme.";
                        }
                        //Check Dominio
                        //if (!(string.Equals(dominio, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(dominio, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(dominio)) && (!(string.Equals(dominio, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(dominio, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DOMINIO non conforme.";
                        }
                        //Check Storica
                        //if (!(string.Equals(storica, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(storica, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(storica)) && (!(string.Equals(storica, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(storica, "N", StringComparison.OrdinalIgnoreCase))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "STORICA non conforme.";
                        }
                        //Check Dato Sensibile
                        //if (!(string.Equals(datoSensibile, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(datoSensibile, "N", StringComparison.OrdinalIgnoreCase)))
                        if (!(string.IsNullOrWhiteSpace(datoSensibile)) && (!(string.Equals(datoSensibile, "S", StringComparison.OrdinalIgnoreCase) || string.Equals(datoSensibile, "N", StringComparison.OrdinalIgnoreCase))))
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

                            if (string.Equals(chiave.ToUpper(), "S"))
                                ValRiga.Chiave = 0;
                            else
                                ValRiga.Chiave = 100;

                            if (string.Equals(unique.ToUpper(), "S"))
                                ValRiga.Unique = unique;
                            else
                                ValRiga.Unique = "N";

                            if (string.Equals(chiaveLogica.ToUpper(), "S"))
                                ValRiga.ChiaveLogica = chiaveLogica;
                            else
                                ValRiga.ChiaveLogica = "N";
                            
                            if (string.Equals(mandatoryFlag.ToUpper(), "S"))
                                ValRiga.MandatoryFlag = 1;
                            else
                                ValRiga.MandatoryFlag = 0;

                            if (string.Equals(dominio.ToUpper(), "S"))
                                ValRiga.Dominio = dominio;
                            else
                                ValRiga.Dominio = "N";

                            if (string.Equals(storica.ToUpper(), "S"))
                                ValRiga.Storica = storica;
                            else
                                ValRiga.Storica = "N";

                            if (string.Equals(datoSensibile.ToUpper(), "S"))
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

                        }
                        else
                        {
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = error;
                            //EmptyRow += 1;
                            //if (EmptyRow >= 10)
                            //    FilesEnd = true;
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
                    }
                    p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    return listaFile;
                }
            }
            return listaFile = null;
        }

        
        public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, int row, int column, int priorityWrite, string text, string sheet = ConfigFile.ATTRIBUTI)
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
                if (fileDaAprire.Extension == ".xls")
                {
                    if (!ConvertXLStoXLSX(file))
                        return false;
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
                    Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
                    return false;
                }

                //bool FilesEnd = false;
                //int EmptyRow = 0;
                //int columns = 0;
                foreach (var worksheet in ws)
                {
                    if (worksheet.Name == sheet)
                    {
                        try
                        {
                            worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[row, column].Style.Font.Bold = true;
                            worksheet.Cells[row, column].Value = "KO";
                            worksheet.Cells[row, column + 1].Value = text;
                            p.SaveAs(fileDaAprire);
                            return true;
                        }
                        catch (Exception exp)
                        {
                            Logger.PrintLC("Error while writing on file " +
                                            fileDaAprire.Name +
                                            ". Description: " +
                                            exp.Message,1, ConfigFile.ERROR);
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



    }
}
