using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop.Excel;
using log4net;

namespace IMF_WebRequestClient
{
    class ProcessFile
    {
        ILog log = null;
        public ProcessFile(ILog p_log)
        {
            log = p_log;
        }
        public string ProcessAndSaveX(string p_strFileName, DateTime p_Date)
        {
            Application app = new Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;
            try
            {
                SqlConnection conn = Utility.GetDataConn(Utility.RIMSDBConnectionString);
                log.Info(string.Format("Connected to connection string {0}", Utility.RIMSDBConnectionString));
                System.Data.DataTable dt = new System.Data.DataTable();
                dt = Utility.getData(conn, Utility.SQLStatementProcDate);
                System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
                string strMonthName = mfi.GetMonthName(p_Date.Month);
                string strYear = p_Date.Year.ToString();
                string strFilterOne = string.Format("Currency units per SDR for {0} {1}", strMonthName, strYear);
                log.Info(string.Format("When date is 15 or < filter is = {0}.", strFilterOne));
                string strFilterTwo = string.Format("Currency units per SDR for {0} {1} Continued", strMonthName, strYear);
                log.Info(string.Format("Date > = Mid Month filter is = {0}.", strFilterTwo));

                string cur_data_date_FILENAME = p_Date.ToString("yyyyMMdd");

                log.Info(string.Format(" Retrieving  {0}.", cur_data_date_FILENAME));
                //string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
                string cur_data_date_COLUMN = p_Date.ToString("d-MMM-yy");

                string strFileNameX = string.Format(Utility.FilePathSaveTo, cur_data_date_FILENAME);
                string strSheetName = string.Format(Utility.SheetName, cur_data_date_FILENAME);
                string strSheetNameIDX = string.Format(Utility.SheetNameIDX, cur_data_date_FILENAME);

                Workbook book = null;
                book = app.Workbooks.Open(p_strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                log.Info(string.Format("{0} opened.", p_strFileName));
                Worksheet EditSheet = book.Worksheets[strSheetNameIDX];
                EditSheet.Rows["1:1"].Delete();
                log.Info(string.Format("Delete header row File = {0}", p_strFileName));


                Range FindRange = (Range)EditSheet.Columns["A", Type.Missing];
                //Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                if (!(FoundRangeContinued is null))
                {

                    int rowNumberContinued = FoundRangeContinued.Row + 1;
                    string rangeContinued = string.Format("1:{0}", rowNumberContinued.ToString());
                    EditSheet.Rows[rangeContinued].Delete();
                    log.Info(string.Format("Delete Continue = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));

                }

                Range FoundRangeNotes = FindRange.Find("Notes:");
                if (!(FoundRangeNotes is null))
                {
                    int RowNumberNotes = FoundRangeNotes.Row - 1;
                    string rangeNotes = string.Format("{0}:{1}", RowNumberNotes.ToString(), (RowNumberNotes + 50).ToString());
                    EditSheet.Rows[rangeNotes].Delete();
                    log.Info(string.Format("Delete Notes = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));
                }




                //EditSheet.Range["1","1:1"].Find()
                //book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //book.SaveAs(p_strFileName, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Save();
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                EditSheet = null;
                book = null;


                //log.Info(string.Format("Openning OleDBConnection to {0}.", p_strFileName));
                log.Info(string.Format("Openning OleDBConnection to {0}.", strFileNameX));

                //OleDbConnection OleDBconn = new OleDbConnection(string.Format(Utility.XLSConnectionString, p_strFileName));
                OleDbConnection OleDBconn = new OleDbConnection(string.Format(Utility.XLSXConnectionString, strFileNameX));
                OleDBconn.Open();

                //Data table to be filled by excel spreadsheet
                System.Data.DataTable dtXLS = new System.Data.DataTable();
                
                //OleCommand.Parameters.Add(new OleDbParameter("Currency", "Chinese yuan,Euro,Japanese yen,U.K. pound,U.K. Pound Sterling,U.S. dollar,Mexican peso,Canadian dollar,Australian dollar,Swedish krona,Norwegian krone"));
               
                System.Data.DataTable ExcelSheets = OleDBconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string EXTableName = ExcelSheets.Rows[0]["TABLE_NAME"].ToString();

                log.Info(string.Format("Worksheet to fill data table with : {0}", EXTableName));
                
                OleDbCommand OleCommand = new OleDbCommand(string.Format(Utility.ExcelSQLQuery, EXTableName, cur_data_date_COLUMN));
                OleCommand.Connection = OleDBconn;
                //log.Info(string.Format("OleDBConnection assigned to OleCommand :{0}", p_strFileName));
                log.Info("QueryString for EXCEL - " + string.Format(Utility.ExcelSQLQuery, EXTableName, cur_data_date_COLUMN));
                log.Info(string.Format("OleDBConnection String assigned to OleCommand :{0}", Utility.XLSXConnectionString));

                OleDbDataAdapter adp = new OleDbDataAdapter(OleCommand);

                //log.Info(string.Format("Fill Data Table with data from File :{0}", p_strFileName));
                log.Info(string.Format("Fill Data Table with data from File :{0}", strFileNameX));
                adp.Fill(dtXLS);

                //log.Info(string.Format("Data Table filled with data from File :{0}", p_strFileName));
                log.Info(string.Format("Data Table filled with data from File :{0}", strFileNameX));
                OleDBconn.Close();

                //log.Info(string.Format("Open file in Excel :{0}", p_strFileName));
                log.Info(string.Format("Open file in Excel :{0}", strFileNameX));

                //book = app.Workbooks.Open(p_strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book = app.Workbooks.Open(strFileNameX, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //log.Info(string.Format("File openned in Excel :{0}", p_strFileName));
                log.Info(string.Format("File openned in Excel :{0}", strFileNameX));


                book.Worksheets.Add();
                book.Worksheets["Sheet1"].Name = "DataToLoadFX";
                log.Info(string.Format("Loading data to {0}.", "DataToLoadFX"));
                Worksheet DataToLoadFX = book.Worksheets["DataToLoadFX"];
                DataToLoadFX.Cells[1, 1] = "Code";
                DataToLoadFX.Cells[1, 2] = "Category";
                DataToLoadFX.Cells[1, 3] = "CodeType";
                DataToLoadFX.Cells[1, 4] = "Date";
                DataToLoadFX.Cells[1, 5] = "Value";
                DataToLoadFX.Cells[1, 6] = "Type";
                DataToLoadFX.Cells[1, 7] = "Source";
                DataToLoadFX.Cells[1, 8] = "Email";
                DataToLoadFX.Cells[1, 9] = "Overwrite";
                string strdtLastUpdate = p_Date.ToString("MM/dd/yyyy");
                //DateTime dtLastUpdated = DateTime.ParseExact(strdtLastUpdate, "dd-MM-yyyy hh:mm",)
                //string filterString = string.Format("[{0}]", cur_data_date_COLUMN);
                //string filterString = string.Format("[{0}] = 'NA'", cur_data_date_COLUMN);
                int i = 2;
                int offset = dtXLS.Rows.Count;
                DataRow[] drUS = dtXLS.Select("Currency ='U.S. dollar'");
                double USD_val = System.Convert.ToDouble(drUS[0][2]);
                foreach (DataRow dr in dtXLS.Rows)
                {
                    DataToLoadFX.Cells[i, 1] = dr[1];
                    DataToLoadFX.Cells[i + offset, 1] = dr[1];
                    DataToLoadFX.Cells[i, 2] = "Currency";
                    DataToLoadFX.Cells[i + offset, 2] = "Currency";
                    DataToLoadFX.Cells[i, 3] = "Rogge.Code.Default";
                    DataToLoadFX.Cells[i + offset, 3] = "Rogge.Code.Default";
                    DataToLoadFX.Cells[i, 4] = strdtLastUpdate;
                    DataToLoadFX.Cells[i + offset, 4] = strdtLastUpdate;
                    string testValue = dr[1].ToString().Trim();
                    string tempCurrencyVal = dr[2].ToString().Trim();
                    if (tempCurrencyVal.Trim() != "NA" && tempCurrencyVal.Trim() != "")
                    {
                        Double CurrencyVal = Convert.ToDouble(tempCurrencyVal);
                        if (testValue.Equals("EUR") || testValue.Equals("GBP") || testValue.Equals("AUD"))
                            DataToLoadFX.Cells[i, 5] = USD_val / CurrencyVal;
                        else if (testValue.Equals("XDR"))
                            DataToLoadFX.Cells[i, 5] = USD_val;
                        else
                            DataToLoadFX.Cells[i, 5] = CurrencyVal / USD_val;

                    }
                    DataToLoadFX.Cells[i + offset, 5] = 0;
                    DataToLoadFX.Cells[i, 6] = "MID";
                    DataToLoadFX.Cells[i + offset, 6] = "BASK";
                    DataToLoadFX.Cells[i, 7] = "IMF";
                    DataToLoadFX.Cells[i + offset, 7] = "IMF";
                    DataToLoadFX.Cells[i, 8] = "ITDataAdmin@RoggeGlobal.com";
                    DataToLoadFX.Cells[i + offset, 8] = "ITDataAdmin@RoggeGlobal.com";
                    DataToLoadFX.Cells[i, 9] = "N";
                    DataToLoadFX.Cells[i + offset, 9] = "N";




                    i++;
                }
                i = i + offset;
                //book.Worksheets.Add();
                //book.Worksheets["Sheet2"].Name = "rms_sdrv";


                //wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                ////wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                //wsRMS_SDRV.QueryTables[1].Name = "rms_sdrv";
                //wsRMS_SDRV.QueryTables[1].FieldNames = true;
                //wsRMS_SDRV.QueryTables[1].RowNumbers = false;
                //wsRMS_SDRV.QueryTables[1].FillAdjacentFormulas = false;
                //wsRMS_SDRV.QueryTables[1].WebFormatting = XlWebFormatting.xlWebFormattingNone;
                ////wsRMS_SDRV.QueryTables[1].PreserveFormatting = false;
                //wsRMS_SDRV.QueryTables[1].SaveData = true;
                //wsRMS_SDRV.QueryTables[1].Refresh(false);
                log.Info(string.Format("Loading data to {0}.", "rms_sdrv"));
                Worksheet wsRMS_SDRV = book.Worksheets["rms_sdrv"];
                Range ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                Range rgCurrencyUnitFind = ARange.Find("Currency Unit");

                int CurrencyRow = rgCurrencyUnitFind.Row;
                string deleteCurrencyRows = string.Format("1:{0}", (CurrencyRow - 1).ToString());
                wsRMS_SDRV.Rows[deleteCurrencyRows].Delete();


                ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                Range rgNotes = (Range)ARange.Find("Footnotes");
                if (!(rgNotes is null))
                {
                    int NotesRow = rgNotes.Row;
                    string deleteNotesRows = string.Format("{0}:{1}", (NotesRow - 1).ToString(), (NotesRow + 30).ToString());
                    wsRMS_SDRV.Rows[deleteNotesRows].Delete();
                }


                //switch()
                //{

                //}

                for (int j = 2; j < 10; j++)
                {
                    double result;
                    var TempValParse = (wsRMS_SDRV.Cells[j, 1] as Range).Value;
                    string strTempValParse = TempValParse.ToString();

                    var tempVal = Double.TryParse(strTempValParse, out result) ? "SDR1=US$" : strTempValParse;
                    if (tempVal != "U.S.$1.00 = SDR" && tempVal != "U.S. dollar" && !Double.TryParse(strTempValParse, out result))
                    {
                        switch (tempVal)
                        {
                            case "Chinese yuan":
                                DataToLoadFX.Cells[i, 1] = "XDRCNY";
                                break;
                            case "Euro":
                                DataToLoadFX.Cells[i, 1] = "XDREUR";
                                break;
                            case "Japanese yen":
                                DataToLoadFX.Cells[i, 1] = "XDRJPY";
                                break;
                            case "U.K. pound":
                                DataToLoadFX.Cells[i, 1] = "XDRGBP";
                                break;
                            case "SDR1 = US$":
                                DataToLoadFX.Cells[i, 1] = "XDRUSD";
                                break;

                            default:
                                break;
                        }
                        DataToLoadFX.Cells[i, 2] = "Index";
                        DataToLoadFX.Cells[i, 3] = "Rogge.Code.Default";
                        DataToLoadFX.Cells[i, 4] = strdtLastUpdate;
                        if (tempVal.Equals("SDR1 = US$"))
                        {
                            string tempXDRUSA = (string)(wsRMS_SDRV.Cells[j, 4] as Range).Value;
                            int lentempXDRUSA = tempXDRUSA.Length - 2;
                            tempXDRUSA = tempXDRUSA.Substring(0, lentempXDRUSA);

                            double tempXDRUSADBL = Convert.ToDouble(tempXDRUSA);
                            DataToLoadFX.Cells[i, 5] = tempXDRUSADBL;

                        }
                        else
                            DataToLoadFX.Cells[i, 5] = wsRMS_SDRV.Cells[j, 3];
                        DataToLoadFX.Cells[i, 6] = "MID";
                        DataToLoadFX.Cells[i, 7] = "IMF";
                        DataToLoadFX.Cells[i, 8] = "ITDataAdmin@RoggeGlobal.com";
                        DataToLoadFX.Cells[i, 9] = "N";
                        i++;
                    }

                }

                //book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Save();


                app.Quit();
                app = null;

                dtXLS = null;

                return strFileNameX;

            }
            catch (Exception e)
            {
                app.Workbooks.Close();
                app.Quit();
                log.Info(e.Message);
                throw new Exception(e.Message, e.InnerException);
            }
        }
        public string ProcessAndSave(string p_strFileName, DateTime p_Date)
        {
            Application app = new Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;
            try
            {
                SqlConnection conn = Utility.GetDataConn(Utility.RIMSDBConnectionString);
                log.Info(string.Format("Connected to connection string {0}", Utility.RIMSDBConnectionString));
                System.Data.DataTable dt = new System.Data.DataTable();
                dt = Utility.getData(conn, Utility.SQLStatementProcDate);
                System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
                string strMonthName = mfi.GetMonthName(p_Date.Month);
                string strYear = p_Date.Year.ToString();
                string strFilterOne = string.Format("Currency units per SDR for {0} {1}", strMonthName, strYear);
                log.Info(string.Format("When date is 15 or < filter is = {0}.",strFilterOne));
                string strFilterTwo = string.Format("Currency units per SDR for {0} {1} Continued", strMonthName, strYear);
                log.Info(string.Format("Date > = Mid Month filter is = {0}.", strFilterTwo));

                string cur_data_date_FILENAME = p_Date.ToString("yyyyMMdd");

                log.Info(string.Format(" Retrieving  {0}.", cur_data_date_FILENAME));
                //string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
                string cur_data_date_COLUMN = p_Date.ToString("d-MMM-yy");
                
                string strFileNameX = string.Format(Utility.FilePathSaveTo, cur_data_date_FILENAME);
                string strSheetName = string.Format(Utility.SheetName, cur_data_date_FILENAME);
                string strSheetNameIDX = string.Format(Utility.SheetNameIDX, cur_data_date_FILENAME);

                Workbook book = null;
                book = app.Workbooks.Open(p_strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                log.Info(string.Format("{0} opened.", p_strFileName));
                Worksheet EditSheet = book.Worksheets[strSheetNameIDX];
                EditSheet.Rows["1:1"].Delete();
                log.Info(string.Format("Delete header row File = {0}", p_strFileName));


                Range FindRange = (Range)EditSheet.Columns["A", Type.Missing];
                //Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                if (!(FoundRangeContinued is null))
                {

                    int rowNumberContinued = FoundRangeContinued.Row + 1;
                    string rangeContinued = string.Format("1:{0}", rowNumberContinued.ToString());
                    EditSheet.Rows[rangeContinued].Delete();
                    log.Info(string.Format("Delete Continue = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));

                }

                Range FoundRangeNotes = FindRange.Find("Notes:");
                if (!(FoundRangeNotes is null))
                {
                    int RowNumberNotes = FoundRangeNotes.Row - 1;
                    string rangeNotes = string.Format("{0}:{1}", RowNumberNotes.ToString(), (RowNumberNotes + 50).ToString());
                    EditSheet.Rows[rangeNotes].Delete();
                    log.Info(string.Format("Delete Notes = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));
                }




                //EditSheet.Range["1","1:1"].Find()
                //book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.SaveAs(p_strFileName, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //book.Save();
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                EditSheet = null;
                book = null;


                log.Info(string.Format("Openning OleDBConnection to {0}.", p_strFileName));
                //log.Info(string.Format("Openning OleDBConnection to {0}.", strFileNameX));

                OleDbConnection OleDBconn = new OleDbConnection(string.Format(Utility.XLSConnectionString, p_strFileName));
                //OleDbConnection OleDBconn = new OleDbConnection(string.Format(Utility.XLSXConnectionString, strFileNameX));
                System.Data.DataTable dtXLS = new System.Data.DataTable();

                OleDbCommand OleCommand = new OleDbCommand(string.Format(Utility.ExcelSQLQuery, strSheetName, cur_data_date_COLUMN));
                //OleCommand.Parameters.Add(new OleDbParameter("Currency", "'Chinese yuan','Euro','Japanese yen','U.K. pound','U.K. Pound Sterling','U.S. dollar','Mexican peso','Canadian dollar','Australian dollar','Swedish krona','Norwegian krone'"));
                log.Info("QueryString for EXCEL - " + string.Format(Utility.ExcelSQLQuery, strSheetName, cur_data_date_COLUMN));
                OleCommand.Connection = OleDBconn;

                log.Info(string.Format("OleDBConnection assigned to OleCommand :{0}", p_strFileName));
                //log.Info(string.Format("OleDBConnection assigned to OleCommand :{0}", strFileNameX));

                OleDbDataAdapter adp = new OleDbDataAdapter(OleCommand);
                log.Info(string.Format("Fill Data Table with data from File :{0}", p_strFileName));
                //log.Info(string.Format("Fill Data Table with data from File :{0}", strFileNameX));
                adp.Fill(dtXLS);

                log.Info(string.Format("Data Table filled with data from File :{0}", p_strFileName));
                //log.Info(string.Format("Data Table filled with data from File :{0}", strFileNameX));
                OleDBconn.Close();

                log.Info(string.Format("Open file in Excel :{0}", p_strFileName));
                //log.Info(string.Format("Open file in Excel :{0}", strFileNameX));

                book = app.Workbooks.Open(p_strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //book = app.Workbooks.Open(strFileNameX, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                log.Info(string.Format("File openned in Excel :{0}", p_strFileName));
                //log.Info(string.Format("File openned in Excel :{0}", strFileNameX));


                book.Worksheets.Add();
                book.Worksheets["Sheet1"].Name = "DataToLoadFX";
                log.Info(string.Format("Loading data to {0}.", "DataToLoadFX"));
                Worksheet DataToLoadFX = book.Worksheets["DataToLoadFX"];
                DataToLoadFX.Cells[1, 1] = "Code";
                DataToLoadFX.Cells[1, 2] = "Category";
                DataToLoadFX.Cells[1, 3] = "CodeType";
                DataToLoadFX.Cells[1, 4] = "Date";
                DataToLoadFX.Cells[1, 5] = "Value";
                DataToLoadFX.Cells[1, 6] = "Type";
                DataToLoadFX.Cells[1, 7] = "Source";
                DataToLoadFX.Cells[1, 8] = "Email";
                DataToLoadFX.Cells[1, 9] = "Overwrite";
                string strdtLastUpdate = p_Date.ToString("MM/dd/yyyy");
                //DateTime dtLastUpdated = DateTime.ParseExact(strdtLastUpdate, "dd-MM-yyyy hh:mm",)
                //string filterString = string.Format("[{0}]", cur_data_date_COLUMN);
                //string filterString = string.Format("[{0}] = 'NA'", cur_data_date_COLUMN);
                int i = 2;
                int offset = dtXLS.Rows.Count;
                DataRow[] drUS = dtXLS.Select("Currency ='U.S. dollar'");
                double USD_val = System.Convert.ToDouble(drUS[0][2]);
                foreach (DataRow dr in dtXLS.Rows)
                {
                    DataToLoadFX.Cells[i, 1] = dr[1];
                    DataToLoadFX.Cells[i + offset, 1] = dr[1];
                    DataToLoadFX.Cells[i, 2] = "Currency";
                    DataToLoadFX.Cells[i + offset, 2] = "Currency";
                    DataToLoadFX.Cells[i, 3] = "Rogge.Code.Default";
                    DataToLoadFX.Cells[i + offset, 3] = "Rogge.Code.Default";
                    DataToLoadFX.Cells[i, 4] = strdtLastUpdate;
                    DataToLoadFX.Cells[i + offset, 4] = strdtLastUpdate;
                    string testValue = dr[1].ToString().Trim();
                    string tempCurrencyVal = dr[2].ToString().Trim();
                    if (tempCurrencyVal.Trim() != "NA" && tempCurrencyVal.Trim() != "")
                    {
                        Double CurrencyVal = Convert.ToDouble(tempCurrencyVal);
                        if (testValue.Equals("EUR") || testValue.Equals("GBP") || testValue.Equals("AUD"))
                            DataToLoadFX.Cells[i, 5] = USD_val / CurrencyVal;
                        else if (testValue.Equals("XDR"))
                            DataToLoadFX.Cells[i, 5] = USD_val;
                        else
                            DataToLoadFX.Cells[i, 5] = CurrencyVal / USD_val;

                    }
                    DataToLoadFX.Cells[i + offset, 5] = 0;
                    DataToLoadFX.Cells[i, 6] = "MID";
                    DataToLoadFX.Cells[i + offset, 6] = "BASK";
                    DataToLoadFX.Cells[i, 7] = "IMF";
                    DataToLoadFX.Cells[i + offset, 7] = "IMF";
                    DataToLoadFX.Cells[i, 8] = "ITDataAdmin@RoggeGlobal.com";
                    DataToLoadFX.Cells[i + offset, 8] = "ITDataAdmin@RoggeGlobal.com";
                    DataToLoadFX.Cells[i, 9] = "N";
                    DataToLoadFX.Cells[i + offset, 9] = "N";




                    i++;
                }
                i = i + offset;
                //book.Worksheets.Add();
                //book.Worksheets["Sheet2"].Name = "rms_sdrv";
                log.Info(string.Format("Loading data to {0}.", "rms_sdrv"));
                Worksheet wsRMS_SDRV = book.Worksheets["rms_sdrv"];

                //wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                ////wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                //wsRMS_SDRV.QueryTables[1].Name = "rms_sdrv";
                //wsRMS_SDRV.QueryTables[1].FieldNames = true;
                //wsRMS_SDRV.QueryTables[1].RowNumbers = false;
                //wsRMS_SDRV.QueryTables[1].FillAdjacentFormulas = false;
                //wsRMS_SDRV.QueryTables[1].WebFormatting = XlWebFormatting.xlWebFormattingNone;
                ////wsRMS_SDRV.QueryTables[1].PreserveFormatting = false;
                //wsRMS_SDRV.QueryTables[1].SaveData = true;
                //wsRMS_SDRV.QueryTables[1].Refresh(false);

                Range ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                Range rgCurrencyUnitFind = ARange.Find("Currency Unit");

                int CurrencyRow = rgCurrencyUnitFind.Row;
                string deleteCurrencyRows = string.Format("1:{0}", (CurrencyRow - 1).ToString());
                wsRMS_SDRV.Rows[deleteCurrencyRows].Delete();


                ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                Range rgNotes = (Range)ARange.Find("Footnotes");
                if (!(rgNotes is null))
                {
                    int NotesRow = rgNotes.Row;
                    string deleteNotesRows = string.Format("{0}:{1}", (NotesRow - 1).ToString(), (NotesRow + 30).ToString());
                    wsRMS_SDRV.Rows[deleteNotesRows].Delete();
                }


                //switch()
                //{

                //}

                for (int j = 2; j < 10; j++)
                {
                    double result;
                    var TempValParse = (wsRMS_SDRV.Cells[j, 1] as Range).Value;
                    string strTempValParse = TempValParse.ToString();

                    var tempVal = Double.TryParse(strTempValParse, out result) ? "SDR1=US$" : strTempValParse;
                    if (tempVal != "U.S.$1.00 = SDR" && tempVal != "U.S. dollar" && !Double.TryParse(strTempValParse, out result))
                    {
                        switch (tempVal)
                        {
                            case "Chinese yuan":
                                DataToLoadFX.Cells[i, 1] = "XDRCNY";
                                break;
                            case "Euro":
                                DataToLoadFX.Cells[i, 1] = "XDREUR";
                                break;
                            case "Japanese yen":
                                DataToLoadFX.Cells[i, 1] = "XDRJPY";
                                break;
                            case "U.K. pound":
                                DataToLoadFX.Cells[i, 1] = "XDRGBP";
                                break;
                            case "SDR1 = US$":
                                DataToLoadFX.Cells[i, 1] = "XDRUSD";
                                break;

                            default:
                                break;
                        }
                        DataToLoadFX.Cells[i, 2] = "Index";
                        DataToLoadFX.Cells[i, 3] = "Rogge.Code.Default";
                        DataToLoadFX.Cells[i, 4] = strdtLastUpdate;
                        if (tempVal.Equals("SDR1 = US$"))
                        {
                            string tempXDRUSA = (string)(wsRMS_SDRV.Cells[j, 4] as Range).Value;
                            int lentempXDRUSA = tempXDRUSA.Length - 2;
                            tempXDRUSA = tempXDRUSA.Substring(0, lentempXDRUSA);

                            double tempXDRUSADBL = Convert.ToDouble(tempXDRUSA);
                            DataToLoadFX.Cells[i, 5] = tempXDRUSADBL;

                        }
                        else
                            DataToLoadFX.Cells[i, 5] = wsRMS_SDRV.Cells[j, 3];
                        DataToLoadFX.Cells[i, 6] = "MID";
                        DataToLoadFX.Cells[i, 7] = "IMF";
                        DataToLoadFX.Cells[i, 8] = "ITDataAdmin@RoggeGlobal.com";
                        DataToLoadFX.Cells[i, 9] = "N";
                        i++;
                    }

                }

                //book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Save();
                

                app.Quit();
                app = null;

                dtXLS = null;

                //return strFileNameX;
                return p_strFileName;

            }
            catch (Exception e)
            {
                app.Workbooks.Close();
                app.Quit();
                log.Info(e.Message);
                throw new Exception(e.Message, e.InnerException);
            }
        }
            public string ProcessAndLoad(string p_strFileName, DateTime p_Date)
            {
                Application app = new Application();
                app.Visible = false;
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                try
                {
                    SqlConnection conn = Utility.GetDataConn(Utility.RIMSDBConnectionString);
                    log.Info(string.Format("Connected to connection string {0}", Utility.RIMSDBConnectionString));
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt = Utility.getData(conn, Utility.SQLStatementProcDate);
                    System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
                    string strMonthName = mfi.GetMonthName(p_Date.Month);
                    string strYear = p_Date.Year.ToString();
                    string strFilterOne = string.Format("Currency units per SDR for {0} {1}", strMonthName, strYear);
                    log.Info(string.Format("When date is 15 or < filter is = {0}.", strFilterOne));
                    string strFilterTwo = string.Format("Currency units per SDR for {0} {1} Continued", strMonthName, strYear);
                    log.Info(string.Format("Date > = Mid Month filter is = {0}.", strFilterTwo));

                    string cur_data_date_FILENAME = p_Date.ToString("yyyyMMdd");

                    log.Info(string.Format(" Retrieving  {0}.", cur_data_date_FILENAME));
                    //string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
                    string cur_data_date_COLUMN = p_Date.ToString("d-MMM-yy");
                    string cur_data_date_search = p_Date.ToString("M/d/yyyy");
                    string strFileNameX = string.Format(Utility.FilePathSaveTo, cur_data_date_FILENAME);
                    string strSheetName = string.Format(Utility.SheetName, cur_data_date_FILENAME);
                    string strSheetNameIDX = string.Format(Utility.SheetNameIDX, cur_data_date_FILENAME);


                Workbook book = null;
                    book = app.Workbooks.Open(p_strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    log.Info(string.Format("{0} opened.", p_strFileName));
                    Worksheet EditSheet = book.Worksheets[strSheetNameIDX];
                    EditSheet.Rows["1:1"].Delete();
                    log.Info(string.Format("Delete header row File = {0}", p_strFileName));


                    Range FindRange = (Range)EditSheet.Columns["A", Type.Missing];
                    //Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                    Range FoundRangeContinued = FindRange.Find(strFilterTwo);
                    if (!(FoundRangeContinued is null))
                    {

                        int rowNumberContinued = FoundRangeContinued.Row + 1;
                        string rangeContinued = string.Format("1:{0}", rowNumberContinued.ToString());
                        EditSheet.Rows[rangeContinued].Delete();
                        log.Info(string.Format("Delete Continue = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));

                    }

                    Range FoundRangeNotes = FindRange.Find("Notes:");
                    if (!(FoundRangeNotes is null))
                    {
                        int RowNumberNotes = FoundRangeNotes.Row - 1;
                        string rangeNotes = string.Format("{0}:{1}", RowNumberNotes.ToString(), (RowNumberNotes + 50).ToString());
                        EditSheet.Rows[rangeNotes].Delete();
                        log.Info(string.Format("Delete Notes = {0} : Checkpoint {1}", strFilterTwo, p_strFileName));
                    }

                    book.SaveAs(p_strFileName, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    book.Save();
                    
                    
                    
                    log.Info(string.Format("Find Date in header row {0} For Sheet {1}.", cur_data_date_search,EditSheet.Name));
                    FindRange = (Range)EditSheet.Rows["1:1"];
                    log.Info(string.Format("Find Range set for date header search."));

                    Range FoundRangeTheDate = FindRange.Find(cur_data_date_search);
                    int ColumnLimit = FoundRangeTheDate.Column;

                    log.Info(string.Format("Date found in column {0}. This is the data column.", ColumnLimit.ToString()));

                    FindRange = (Range)EditSheet.Columns["A:A"];

                    Range FoundRangeUSD = FindRange.Find("U.S. dollar");
                
                    //Range FoundRangelastRow = FoundRangeTheDate.Rows["1:" + EditSheet.Cells.Rows.Count.ToString()].End[XlDirection.xlUp];
                    Range FoundRangelastRowbyRow = EditSheet.Cells.Find("*", After: EditSheet.Range["A1"], LookAt: XlLookAt.xlPart, LookIn: XlFindLookIn.xlFormulas, 
                        SearchOrder: XlSearchOrder.xlByRows,SearchDirection:XlSearchDirection.xlPrevious, MatchCase: false);
                    
                    
                    
                    int Rowlimit = FoundRangelastRowbyRow.Row;
                    
                    int USD_row= FoundRangeUSD.Row;

                    log.Info(string.Format("USD FX Rate found in Row {0} Column {1}.", USD_row.ToString(), ColumnLimit.ToString()));

                    double USD_val = Convert.ToDouble(EditSheet.Cells[USD_row, ColumnLimit].Value);


                    //ProcessAndSave the book

                //book.Close(Type.Missing, Type.Missing, Type.Missing);

                
                    book.Worksheets.Add();
                    book.Worksheets["Sheet1"].Name = "DataToLoadFX";
                    log.Info(string.Format("Loading data to {0}.", "DataToLoadFX"));
                    Worksheet DataToLoadFX = book.Worksheets["DataToLoadFX"];
                    DataToLoadFX.Cells[1, 1] = "Code";
                    DataToLoadFX.Cells[1, 2] = "Category";
                    DataToLoadFX.Cells[1, 3] = "CodeType";
                    DataToLoadFX.Cells[1, 4] = "Date";
                    DataToLoadFX.Cells[1, 5] = "Value";
                    DataToLoadFX.Cells[1, 6] = "Type";
                    DataToLoadFX.Cells[1, 7] = "Source";
                    DataToLoadFX.Cells[1, 8] = "Email";
                    DataToLoadFX.Cells[1, 9] = "Overwrite";
                    string strdtLastUpdate = p_Date.ToString("MM/dd/yyyy");
                    log.Info(string.Format("Parse data to spread sheet {0}.", "DataToLoadFX"));
                    int offset = 10;
                    int i = 2;
                    int j = 2;
                    for (i = 2; i <= Rowlimit; i++)
                    {
                        bool update = false;
                        switch (EditSheet.Cells[i, 1].Value)
                        {
                            case "Chinese yuan":
                                update = true;
                                DataToLoadFX.Cells[j,1] = "CNY";
                                DataToLoadFX.Cells[j + offset, 1] = "CNY";
                                break;
                            case "Euro":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "EUR";
                                DataToLoadFX.Cells[j + offset, 1] = "EUR";
                                break;
                            case "Japanese yen":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "JPY";
                                DataToLoadFX.Cells[j + offset, 1] = "JPY";
                                break;
                            case "U.K. pound":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "GBP";
                                DataToLoadFX.Cells[j + offset, 1] = "GBP";
                                break;
                            case "U.K. Pound Sterling":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "GBP";
                                DataToLoadFX.Cells[j + offset, 1] = "GBP";
                                break;
                            case "U.S. dollar":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "XDR";
                                DataToLoadFX.Cells[i + offset, 1] = "XDR";
                                
                                break;
                            case "Australian dollar":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "AUD";
                                DataToLoadFX.Cells[j + offset, 1] = "AUD";;
                                break;
                            case "Canadian dollar":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "CAD";
                                DataToLoadFX.Cells[j + offset, 1] = "CAD";
                                break;
                            case "Mexican peso":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "MXN";
                                DataToLoadFX.Cells[j + offset, 1] = "MXN";
                                break;
                            case "Norwegian krone":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "NOK";
                                DataToLoadFX.Cells[j + offset, 1] = "NOK";
                                break;
                            case "Swedish krona":
                                update = true;
                                DataToLoadFX.Cells[j, 1] = "SEK";
                                DataToLoadFX.Cells[j + offset, 1] = "SEK";
                                break;
                            default:
                                break;

                        }
                        if (update)
                        {
                            DataToLoadFX.Cells[j, 2] = "Currency";
                            DataToLoadFX.Cells[j + offset, 2] = "Currency";
                            DataToLoadFX.Cells[j, 3] = "Rogge.Code.Default";
                            DataToLoadFX.Cells[j + offset, 3] = "Rogge.Code.Default";
                            DataToLoadFX.Cells[j, 4] = strdtLastUpdate;
                            DataToLoadFX.Cells[j + offset, 4] = strdtLastUpdate;
                            string testValue = EditSheet.Cells[i, 1].Value;
                            string tempCurrencyVal = EditSheet.Cells[i, ColumnLimit].Value.ToString();
                            if (tempCurrencyVal.Trim() != "NA" && tempCurrencyVal.Trim() != "")
                            {
                                Double CurrencyVal = Convert.ToDouble(tempCurrencyVal);
                                if (EditSheet.Cells[i, 1].Value.Equals("Euro") || EditSheet.Cells[i, 1].Value.Equals("U.K. pound") || EditSheet.Cells[i, 1].Value.Equals("Australian dollar") || EditSheet.Cells[i, 1].Value.Equals("U.K. Pound Sterlin"))
                                    DataToLoadFX.Cells[i, 5] = USD_val / CurrencyVal;
                                else if (testValue.Equals("XDR"))
                                    DataToLoadFX.Cells[j, 5] = USD_val;
                                else
                                    DataToLoadFX.Cells[j, 5] = CurrencyVal / USD_val;

                            }
                            DataToLoadFX.Cells[j + offset, 5] = 0;
                            DataToLoadFX.Cells[j, 6] = "MID";
                            DataToLoadFX.Cells[j + offset, 6] = "BASK";
                            DataToLoadFX.Cells[j, 7] = "IMF";
                            DataToLoadFX.Cells[j + offset, 7] = "IMF";
                            DataToLoadFX.Cells[j, 8] = "ITDataAdmin@RoggeGlobal.com";
                            DataToLoadFX.Cells[j + offset, 8] = "ITDataAdmin@RoggeGlobal.com";
                            DataToLoadFX.Cells[j, 9] = "N";
                            DataToLoadFX.Cells[j + offset, 9] = "N";
                            j++;
                        }



                    }


                //DateTime dtLastUpdated = DateTime.ParseExact(strdtLastUpdate, "dd-MM-yyyy hh:mm",)
                //string filterString = string.Format("[{0}]", cur_data_date_COLUMN);
                //string filterString = string.Format("[{0}] = 'NA'", cur_data_date_COLUMN);

                // dtXLS.Rows.Count;
                   /* DataRow[] drUS = dtXLS.Select("Currency ='U.S. dollar'");
                    double USD_val = System.Convert.ToDouble(drUS[0][2]);
                    foreach (DataRow dr in dtXLS.Rows)
                    {
                        DataToLoadFX.Cells[i, 1] = dr[1];
                        DataToLoadFX.Cells[i + offset, 1] = dr[1];
                        DataToLoadFX.Cells[i, 2] = "Currency";
                        DataToLoadFX.Cells[i + offset, 2] = "Currency";
                        DataToLoadFX.Cells[i, 3] = "Rogge.Code.Default";
                        DataToLoadFX.Cells[i + offset, 3] = "Rogge.Code.Default";
                        DataToLoadFX.Cells[i, 4] = strdtLastUpdate;
                        DataToLoadFX.Cells[i + offset, 4] = strdtLastUpdate;
                        string testValue = dr[1].ToString().Trim();
                        string tempCurrencyVal = dr[2].ToString().Trim();
                        if (tempCurrencyVal.Trim() != "NA" && tempCurrencyVal.Trim() != "")
                        {
                            Double CurrencyVal = Convert.ToDouble(tempCurrencyVal);
                            if (testValue.Equals("EUR") || testValue.Equals("GBP") || testValue.Equals("AUD"))
                                DataToLoadFX.Cells[i, 5] = USD_val / CurrencyVal;
                            else if (testValue.Equals("XDR"))
                                DataToLoadFX.Cells[i, 5] = USD_val;
                            else
                                DataToLoadFX.Cells[i, 5] = CurrencyVal / USD_val;

                        }
                        DataToLoadFX.Cells[i + offset, 5] = 0;
                        DataToLoadFX.Cells[i, 6] = "MID";
                        DataToLoadFX.Cells[i + offset, 6] = "BASK";
                        DataToLoadFX.Cells[i, 7] = "IMF";
                        DataToLoadFX.Cells[i + offset, 7] = "IMF";
                        DataToLoadFX.Cells[i, 8] = "ITDataAdmin@RoggeGlobal.com";
                        DataToLoadFX.Cells[i + offset, 8] = "ITDataAdmin@RoggeGlobal.com";
                        DataToLoadFX.Cells[i, 9] = "N";
                        DataToLoadFX.Cells[i + offset, 9] = "N";




                        i++;
                    }*/
                    j = j + offset;
                    //book.Worksheets.Add();
                    //book.Worksheets["Sheet2"].Name = "rms_sdrv";
                    log.Info(string.Format("Loading data to {0}.", "rms_sdrv"));
                    Worksheet wsRMS_SDRV = book.Worksheets["rms_sdrv"];

                    //wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                    ////wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
                    //wsRMS_SDRV.QueryTables[1].Name = "rms_sdrv";
                    //wsRMS_SDRV.QueryTables[1].FieldNames = true;
                    //wsRMS_SDRV.QueryTables[1].RowNumbers = false;
                    //wsRMS_SDRV.QueryTables[1].FillAdjacentFormulas = false;
                    //wsRMS_SDRV.QueryTables[1].WebFormatting = XlWebFormatting.xlWebFormattingNone;
                    ////wsRMS_SDRV.QueryTables[1].PreserveFormatting = false;
                    //wsRMS_SDRV.QueryTables[1].SaveData = true;
                    //wsRMS_SDRV.QueryTables[1].Refresh(false);

                    Range ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                    Range rgCurrencyUnitFind = ARange.Find("Currency Unit");

                    int CurrencyRow = rgCurrencyUnitFind.Row;
                    string deleteCurrencyRows = string.Format("1:{0}", (CurrencyRow - 1).ToString());
                    wsRMS_SDRV.Rows[deleteCurrencyRows].Delete();


                    ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
                    Range rgNotes = (Range)ARange.Find("Footnotes");
                    if (!(rgNotes is null))
                    {
                        int NotesRow = rgNotes.Row;
                        string deleteNotesRows = string.Format("{0}:{1}", (NotesRow - 1).ToString(), (NotesRow + 30).ToString());
                        wsRMS_SDRV.Rows[deleteNotesRows].Delete();
                    }


                    //switch()
                    //{

                    //}

                    for (i = 2; i < 10; i++)
                    {
                        double result;
                        var TempValParse = (wsRMS_SDRV.Cells[i, 1] as Range).Value;
                        string strTempValParse = TempValParse.ToString();

                        var tempVal = Double.TryParse(strTempValParse, out result) ? "SDR1=US$" : strTempValParse;
                        if (tempVal != "U.S.$1.00 = SDR" && tempVal != "U.S. dollar" && !Double.TryParse(strTempValParse, out result))
                        {
                            switch (tempVal)
                            {
                                case "Chinese yuan":
                                    DataToLoadFX.Cells[j, 1] = "XDRCNY";
                                    break;
                                case "Euro":
                                    DataToLoadFX.Cells[j, 1] = "XDREUR";
                                    break;
                                case "Japanese yen":
                                    DataToLoadFX.Cells[j, 1] = "XDRJPY";
                                    break;
                                case "U.K. pound":
                                    DataToLoadFX.Cells[j, 1] = "XDRGBP";
                                    break;
                                case "SDR1 = US$":
                                    DataToLoadFX.Cells[j, 1] = "XDRUSD";
                                    break;

                                default:
                                    break;
                            }
                            DataToLoadFX.Cells[j, 2] = "Index";
                            DataToLoadFX.Cells[j, 3] = "Rogge.Code.Default";
                            DataToLoadFX.Cells[j, 4] = strdtLastUpdate;
                            if (tempVal.Equals("SDR1 = US$"))
                            {
                                string tempXDRUSA = (string)(wsRMS_SDRV.Cells[i, 4] as Range).Value;
                                int lentempXDRUSA = tempXDRUSA.Length - 2;
                                tempXDRUSA = tempXDRUSA.Substring(0, lentempXDRUSA);

                                double tempXDRUSADBL = Convert.ToDouble(tempXDRUSA);
                                DataToLoadFX.Cells[j, 5] = tempXDRUSADBL;

                            }
                            else
                                DataToLoadFX.Cells[j, 5] = wsRMS_SDRV.Cells[i, 3];
                            DataToLoadFX.Cells[j, 6] = "MID";
                            DataToLoadFX.Cells[j, 7] = "IMF";
                            DataToLoadFX.Cells[j, 8] = "ITDataAdmin@RoggeGlobal.com";
                            DataToLoadFX.Cells[j, 9] = "N";
                            j++;
                        }

                    }
                    if (Utility.FileFormat.Equals("PROCLOADX"))
                        book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    else
                        strFileNameX = p_strFileName;
                    book.Save();

                    EditSheet = null;
                    book = null;

                    app.Quit();
                    app = null;

                    //dtXLS = null;

                    return strFileNameX;
                    

                }
                catch (Exception e)
                {
                    app.Workbooks.Close();
                    app.Quit();
                    log.Info(e.Message);
                    throw new Exception(e.Message, e.InnerException);
                }
            }
        public void LoadToRGP(string p_strFileNameFull, string p_FileName,string p_FileFormat)
        {
            
            string strConnection = Utility.ConnectionString;
            SqlConnection RGPConn = Utility.GetDataConn(strConnection);

            SqlDataAdapter da = new SqlDataAdapter(Utility.SQLQuery, RGPConn);

            System.Data.DataTable dtTimeSeries = new System.Data.DataTable();
            da.Fill(dtTimeSeries);
            System.Data.OleDb.OleDbConnection OleDBconn = null;
            switch (p_FileFormat)
            {
                case "XLSX":
                    OleDBconn = new System.Data.OleDb.OleDbConnection(string.Format(Utility.XLSXConnectionString, p_strFileNameFull));
                    break;
                default:
                    OleDBconn = new System.Data.OleDb.OleDbConnection(string.Format(Utility.XLSConnectionString, p_strFileNameFull));
                    break;
            }

            OleDbCommand  OleCommand = new OleDbCommand("select * FROM [DataToLoadFX$]");
            OleCommand.Connection = OleDBconn;
            OleDbDataAdapter adp = new OleDbDataAdapter(OleCommand);
            System.Data.DataTable dtXLS = new System.Data.DataTable();
            adp.Fill(dtXLS);
            DataColumn FileNameColumn = new DataColumn("FileName");
            FileNameColumn.DefaultValue = p_FileName;
            dtXLS.Columns.Add(FileNameColumn);
            int i;
            i = 0;

            SqlBulkCopy bc;
            bc = new SqlBulkCopy(RGPConn, SqlBulkCopyOptions.TableLock, null);
            bc.DestinationTableName = Utility.Table;
            log.Info(string.Format("map data columns for table ({0})", Utility.Table));
            foreach (DataColumn dc in dtXLS.Columns)
            {
                SqlBulkCopyColumnMapping mapCol = new SqlBulkCopyColumnMapping(dtXLS.Columns[i].ColumnName, dc.ColumnName);
                bc.ColumnMappings.Add(mapCol);

                i++;
            }

            if (Utility.Truncate)
            {
                SqlCommand cmd = new SqlCommand(string.Format("DELETE {0}", Utility.Table), RGPConn);
                log.Info(string.Format("Truncate table  ({0}) before import", Utility.Table));
                cmd.ExecuteNonQuery();
            }
            int rowcount = 0;
            rowcount = dtXLS.Rows.Count;
            int iRowCount = dtXLS.Rows.Count;

            bc.BatchSize = 1000;
            bc.BulkCopyTimeout = 0;
            log.Info(string.Format("Bulk Insert of {0} records started", iRowCount.ToString()));
            log.Info(string.Format("Load {0} rows of data to {1}", rowcount.ToString(), Utility.Table));

            bc.WriteToServer(dtXLS);
            log.Info(string.Format("Data Inserted. {0} records inserted.", iRowCount.ToString()));
            rowcount = dtXLS.Rows.Count;

            //var res = from row in dtXLS.AsEnumerable()
            //      where row.Field<string>("Currency").Equals(strFilterTwo)
            //      select row;

            //var res1 = dtXLS.Select("Currency like '%Continued%'");



            //foreach(DataRow dr in dtXLS.Rows)
            //{

            //}
        }
    }
}
