using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop.Excel;
using log4net;

namespace IMF_WebRequestClient
{
    public class WebRequest_GetFile
    {
        ILog log = null;
        public WebRequest_GetFile(ILog p_log)
        {
            log = p_log + log;
        }
        public void GET_FILE(string p_URL)
        {
            WebClient wcIMF_Request = new WebClient();
            System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
            string strMonthName = mfi.GetMonthName(DateTime.Now.Month);
            string strYear = DateTime.Now.Year.ToString();
            string strFilterOne = string.Format("Currency units per SDR for {0} {1}", strMonthName, strYear);
            log.Info(strFilterOne);
            string strFilterTwo = string.Format("Currency units per SDR for {0} {1} Continued", strMonthName, strYear);
            log.Info(strFilterTwo);
            string cur_data_date = DateTime.Now.ToString("yyyy-MM-dd");
            string cur_data_date_FILENAME = DateTime.Now.ToString("yyyyMMdd");
            string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
            string strFileName = string.Format("\\\\us1.1corp.org\\dc0\\RCM\\prod\\apps\\rimsx\\CCME\\IMF_data_{0}_WEBCLIENT.xls", cur_data_date_FILENAME);
            string strFileNameX = string.Format("\\\\us1.1corp.org\\dc0\\RCM\\prod\\apps\\rimsx\\CCME\\IMF_data_{0}_WEBCLIENT.xlsx", cur_data_date_FILENAME);
            string strSheetName = string.Format("[IMF_data_{0}_WEBCLIENT$]", cur_data_date_FILENAME);
            wcIMF_Request.DownloadFile(string.Format("https://www.imf.org/external/np/fin/data/rms_mth.aspx?SelectDate={0}&reportType=CVSDR&tsvflag=Y", cur_data_date), strFileName);
            //wcIMF_Request.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            //byte[] response = wcIMF_Request.UploadData(string.Format("https://www.imf.org/external/np/fin/data/rms_mth.aspx?SelectDate={0}&reportType=CVSDR&tsvflag=Y", cur_data_date), "POST", new byte[] { });
            //get the response as a string and do something with it...
            //string s = System.Text.Encoding.ASCII.GetString(response);

            //System.IO.StreamWriter Imf_file = new System.IO.StreamWriter("\\\\us1.1corp.org\\dc0\\RCM\\prod\\apps\\rimsx\\CCME\\IMF_data.xls");


            //Imf_file.Write(s);
            //Imf_file.Close();

            //System.IO.File.WriteAllLines(strFileNameX, System.IO.File.ReadAllLines(strFileNameX).Where(strLine => !strLine.Equals(strFilterOne)));
            //System.IO.File.WriteAllLines(strFileName, System.IO.File.ReadAllLines(strFileName).Where(strLine => !strLine.Equals(strFilterOne)));

            //System.IO.File.WriteAllLines(strFileNameX, System.IO.File.ReadAllLines(strFileNameX).Where(strLine => !strLine.Equals(strFilterTwo)));
            //System.IO.File.WriteAllLines(strFileName, System.IO.File.ReadAllLines(strFileName).Where(strLine => !strLine.Equals(strFilterTwo)));

            //System.IO.File.WriteAllLines(strFileNameX, System.IO.File.ReadAllLines(strFileNameX).Where(strLine => !strLine.Equals("")));
            //System.IO.File.WriteAllLines(strFileName, System.IO.File.ReadAllLines(strFileName).Where(strLine => !strLine.Equals("")));



            Application app = new Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;

            Workbook book = null;
            book = app.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet EditSheet = book.Worksheets[1];
            EditSheet.Rows["1:1"].Delete();

            Range FindRange = (Range)EditSheet.Columns["A", Type.Missing];
            //Range FoundRangeContinued = FindRange.Find(strFilterTwo);
            Range FoundRangeContinued = FindRange.Find(strFilterTwo);
            if (!(FoundRangeContinued is null))
            {

                int rowNumberContinued = FoundRangeContinued.Row + 1;
                string rangeContinued = string.Format("1:{0}", rowNumberContinued.ToString());
                EditSheet.Rows[rangeContinued].Delete();
            }

            Range FoundRangeNotes = FindRange.Find("Notes:");
            if (!(FoundRangeNotes is null))
            {
                int RowNumberNotes = FoundRangeNotes.Row - 1;
                string rangeNotes = string.Format("{0}:{1}", RowNumberNotes.ToString(), (RowNumberNotes + 50).ToString());
                EditSheet.Rows[rangeNotes].Delete();
            }


            

            //EditSheet.Range["1","1:1"].Find()
            book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            book.Close(Type.Missing, Type.Missing, Type.Missing);
            EditSheet = null;
            book = null;




            OleDbConnection OleDBconn = new OleDbConnection(string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=1\";", strFileNameX));
            System.Data.DataTable dtXLS = new System.Data.DataTable();
            OleDbCommand OleCommand = new OleDbCommand(string.Format("select [Currency], " +
                "IIF(Currency = 'Chinese yuan', 'CNY',IIF(Currency = 'Euro', 'EUR',IIF(Currency = 'Japanese yen', 'JPY',IIF(Currency = 'U.K. pound','GBP',IIF(Currency = 'U.K. Pound Sterling','GBP',IIF(Currency = 'U.S. dollar','XDR',)))))) , " +
                //"IIF(Currency = 'Japanese yen', 'JPY', " +
                //"IIF(Currency = 'U.K. pound','GBP'," +
                //"IIF(Currency = 'U.K. Pound Sterling','GBP'," +
                //"IIF(Currency = 'U.S. dollar','XDR'," +
                //"IIF(Currency = 'Australian dollar,'AUD'," +
                //"IIF(Currency = 'Canadian dollar','CAD'," +
                //"IIF(Currency = 'Mexican peso','MXN', " +
                //"IIF(Currency = 'Norwegian krone','NOK', " +
                //"IIF(Currency = 'Swedish krona','SEK',''))))))))))) as [Currency_cd]," +
                "[{1}] from {0} WHERE [Currency] IN ('Chinese yuan','Euro','Japanese yen','U.K. pound','U.K. Pound Sterling','U.S. dollar'" +
                ",'Mexican peso','Canadian dollar','Australian dollar','Swedish krona','Norwegian krone')", strSheetName, cur_data_date_COLUMN));
            OleCommand.Connection = OleDBconn;
            OleDbDataAdapter adp = new OleDbDataAdapter(OleCommand);
            adp.Fill(dtXLS);
            OleDBconn.Close();
            book = app.Workbooks.Open(strFileNameX, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

           

            book.Worksheets.Add();
            book.Worksheets["Sheet1"].Name = "DataToLoadFX";
            Worksheet DataToLoadFX = book.Worksheets["DataToLoadFX"];
            DataToLoadFX.Cells[1, 1] = "AS_OF_DATE";
            DataToLoadFX.Cells[1, 2] = "Currency";
            DataToLoadFX.Cells[1, 3] = "Currency_CD";
            DataToLoadFX.Cells[1, 4] = cur_data_date_COLUMN;
            string strdtLastUpdate = DateTime.Now.ToString("dd-MM-yyyy hh:mm");
            //DateTime dtLastUpdated = DateTime.ParseExact(strdtLastUpdate, "dd-MM-yyyy hh:mm",)
            //string filterString = string.Format("[{0}]", cur_data_date_COLUMN);
            //string filterString = string.Format("[{0}] = 'NA'", cur_data_date_COLUMN);
            int i = 2;

            foreach (DataRow dr in dtXLS.Rows)
            {
                DataToLoadFX.Cells[i, 1] = strdtLastUpdate;
                DataToLoadFX.Cells[i, 2] = dr[0];
                DataToLoadFX.Cells[i, 3] = dr[1];
                if (dr[2].ToString().Trim() != "NA")
                    DataToLoadFX.Cells[i, 4] = dr[2];

                i++;
            }

            book.Worksheets.Add();
            book.Worksheets["Sheet2"].Name = "rms_sdrv";

            Worksheet wsRMS_SDRV = book.Worksheets["rms_sdrv"];

            wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
            //wsRMS_SDRV.QueryTables.Add("URL;http://www.imf.org/external/np/fin/data/rms_sdrv.aspx", wsRMS_SDRV.Range["$A$1"]);
            wsRMS_SDRV.QueryTables[1].Name = "rms_sdrv";
            wsRMS_SDRV.QueryTables[1].FieldNames = true;
            wsRMS_SDRV.QueryTables[1].RowNumbers = false;
            wsRMS_SDRV.QueryTables[1].FillAdjacentFormulas = false;
            wsRMS_SDRV.QueryTables[1].WebFormatting = XlWebFormatting.xlWebFormattingNone;
            //wsRMS_SDRV.QueryTables[1].PreserveFormatting = false;
            wsRMS_SDRV.QueryTables[1].SaveData = true;
            wsRMS_SDRV.QueryTables[1].Refresh(false);
            
            Range ARange = (Range)wsRMS_SDRV.Columns["A", Type.Missing];
            Range rgCurrencyUnitFind = ARange.Find("Currency Unit");

            int  CurrencyRow =  rgCurrencyUnitFind.Row;
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


            book.SaveAs(strFileNameX, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
            app = null;

            dtXLS = null;
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
