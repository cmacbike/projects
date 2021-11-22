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
            log = p_log;
        }
        public string GET_FILE(DateTime p_Date)
        {
            Application app = new Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;
            try
            {
                log.Info("create webclient.");
                WebClient wcIMF_Request = new WebClient();
                
                string cur_data_date = p_Date.ToString("yyyy-MM-dd");
                log.Info(string.Format("Current data date - {0}", cur_data_date));

                string cur_data_date_FILENAME = p_Date.ToString("yyyyMMdd");
                //string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
                string strFileName = string.Format(Utility.FilePath, cur_data_date_FILENAME);

                log.Info(string.Format("File to retrieve - {0}", strFileName));
                log.Info(string.Format("URL for file to retrieve {0}", string.Format(Utility.URL, cur_data_date)));
                wcIMF_Request.DownloadFile(string.Format(Utility.URL, cur_data_date), strFileName);
                Workbook book = null;
                Worksheet wsRMS_SDRV = null;


                if (System.IO.File.Exists(strFileName))
                {
                    log.Info(string.Format("File retrieved successfully - {0}", strFileName));
                    book = app.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    log.Info(string.Format("{0} opened.", strFileName));
                    book.Worksheets.Add();
                    book.Worksheets["Sheet1"].Name = "rms_sdrv";
                    log.Info(string.Format("Loading data to {0}.", "rms_sdrv"));
                    wsRMS_SDRV = book.Worksheets["rms_sdrv"];


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
                    log.Info(string.Format("Data loaded to {0}.", "rms_sdrv"));
                }
                book.SaveAs(strFileName, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Save();
                book = null;

                app.Quit();
                app = null;
                return strFileName;
            }
            catch(Exception e)
            {
                app.Workbooks.Close();
                app.Quit();
                log.Info(e.Message);
                throw new Exception(e.Message,e.InnerException);
                
            }
        }

    }

}
