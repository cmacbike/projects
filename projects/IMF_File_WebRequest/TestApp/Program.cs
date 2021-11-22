using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Data;
using System.Data.SqlClient;
using IMF_WebRequestClient;
using System.Windows.Forms;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace IMF_WebRequestClient
{
    class Program
    {
        static string m_ConfigFileName;
        static int Main(string[] args)
        {
            ILog log = null;
            try
            {

                int i;
                for (i = 0; i < args.Length; i++)
                {
                    switch (args[i])
                    {

                        //case "-l":
                        //    strlogFileLOC = args[i + 1];
                        //    break;
                        case "-c":
                            m_ConfigFileName = args[i + 1];
                            Utility.Configfile = m_ConfigFileName;
                            break;

                    }
                }

                log4net.GlobalContext.Properties["LogFileName"] = Utility.LogFilePath;
                Console.WriteLine(string.Format("logging configured for {0}...", log4net.GlobalContext.Properties["LogFileName"]));
                log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(m_ConfigFileName));
                log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                Console.WriteLine(string.Format("Logging is configured at {0}.", log.Logger.Name));
                SqlConnection RIMSConn = Utility.GetDataConn(Utility.RIMSDBConnectionString);



                DataTable proc_dt = Utility.getData(RIMSConn, Utility.SQLStatementProcDate);

                DateTime cur_proc_dt = (DateTime)proc_dt.Rows[0][Utility.ProcDate];



                DateTime the_Date = Utility.DateSpecified ? Utility.SpecifyDate : DateTime.Now;

                the_Date = Utility.use_Proc_date ? cur_proc_dt : the_Date;
                DateTime TodaysDate = DateTime.Now;

                

                if (Utility.use_Proc_date)
                {
                    switch (TodaysDate.DayOfWeek)
                    {
                        case DayOfWeek.Saturday:
                            if (the_Date.DayOfWeek != DayOfWeek.Friday)
                            {
                                the_Date = DateTime.Now.AddDays(-1);
                            }
                            break;
                        case DayOfWeek.Sunday:
                            if (the_Date.DayOfWeek != DayOfWeek.Friday)
                            {
                                the_Date = DateTime.Now.AddDays(-2);
                            }
                            break;
                        case DayOfWeek.Monday:
                            if (the_Date.DayOfWeek != DayOfWeek.Friday)
                            {
                                the_Date = DateTime.Now.AddDays(-3);
                            }
                            break;
                    }
                }

                log.Info(the_Date);

                if (Utility.WebRequest)
                {
                    WebRequest_GetFile GetFileClass = new WebRequest_GetFile(log);
                    string strFileName = GetFileClass.GET_FILE(the_Date);
                }
                if (Utility.LoadFile)
                {
                    string cur_data_date_FILENAME = the_Date.ToString("yyyyMMdd");
                    //string cur_data_date_COLUMN = DateTime.Now.AddDays(-1).ToString("d-MMM-yy");
                    string strFileName = string.Format(Utility.FilePath, cur_data_date_FILENAME);
                    ProcessFile objProcessFile = new ProcessFile(log);

                    string strFileNameX = "";
                    string FileFormatX = "XSLX";

                    switch (Utility.FileFormat)
                    {
                        case "XLS":
                            strFileNameX = objProcessFile.ProcessAndSave(strFileName, the_Date);
                            FileFormatX = "XLS";
                            break;
                        case "PROCLOAD":
                            strFileNameX = objProcessFile.ProcessAndLoad(strFileName, the_Date);
                            FileFormatX = "XLS";
                            break;
                        case "PROCLOADX":
                            strFileNameX = objProcessFile.ProcessAndLoad(strFileName, the_Date);
                            FileFormatX = "XLSX";
                            break;
                        default:
                            strFileNameX = objProcessFile.ProcessAndSaveX(strFileName, the_Date);
                            FileFormatX = "XLSX";
                            break;

                    }


                    string strFile = string.Format(Utility.FileName, the_Date.ToString("yyyyMMdd"));
                    
                    objProcessFile.LoadToRGP(strFileNameX, strFile, FileFormatX);
                }
                string logfileRoot = log4net.GlobalContext.Properties["LogFileName"].ToString();
                string logFileAttach = logfileRoot + "Attach.log";
              

                System.IO.File.Copy(logfileRoot.ToString() + ".log", logFileAttach, true);
                Utility.AttachmentFileName = logFileAttach;
                if (Utility.LoadFile)
                    Utility.send(Utility.AUTOSYSJOBNAME + "- Success", Utility.AUTOSYSJOBNAME + " has Run Successfully. Data was loaded to database.");
                else
                    Utility.send(Utility.AUTOSYSJOBNAME + "- Success", Utility.AUTOSYSJOBNAME + " has Run Successfully. Data was not loaded to database. Data will load in DN4UK_LOAD_TIMESERIES_DATA.");
                return 0;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                //MessageBox.Show(ex.InnerException.ToString());
                log.Error(ex);
                string logfileRoot = log4net.GlobalContext.Properties["LogFileName"].ToString();
                string logFileAttach = logfileRoot + "Attach.log";

                System.IO.File.Copy(logfileRoot.ToString() + ".log", logFileAttach, true);
                Utility.AttachmentFileName = logFileAttach;
                Utility.send(Utility.AUTOSYSJOBNAME + "- Failure", ex.Message);
                
                return 1;
            }
        }
    }
}
