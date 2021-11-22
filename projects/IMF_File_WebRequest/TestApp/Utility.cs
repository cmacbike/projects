using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using log4net;
using StaticDust.Configuration;
using System.Net.Mail;
using System.Net;
using System.Reflection;


namespace IMF_WebRequestClient
{
   
    public static class Utility
    {
        
            static string m_strConnectionString;
            static string m_ConfigFile;
            private static System.Data.DataTable m_dt;
       
        
            public static  string getDateString(string p_DatabaseServer, string p_DatabaseName, bool p_ReformatDate)
            {



                           
                SqlConnection objConn = new SqlConnection(m_strConnectionString);
                objConn.Open();
                System.Data.DataTable dt = getData(objConn, "select * from rims..proc_dt");
                m_dt = dt;
                string strDate = "";
                if (p_ReformatDate)
                    strDate = ((DateTime)(dt.Rows[0][ProcDate])).ToString("yyyy-MM-dd") + " 00:00:00.000";
                else
                    strDate = ((DateTime)(dt.Rows[0][ProcDate])).ToString("yyyyMMdd");
                //strFileDate = DateTime.ParseExact(strDate,, System.Globalization.CultureInfo.InvariantCulture).ToString();

                return strDate;
            }

            public static System.Data.DataTable getData( SqlConnection p_DatabaseCon,string p_SQL)
            {
                SqlCommand cmd = new SqlCommand(p_SQL, p_DatabaseCon);
                cmd.CommandType = CommandType.Text;
                cmd.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                return dt;
            
            
            }
            public static SqlConnection GetDataConn()
            {
                SqlConnection DataConnection = new SqlConnection(m_strConnectionString);
                DataConnection.Open();
                return DataConnection;

            }
            public static SqlConnection GetDataConn(string p_Connectionstring)
            {
                SqlConnection DataConnection = new SqlConnection(p_Connectionstring);
                DataConnection.Open();
                return DataConnection;

            }
            

           
            public static bool use_Proc_date
            {
                get
                {
                    try
                    {
                        return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["use_Proc_date"]);
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            public static System.Data.DataTable ProcDateDataTable
            {
                get
                {
                    return m_dt;
                }
                set
                {
                    value = m_dt;
                }
            }
            public static string ProcDate
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["ProcDate"].ToString();
                    }
                    catch
                    {
                        return "nxt_proc_dt";
                    }
                }
            }
            public static string SQLStatementProcDate
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SQLStatementProcDate"].ToString();
                    }
                    catch
                    {
                        return "SELECT * FROM rims..proc_dt";
                    }
                }
            }
        public static DateTime SpecifyDate
        {
            get
            {
                try
                {
                    return Convert.ToDateTime(CustomConfigurationSettings.AppSettings(m_ConfigFile)["SpecifyDate"].ToString());
                }
                catch
                {
                    return DateTime.Now;
                }
            }
        }
        public static bool DateSpecified
        {
            get
            {
                try
                {
                    return System.Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["DateSpecified"].ToString());
                }
                catch
                {
                    return false;
                }
            }
        }
        public static bool FormatString
        {
            get
            {
                try
                {
                    return System.Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["FormatString"].ToString());
                }
                catch
                {
                    return false;
                }
            }
        }
        public static string ExcelSQLQuery
        {
            get
            {
                try
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["ExcelSQLQuery"].ToString();
                }
                catch
                {
                    return "select [Currency], " +
                "IIF(Currency = 'Chinese yuan', 'CNY',IIF(Currency = 'Euro', 'EUR',IIF(Currency = 'Japanese yen', 'JPY'," +
                "IIF(Currency = 'U.K. pound','GBP',IIF(Currency = 'U.K. Pound Sterling','GBP',IIF(Currency = 'U.S. dollar','XDR',IIF(Currency = 'Australian dollar','AUD'," +
                "IIF(Currency = 'Canadian dollar', 'CAD',IIF(Currency = 'Mexican peso','MXN',IIF(Currency = 'Norwegian krone','NOK',IIF(Currency = 'Swedish krona','SEK',))))))))))) , " +
                "[{1}] from {0} WHERE [Currency] IN ('Chinese yuan','Euro','Japanese yen','U.K. pound','U.K. Pound Sterling','U.S. dollar'" +
                ",'Mexican peso','Canadian dollar','Australian dollar','Swedish krona','Norwegian krone')";
                }
            }
        }
        public static string SQLQuery
        {
            get
            {
                try
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SQLQuery"].ToString();
                }
                catch
                {
                    return "";
                }
            }
        }
        public static string URL
        {
            get
            {
                try
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["URL"].ToString();
                }
                catch
                {
                    return "https://www.imf.org/external/np/fin/data/rms_mth.aspx?SelectDate={0}&reportType=CVSDR&tsvflag=Y";
                }
            }
        }
        public static string SheetName
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SheetName"].ToString();
                    }
                    catch
                    {
                        return "[IMF_data_{0}_WEBCLIENT$]";
                    }
                }
            }
        public static string SheetNameIDX
        {
            get
            {
                try
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SheetNameIDX"].ToString();
                }
                catch
                {
                    return "IMF_data_{0}_WEBCLIENT";
                }
            }
        }
        public static string XLSXConnectionString
        {
            get
            {
                try
                {
                    return (CustomConfigurationSettings.AppSettings(m_ConfigFile)["XLSXConnectionString"].ToString());
                }
                catch
                {
                    return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 XML;HDR=YES;IMEX=1\";";
                }
            }
        }
        public static string XLSConnectionString
        {
            get
            {
                try
                {
                    return (CustomConfigurationSettings.AppSettings(m_ConfigFile)["XLSConnectionString"].ToString());
                }
                catch
                {
                    return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";";
                }
            }
        }
        public static bool FormatColumns
        {
            get
            {
                try
                {
                    return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["FormatColumns"].ToString());
                }
                catch
                {
                    return false;
                }
            }
        }
        public static string[] ColumnsToFormat
        {
            get
            {
                try
                {
                    return new string[] { CustomConfigurationSettings.AppSettings(m_ConfigFile)["ColumnsToFormat"].ToString() };
                }
                catch
                {
                    return new string[]{ };
                }
            }
        }
        public static string[] Formats
        {
            get
            {
                try
                {
                    return new string[] { CustomConfigurationSettings.AppSettings(m_ConfigFile)["Formats"].ToString() };
                }
                catch
                {
                    return new string[] {"@"};
                }
            }
        }
        static public string strConnectionString
            {
                set
                {
                    m_strConnectionString = value;
                }

            
            }
            static public string Configfile
            {
                set
                {
                    m_ConfigFile = value;
                }

            }
            public static string LogFilePath
            {
                
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["LogFilePath"].ToString();
                    }
                    catch
                    {
                        return "\\\\us1.1corp.org\\dc0\\RCM\\prod\\apps\\rimsx\\IMF\\logs";
                    }
                }
               
            }
            public static string RIMSDBConnectionString
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["RIMSDBConnectionString"].ToString();
                }
            }
            public static string ConnectionString
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["ConnectionString"].ToString();
                }
            }
            public static string DatabaseServer
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["DatabaseServer"].ToString();
                }
            }
            public static string DatabaseName
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["DatabaseName"].ToString();
                }
            }
            public static string Source
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["Source"].ToString();
                }
            }
            public static string UseDate
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["UseDate"].ToString();
                    }
                    catch
                    {
                        return "NOTUSED";
                    }
                }
            }
            public static string InFile
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["InFile"].ToString();
                }
            }
            public static string FileName
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["FileName"].ToString();
                }
            }
            public static string SourceFilePath
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SourceFilePath"].ToString();
                }
            }
            public static string FilePath
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["FilePath"].ToString();
                }
            }
            public static string FilePathSaveTo
        {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["FilePathSaveTo"].ToString();
                }
            }
            public static string AUTOSYSJOBNAME
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["AUTOSYSJOBNAME"].ToString();

                    }
                    catch
                    {
                        return "DN4_IPMS_LOAD_IMF_FX_BOX";
                    }
                }
            }
            public static string FileFormat
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["FileFormat"].ToString();

                    }
                    catch
                    {
                        return "XLSX";
                    }
                }
                
            }
            public static bool LoadFile
            {
                get
                {
                    try
                    {
                        return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["LoadFile"].ToString());

                    }
                    catch
                    {
                        return true;
                    }
                }
            }
            public static bool WebRequest
            {
                get
                {
                    try
                    {
                        return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["WebRequest"].ToString());

                    }
                    catch
                    {
                        return true;
                    }
                }
            }
        public static bool Truncate
            {
                get
                {
                    return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["Truncate"]);
                }
            }
                
            
            public static string Table
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["Table"].ToString();
                }
            }
            public static string DataColumns
            {
                get
                {
                    try
                    {
                        return CustomConfigurationSettings.AppSettings(m_ConfigFile)["DataColumns"].ToString();
                    }
                    catch
                    {
                        return "*";
                    }
                }
            }
        public static string From
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["From"].ToString();
                }
            }
            public static string TO
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["TO"].ToString();
                }
            }
            public static string SMTPServer
            {
                get
                {

                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SMTPServer"].ToString();
                }
            }
            public static string SMTPUserName
            {
                get
                {

                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SMTPUserName"].ToString();
                }
            }

            public static string SMTPPassword
            {
                get
                {
                    return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SMTPPassword"].ToString();
                }
            }
            public static bool ReFormatDate
            {
                get
                {
                    return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["ReFormatDate"]);
                }
            }
            public static void send(string p_strSubject, string p_strBody)
            {


                SmtpClient objClient = new SmtpClient(SMTPServer);
                objClient.Credentials = new NetworkCredential(SMTPUserName, SMTPPassword);


                MailMessage objMessage = new MailMessage(From, TO, p_strSubject, p_strBody);

                if (AttachmentFileName == "")
                {
                    objClient.Send(objMessage);
                }
                else
                {
                    Attachment objAttachment = new System.Net.Mail.Attachment(AttachmentFileName);
                    objMessage.Attachments.Add(objAttachment);
                    objClient.Send(objMessage);
                }


            }
            public static string AttachmentFileName { get; set; }
       }
}
