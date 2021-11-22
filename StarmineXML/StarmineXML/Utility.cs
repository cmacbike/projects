using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StaticDust.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;
using System.Reflection;

namespace StarmineXML
{
    public static class Utility
    {
        static string m_strConnectionString;
        static  string m_strMessage;
        static string m_ConfigFile;
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
        public  static SqlConnection GetDataConn()
        {
            SqlConnection DataConnection = new SqlConnection(m_strConnectionString);
            DataConnection.Open();
            return DataConnection;

        }
        public static DateTime GetLocalDateTime()
        {
            SqlConnection conn = GetDataConn();
            DataTable dt = getData(conn, "select getdate()");
            DateTime dtDBDate = System.Convert.ToDateTime(dt.Rows[0][0].ToString());
            return dtDBDate;


        }
          public static void insertTask(string p_TaskName)
         {

             string strSQL;
             SqlConnection connRIMS = GetDataConn();
             strSQL = "DELETE rims..TaskDailySchedule where TaskName = '" + p_TaskName + "'";
             SqlCommand cmd = new SqlCommand(strSQL, connRIMS);
             cmd.CommandType = CommandType.Text;
             cmd.CommandTimeout = 0;
             cmd.ExecuteNonQuery();

             string strInsertSQL = @"INSERT INTO rims..TaskDailySchedule(TaskName,SchedStartTime,ActualStarttime,Status, OLEExecName, BaseFileName, SuffixFormat,
            FileSuffix, FileLocation, LocalFileName, ProxyName, Processor1, Processor2, Processor3,Dependency1,
            Dependency2, Dependency3, ArbParm1, ArbParm2, ArbParm3,DatabaseServer, DatabaseName, StagingTable, 
            StoredProcedure, RetryDelayMinutes, TaskMasterID) 
            SELECT TaskName, convert(varchar, getdate(),100),convert(varchar, getdate(),100), 'AUTOSYSRunning', OLEExecName, BaseFileName, SuffixFormat,
            FileSuffix, FileLocation, LocalFileName, ProxyName, Processor1, Processor2, Processor3,
            Dependency1, Dependency2, Dependency3, ArbParm1, ArbParm2, ArbParm3,DatabaseServer , DatabaseName, 
            StagingTable, StoredProcedure, RetryDelayMinutes, TaskMasterID 
            From rims..TaskMaster WHERE TaskName = '" + p_TaskName + "'";
             SqlCommand inscmd = new SqlCommand(strInsertSQL, connRIMS);
             inscmd.CommandType = CommandType.Text;
             inscmd.CommandTimeout = 0;
             inscmd.ExecuteNonQuery();
             connRIMS.Close();
             connRIMS.Dispose();


         }

         public static void updateTask(string p_DataProcessed, string p_Status, string p_Phase, string p_Taskname,Boolean p_ErrorStatus)
         {
             string strSQL = "";

             if (p_Status != "Complete" || !p_ErrorStatus)
                 strSQL = @"update rims..TaskDailySchedule  
                set DataProcessed = '" + p_DataProcessed + @"',  
                Status = '" + p_Status + @"',  
                Phase = '" + p_Phase + @"' 
                WHERE TaskName = '" + p_Taskname + "'";
             else 
                strSQL = @"update rims..TaskDailySchedule  
                set DataProcessed = '" + p_DataProcessed + @"',  
                Status = '" + p_Status + @"',  
                Phase = '" + p_Phase + @"',
                EndTime = convert(varchar, getdate(),100) 
                WHERE TaskName = '" + p_Taskname + "'";

             SqlConnection connRIMS = GetDataConn();

             SqlCommand cmd = new SqlCommand(strSQL, connRIMS);
             cmd.CommandType = CommandType.Text;
             cmd.ExecuteNonQuery();
             connRIMS.Close();
             connRIMS.Dispose();
         }

         public static void executeSQL(SqlConnection p_Connection,string p_strSQL)
         {
             

             SqlConnection connRIMS = GetDataConn();

             SqlCommand cmd = new SqlCommand(p_strSQL, p_Connection);
             cmd.CommandType = CommandType.Text;
             cmd.ExecuteNonQuery();
             connRIMS.Close();
             connRIMS.Dispose();
         }
         public static void send(string p_strFrom, string p_strTO, string p_strSMTPServer, string p_strSubject, string p_strBody, string p_UserName, string p_Password)
         {
            

                 SmtpClient objClient = new SmtpClient(p_strSMTPServer);
                 objClient.Credentials = new NetworkCredential(p_UserName, p_Password);


                 MailMessage objMessage = new MailMessage(p_strFrom, p_strTO, p_strSubject, p_strBody);

                 
                     objClient.Send(objMessage);
                 


         }
         public static string getTimeZone()
         {
             TimeZoneInfo tzInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
             DateTime dtLocalTime = TimeZoneInfo.ConvertTime(DateTime.Now, tzInfo);
             return (-tzInfo.GetUtcOffset(dtLocalTime)).ToString().Substring(0, 5);

         }
         static public string strConnectionString
         {
             set
             {
                 m_strConnectionString = value;
             }

         }
         public static string ConfigFile
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
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["LogFilePath"].ToString();
             }
         }

         public static string AnalystFilename
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["AnalystFilename"].ToString();
             }
         }
         public static string SendTo
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SendTo"].ToString();
             }
         }

         public static string SendFrom
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SendFrom"].ToString();
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
         public static string  SecurityFilename
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SecurityFileName"].ToString();
             }
         }
         public static string RecommendationFilename
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["RecommendationFilename"].ToString();
             }
         }

         public static string GetAnalystStoredProcedure
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["GetAnalystStoredProcedure"].ToString();
             }
         }
         public static string GetSecurityStoredProcedure
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["GetSecurityStoredProcedure"].ToString();
             }
         }
         public static string STPDeleteRecommendations
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["STPDeleteRecommendations"].ToString();
             }
         }
         public static string GetRecommendationStoredProcedure
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["GetRecommendationStoredProcedure"].ToString();
             }
         }
         public static string PublicKey
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["PublicKey"].ToString();
             }
         }
         public static string PrivateKey
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["PrivateKey"].ToString();
             }
         }
         public static string secretPassPhrase
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["secretPassPhrase"].ToString();
             }
         }
         public static string URI
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["URI"].ToString();
             }
         }
         public static string UserName
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["UserName"].ToString();
             }
         }
         public static string UserPass
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["UserPass"].ToString();
             }
         }
         public static string From
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["From"].ToString();
             }
         }
         public static string To
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["To"].ToString();
             }
         }
         public static string SMTPServer
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["SMTPServer"].ToString();
             }
         }
         public static string supplierID
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["supplierID"].ToString();
             }
         }
         public static string UploadConfig
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["UploadConfig"].ToString();
             }
         } 
         public static string PGPLocation
         {
             get
             {
                 return CustomConfigurationSettings.AppSettings(m_ConfigFile)["PGPLocation"].ToString();
             }
         }
         public static bool UseXMLHttp
         {
             get
             {
                 return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["UseXMLHttp"]);
             }
         }
         public static bool WriteFilesOnly
         {
             get
             {
                 return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["WriteFilesOnly"]);
             }
         }
        public static Int32  SleepInterval
        {
            get
            {
                return Convert.ToInt32(CustomConfigurationSettings.AppSettings(m_ConfigFile)["SleepInterval"]);
            }
        }
        public static bool UploadOnly
        {

            get
            {
                try
                {
                    return Convert.ToBoolean(CustomConfigurationSettings.AppSettings(m_ConfigFile)["UploadOnly"]);
                }
                catch
                {
                    return false;
                }

            }
        }
        public static string AttachmentFileName { get; set; }
         
    }
}

    