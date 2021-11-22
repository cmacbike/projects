using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data;
using System.Xml.XPath;
using System.Text.RegularExpressions;
//using StarMine_Generic;

namespace StarmineXML
{
    class Program
    {
        static string m_strDatabaseServer;
        static string m_strDatabase;
        private static string m_TaskName = "";
        private static log4net.ILog log;
        private static string m_strLogFileLocation;
        private static string m_strAnalystFilePath;
        private static string m_strSecurityFilePath;
        private static string m_strRecommendationFilePath;
        private static string m_strDate = "";
        private static string m_strParmDate = "";
        private static string m_strConfigFileName = "";
        private static Dictionary<string, string> Fields;
        private static Dictionary<string, string> CompletionFields;
        private static string m_Results;
        private static bool m_USEXmlHTTP = true;
        private const string Success = "Pretty Good Privacy(tm) Version 6.5.2\r\n(c) 1999 Network Associates Inc.\r\nUses the BSafe(tm) Toolkit, which is copyright RSA Data Security, Inc.\r\nExport of this software may be restricted by the U.S. government.\r\n\r\n";
        // 0 = Recipient , 1 = pgp output file , 2 = passPhrase , 3 = file to encrypt
        //private static string m_CommandEncrypt = @"D:\LocalDeveloper\DataFeeds\external\common\bin\gnupg\gpg -r {0} --trust-model always -s -o {1} --passphrase {2} -e {3}";

        //0 = PGP Encryption Software location; 1 = Filename; 2 =  sendto; 3 = sendFrom; 4 = password;

        private static string m_CommandEncrypt = @"{0} -es ""{1}"" {2} -u ""{3}"" -z {4}";
        private static string m_strPrivateKeyRing = @"\\us1.1corp.org\DC0\RCM\prod\apps\rimsx\StarMine_dev\Keyrings\RIMSSchedsecring.skr";
        private static string m_strPublicKeyRing = @"\\us1.1corp.org\DC0\RCM\prod\apps\rimsx\StarMine_dev\Keyrings\RIMSSchedpubring.pkr";
        //@"D:\LocalDeveloper\DataFeeds\external\common\bin\gnupg\gpg --passphrase dresdnerRCMSecret --import \\us1.1corp.org\DC0\RCM\prod\apps\rimsx\StarMine_dev\Keyrings\RIMSSchedsecring.skr";
        //private static string m_recipient = "StarMine (Default key for web signing and encryption) <webmaster@starmine.com>";
        private static string m_recipient = "RIMS Scheduler";
        private static string m_CommandImport = @"D:\LocalDeveloper\DataFeeds\external\common\bin\gnupg\gpg --passphrase {0} --import {1}";

        

        static int Main(string[] args)
        {
            try
            {
                //get the data
                //Console.WriteLine(System.Convert.ToInt32(DateTime.Now.AddDays(-3).DayOfWeek).ToString() + DateTime.Now.AddDays(-2).DayOfWeek);
                string ConfigFilePath = "";

                //Connect to Database
                //Query to determine reddiness
                int i;
                for (i = 0; i < args.Length; i++)
                {
                    switch (args[i])
                    {

                        case "-C":
                            ConfigFilePath = args[i + 1];

                            //TestFile.WriteLine(m_strStoredProcedure);
                            break;
                        case "-T":
                            m_TaskName = args[i + 1];
                            //TestFile.WriteLine(m_LoadType);
                            break;

                        case "-G":
                            m_strConfigFileName = args[i + 1];
                            //TestFile.WriteLine(m_LoadType);
                            break;
                      
                        default:
                            break;


                    }
                }

                ConfigFilePath = ConfigFilePath.Substring(ConfigFilePath.Length - 1, 1) == "\\" ? ConfigFilePath : ConfigFilePath + "\\";
                
                Utility.ConfigFile = ConfigFilePath + m_strConfigFileName;
               
                m_strDatabaseServer = Utility.DatabaseServer;
                m_strDatabase = Utility.DatabaseName;
                m_strAnalystFilePath = Utility.AnalystFilename;
                m_strSecurityFilePath = Utility.SecurityFilename;
                m_strRecommendationFilePath = Utility.RecommendationFilename;
                m_strLogFileLocation = Utility.LogFilePath;

                Utility.strConnectionString = String.Format("Data Source={0};initial catalog={1};Integrated Security=SSPI;", m_strDatabaseServer, m_strDatabase);
                m_strLogFileLocation = Utility.LogFilePath;
                m_strLogFileLocation = m_strLogFileLocation.Substring(m_strLogFileLocation.Length - 1, 1) == "\\" ? m_strLogFileLocation : m_strLogFileLocation + "\\";
                log4net.GlobalContext.Properties["LogFileName"] = m_strLogFileLocation + m_TaskName + "log";
                //log4net.GlobalContext.Properties["LogFileName"] = m_strLogFileLocation + "FTP";
                log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(ConfigFilePath + m_strConfigFileName));
                log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                DayOfWeek IsMonday = new DayOfWeek();


                IsMonday = DayOfWeek.Monday;
                //DayOfWeek IsSaturday = new DayOfWeek();
                //IsSaturday = DayOfWeek.Sunday;
                if (DateTime.Now.Hour < 13)
                {
                    if (DateTime.Now.DayOfWeek.Equals(IsMonday))
                    {
                        m_strDate = DateTime.Now.AddDays(-3).ToString("yyyy-MM-dd");
                        m_strParmDate = DateTime.Now.AddDays(-3).ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        m_strDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                        m_strParmDate = DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd");
                        //m_strDate = DateTime.Now.ToString("yyyy-MM-dd");
                        //m_strParmDate = DateTime.Now.ToString("yyyy/MM/dd");
                    }
                }
                else
                {
                    m_strDate = DateTime.Now.ToString("yyyy-MM-dd");
                    m_strParmDate = DateTime.Now.ToString("yyyy/MM/dd");
                }

                Fields = new Dictionary<string, string>();
                Fields.Add("MAX_FILE_SIZE", "10000000");
                Fields.Add("supplier", Utility.supplierID);
                Fields.Add("action", "upload");
                Fields.Add("file_packaging", "none");
                Fields.Add("sec_method", "pgp-encrypted-signed");
                Fields.Add("file_format", "xml");

                CompletionFields = new Dictionary<string, string>();
                CompletionFields.Add("supplier", Utility.supplierID);
                CompletionFields.Add("action", "upload-complete");
                CompletionFields.Add("date", m_strDate);

                if(Utility.UploadOnly)
                {
                    m_strAnalystFilePath = m_strAnalystFilePath.Replace("YYYY-MM-DD", m_strDate);
                    string strXMLResult = HTTPUpLoad(m_strAnalystFilePath);
                    m_Results = "Analyst Document\r\n" + strXMLResult;
                    if (!GetStatusXML(strXMLResult).Equals("0"))
                    {
                        string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strAnalystFilePath, strXMLResult);
                        Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                    }

                    m_strSecurityFilePath = m_strSecurityFilePath.Replace("YYYY-MM-DD", m_strDate);
                    strXMLResult = HTTPUpLoad(m_strSecurityFilePath);
                    m_Results = m_Results + "\r\n\r\nSecurity Document\r\n" + strXMLResult;
                    if (!GetStatusXML(strXMLResult).Equals("0"))
                    {
                        string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strSecurityFilePath, strXMLResult);
                        Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                    }

                    m_strRecommendationFilePath = m_strRecommendationFilePath.Replace("YYYY-MM-DD", m_strDate);
                    strXMLResult = HTTPUpLoad(m_strRecommendationFilePath);
                    m_Results = m_Results + "\r\n\r\nRecommendation Document\r\n" + strXMLResult;
                    if (!GetStatusXML(strXMLResult).Equals("0"))
                    {
                        string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strRecommendationFilePath, strXMLResult);
                        Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                    }
                }
                else
                { 
                    WriteAnalystFile();
                    writeSecurityFile();
                    writeRecommendationFile();
                }
                if (!Utility.WriteFilesOnly)
                {
                    UploadToStarmine DoUpload = new UploadToStarmine(log);

                    string strResultCompletion = DoUpload.UploadFileHTTP(Utility.URI, "", CompletionFields);

                    m_Results = m_Results + "\r\n\r\nCompletion\r\n" + strResultCompletion;
                    Utility.send(Utility.From, Utility.To, Utility.SMTPServer, string.Format("StarMine Upload Result Production - {0}", Utility.supplierID), m_Results, "", "");
                
                    SqlConnection RIMSConn = Utility.GetDataConn();
                    string strSQL = string.Format("exec {0} '{1}'",Utility.STPDeleteRecommendations,m_strParmDate) ;

                    Utility.executeSQL(RIMSConn, strSQL);
                    

                }
                return 0;
                //HTTPUpLoad();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                log.Info(ex.Message);
                return 1;
            }
           
           

        }
        private static void WriteAnalystFile()
        {
            
            m_strAnalystFilePath = m_strAnalystFilePath.Replace("YYYY-MM-DD", m_strDate);
            log.Info(string.Format("Writing {0}",m_strAnalystFilePath));
            System.IO.StreamWriter XMLStream = new System.IO.StreamWriter(m_strAnalystFilePath);
            XmlWriterSettings Settings = new XmlWriterSettings();
            Settings.NewLineOnAttributes = true;
            Settings.NewLineHandling = NewLineHandling.Replace;
            Settings.Encoding = Encoding.Unicode;
            using (XmlWriter objWriteAnalystFile = XmlWriter.Create(XMLStream,Settings))
            {
               
                           
                SqlConnection RIMSConn = Utility.GetDataConn();
            

                DataTable dt = Utility.getData(RIMSConn, string.Format ("exec rims..{0}",Utility.GetAnalystStoredProcedure));
                objWriteAnalystFile.WriteProcessingInstruction("xml", @"version=""1.0"" encoding=""ISO-8859-1""");

                objWriteAnalystFile.WriteStartElement("AnalystSet","http://www.starmine.com/2002/6/SMXML");
                objWriteAnalystFile.WriteAttributeString("loadType", "incremental");
                objWriteAnalystFile.WriteAttributeString("supplierID", Utility.supplierID);
                objWriteAnalystFile.WriteAttributeString("date", m_strDate);
                objWriteAnalystFile.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");

                foreach (DataRow dr in dt.Rows)
                {
                    objWriteAnalystFile.WriteStartElement("Analyst");
                    objWriteAnalystFile.WriteAttributeString("action", "update");
                    objWriteAnalystFile.WriteAttributeString("analystType", "single");
                    
                    Regex r = new Regex("(_L)$");
                   
                    objWriteAnalystFile.WriteStartElement("AnalystID");
                    objWriteAnalystFile.WriteString(dr["Office"].ToString().Equals("SF") && !Utility.supplierID.Equals("rcmgrassroots") && !r.IsMatch(dr["anlst_id"].ToString()) ? dr["anlst_id"].ToString() + "_" + "L" : dr["anlst_id"].ToString());
                    objWriteAnalystFile.WriteEndElement();

                    objWriteAnalystFile.WriteStartElement("DisplayName");
                    objWriteAnalystFile.WriteString(dr["Office"].ToString().Equals("SF") ? dr["anlst_fname"].ToString(): dr["anlst_fname"].ToString().Replace(", CFA", ""));
                    objWriteAnalystFile.WriteEndElement();

                    objWriteAnalystFile.WriteEndElement();
                   
                }

                objWriteAnalystFile.WriteEndElement();
                objWriteAnalystFile.Close();

                XMLStream.Close();
                //log.Info(string.Format("Encrypting {0}", m_strAnalystFilePath));
                //EncryptSign encrypter = new EncryptSign();
                //encrypter.Encrypt(m_strAnalystFilePath, Utility.PublicKey, Utility.PrivateKey, Utility.secretPassPhrase);
                //log.Info(string.Format("{0} Encrypted", m_strAnalystFilePath));
            }

            if (!Utility.WriteFilesOnly)
            {
                Encrypt(m_strAnalystFilePath);
                string strXMLResult = HTTPUpLoad(m_strAnalystFilePath);
                m_Results = "Analyst Document\r\n" + strXMLResult;
                if (!GetStatusXML(strXMLResult).Equals("0"))
                {
                    string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strAnalystFilePath, strXMLResult);
                    Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                }
            }
         }
        private static void writeSecurityFile()
        {
            
            m_strSecurityFilePath = m_strSecurityFilePath.Replace("YYYY-MM-DD", m_strDate);
            log.Info(string.Format("Writing {0}", m_strSecurityFilePath));
            XmlWriterSettings Settings = new XmlWriterSettings();
            Settings.NewLineOnAttributes = true;
            Settings.NewLineHandling = NewLineHandling.Replace;
            Settings.Encoding = Encoding.Unicode;
            System.IO.StreamWriter XMLStream = new System.IO.StreamWriter(m_strSecurityFilePath);
            using (XmlWriter objWriteSecFile = XmlWriter.Create(XMLStream, Settings))
            {
                 


                SqlConnection RIMSConn = Utility.GetDataConn();


                DataTable dt = Utility.getData(RIMSConn, string.Format("exec rims..{0}",Utility.GetSecurityStoredProcedure));
                objWriteSecFile.WriteProcessingInstruction("xml", @"version=""1.0"" encoding=""ISO-8859-1""");

                objWriteSecFile.WriteStartElement("SecuritySet", "http://www.starmine.com/2002/6/SMXML");
                objWriteSecFile.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
                objWriteSecFile.WriteAttributeString("date", m_strDate);
                objWriteSecFile.WriteAttributeString("supplierID", Utility.supplierID);
                objWriteSecFile.WriteAttributeString("loadType", "full");
                

                foreach (DataRow dr in dt.Rows)
                {
                    objWriteSecFile.WriteStartElement("Security");
                    objWriteSecFile.WriteAttributeString("action", "update");
                    objWriteSecFile.WriteAttributeString("date", m_strDate);

                    objWriteSecFile.WriteStartElement("SupplierSecurityID");
                    objWriteSecFile.WriteString(dr["rims_id"].ToString());
                    objWriteSecFile.WriteEndElement();

                    objWriteSecFile.WriteStartElement("Company");
                    objWriteSecFile.WriteStartElement("DisplayName");
                    objWriteSecFile.WriteString(dr["name"].ToString());
                    objWriteSecFile.WriteEndElement();
                    objWriteSecFile.WriteStartElement("Country");
                    objWriteSecFile.WriteAttributeString("source", "iso");
                    objWriteSecFile.WriteString(dr["iso_country_id"].ToString());
                    objWriteSecFile.WriteEndElement();
                    objWriteSecFile.WriteEndElement();

                    objWriteSecFile.WriteStartElement("Country");
                    objWriteSecFile.WriteAttributeString("source", "iso");
                    objWriteSecFile.WriteString(dr["exch_country_id"].ToString());
                    objWriteSecFile.WriteEndElement();

                    objWriteSecFile.WriteStartElement("SecurityID");
                    objWriteSecFile.WriteAttributeString("idType", dr["iso_country_id"].ToString().Equals("US")? "cusip":"sedol");
                    objWriteSecFile.WriteString(dr["exch_country_id"].ToString().Equals("US") && !dr["cusip"].Equals(null)? dr["cusip"].ToString() : dr["sedol"].ToString());
                    objWriteSecFile.WriteEndElement();

                    objWriteSecFile.WriteStartElement("SecurityID");
                    objWriteSecFile.WriteAttributeString("idType", "listed_ticker");
                    objWriteSecFile.WriteString(dr["TKR"].ToString());
                    objWriteSecFile.WriteEndElement();

                    

                    objWriteSecFile.WriteEndElement();

                }

                objWriteSecFile.WriteEndElement();
                objWriteSecFile.Close();
                XMLStream.Close();
                //log.Info(string.Format("Encrypting {0}", m_strSecurityFilePath));
                //EncryptSign encrypter = new EncryptSign();
                //encrypter.Encrypt(m_strSecurityFilePath, Utility.PublicKey, Utility.PrivateKey, Utility.secretPassPhrase);
                //log.Info(string.Format("{0} Encrypted", m_strSecurityFilePath));
                if (!Utility.WriteFilesOnly)
                {
                    Encrypt(m_strSecurityFilePath);
                    string strXMLResult = HTTPUpLoad(m_strSecurityFilePath);
                    m_Results = m_Results + "\r\n\r\nSecurity Document\r\n" + strXMLResult;
                    if (!GetStatusXML(strXMLResult).Equals("0"))
                    {
                        string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strSecurityFilePath, strXMLResult);
                        Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                    }
                }
            }
           
        }
        private static void writeRecommendationFile()
        {
            //DateTime dtDBDateTime = Utility.GetLocalDateTime();
            m_strRecommendationFilePath = m_strRecommendationFilePath.Replace("YYYY-MM-DD", m_strDate);
            log.Info(string.Format("Writing {0}", m_strRecommendationFilePath));
            string strRecommendfilename = m_strRecommendationFilePath.Substring(m_strRecommendationFilePath.LastIndexOf("\\"),m_strRecommendationFilePath.Length -m_strRecommendationFilePath.LastIndexOf("\\"));
            System.IO.StreamWriter RecommendLogFile = new System.IO.StreamWriter(m_strLogFileLocation + strRecommendfilename + ".log");

            System.IO.StreamWriter XMLStream = new System.IO.StreamWriter(m_strRecommendationFilePath);
            XmlWriterSettings Settings = new XmlWriterSettings();
            Settings.NewLineOnAttributes = true;
            Settings.NewLineHandling = NewLineHandling.Replace;
            Settings.Encoding  = Encoding.Unicode;
            using (XmlWriter objWriteRecommendationFile = XmlWriter.Create(XMLStream, Settings))
            {



                SqlConnection RIMSConn = Utility.GetDataConn();
                DataTable dt = Utility.getData(RIMSConn, string.Format("exec rims..{0} '{1}', 'drcm'",Utility.GetRecommendationStoredProcedure,m_strParmDate));

                objWriteRecommendationFile.WriteProcessingInstruction("xml", @"version=""1.0"" encoding=""ISO-8859-1""");

                objWriteRecommendationFile.WriteStartElement("RecommendationSet", "http://www.starmine.com/2002/6/SMXML");
                objWriteRecommendationFile.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
                objWriteRecommendationFile.WriteAttributeString("date", m_strDate);
                objWriteRecommendationFile.WriteAttributeString("supplierID", Utility.supplierID);
                objWriteRecommendationFile.WriteAttributeString("loadType", "incremental");
                
                
                int recordcount = 1;
                //int RowIndex = 0;
                //int rowCount = dt.Rows.Count;
                System.Text.RegularExpressions.Regex rx = new System.Text.RegularExpressions.Regex("\\b(Z|ZZ|ZZZ)\\b");
                foreach (DataRow dr in dt.Rows)
                {

                    




                    RecommendLogFile.WriteLine(string.Format("Writing Record Number {0}",recordcount.ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tstarMinevote = {0}", dr["StarMineVoteValue"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tAnlst_vote = {0}", dr["anlst_vote"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tStarMine_changed_date = {0}", dr["StarMine_changed_date"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tchanged_date = {0}", dr["changed_date"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tchanged_time = {0}", dr["changed_time"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tMarketCap = {0}", dr["cur_MarketCap"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tanlst_id = {0}", dr["anlst_id"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tAnalystStatus = {0}", dr["Status"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tRims_id_value = {0}", dr["rims_id"].ToString()));
                    RecommendLogFile.WriteLine(string.Format("\tbroad_ind_id = {0}", dr["broad_ind_id"].ToString()));


                    if (dr["Status"].ToString().Equals("A"))
                    {

                        DateTime Starmine_Changed_Date = System.Convert.ToDateTime(dr["StarMine_changed_date"].ToString());

                        if (Starmine_Changed_Date.DayOfWeek.Equals(6))
                        {
                            Starmine_Changed_Date = Starmine_Changed_Date.AddDays(2);
                        }
                        else if (Starmine_Changed_Date.DayOfWeek.Equals(0)) 
                        {
                            Starmine_Changed_Date = Starmine_Changed_Date.AddDays(1);
                        }

                         
                        
                        objWriteRecommendationFile.WriteStartElement("Recommendation");
                        objWriteRecommendationFile.WriteAttributeString("action", "update");
                        objWriteRecommendationFile.WriteAttributeString("date", Starmine_Changed_Date.ToString("yyyy-MM-dd"));
                        objWriteRecommendationFile.WriteAttributeString("timestamp", ((DateTime)dr["changed_time"]).ToString(string.Format("yyyy-MM-ddTHH:mm:ss+{0}", Utility.getTimeZone())));

                        if (rx.IsMatch(dr["anlst_id"].ToString().Trim().ToUpper()))
                        {
                            objWriteRecommendationFile.WriteAttributeString("event", "stop");
                            objWriteRecommendationFile.WriteElementString("SecurityID", dr["rims_id"].ToString());
                            objWriteRecommendationFile.WriteEndElement();
                        }
                        else
                        {

                            objWriteRecommendationFile.WriteAttributeString("event", "initial");
                            if (dr["anlst_vote"].Equals(null) || dr["anlst_vote"].ToString().Equals("") || dr["anlst_vote"].ToString().Equals("0"))
                            {
                                objWriteRecommendationFile.WriteElementString("Value", "6");
                                objWriteRecommendationFile.WriteElementString("Text", "NA");
                                objWriteRecommendationFile.WriteElementString("SupplierValue", "NA");
                            }
                            else
                            {
                                objWriteRecommendationFile.WriteElementString("Value", dr["StarMineVoteValue"].ToString());
                                objWriteRecommendationFile.WriteElementString("Text", string.Format("{0:0.0}",Convert.ToInt32(dr["anlst_vote"])));
                                objWriteRecommendationFile.WriteElementString("SupplierValue", string.Format("{0:0.0}", Convert.ToInt32(dr["anlst_vote"])));

                            }

                            objWriteRecommendationFile.WriteElementString("SecurityID", dr["rims_id"].ToString());
                            objWriteRecommendationFile.WriteElementString("AnalystID", dr["anlst_id"].ToString());
                            objWriteRecommendationFile.WriteEndElement();
                        }
                    }
                }

                objWriteRecommendationFile.WriteEndElement();
                objWriteRecommendationFile.Close();
                XMLStream.Close(); 
                //System.Diagnostics.Process.Start("");
                //C:\>D:\LocalDeveloper\DataFeeds\external\common\bin\gnupg\gpg -r "StarMine (Default key for web signing and encryption) <webmaster@starmine.com>" --trust-model always -s -o \\us1.1corp.org\DC0\RCM\prod\apps\rimsx\StarMine_dev\rcmgrassroots-Recommendation-Incremental-2016-06-16.xml --passphrase dresdnerRCMSecret -e \\us1.1corp.org\DC0\RCM\prod\apps\rimsx\StarMine_dev\rcmgrassroots-Recommendation-Incremental-2016-06-16.xml.pgp

          
                

                //log.Info(string.Format("Encrypting {0}", m_strRecommendationFilePath));
                //EncryptSign encrypter = new EncryptSign();
                //encrypter.Encrypt(m_strRecommendationFilePath, Utility.PublicKey, Utility.PrivateKey, Utility.secretPassPhrase);
                //log.Info(string.Format("{0} Encrypted", m_strRecommendationFilePath));
                if (!Utility.WriteFilesOnly)
                {
                    Encrypt(m_strRecommendationFilePath);
                    string strXMLResult = HTTPUpLoad(m_strRecommendationFilePath);
                    m_Results = m_Results + "\r\n\r\nRecommendation Document\r\n" + strXMLResult;
                    if (!GetStatusXML(strXMLResult).Equals("0"))
                    {
                        string Body = string.Format("Errors while sending StareMine file {0} \r\n\r\n {1} ", m_strRecommendationFilePath, strXMLResult);
                        Utility.send(Utility.From, Utility.To, Utility.SMTPServer, "StarMine Upload Failure", Body, "", "");
                    }
                }   
              
            }
            
        }
        private static void Encrypt(string p_FileToEncrypt)
        {

            string EcryptedFileName = p_FileToEncrypt + ".pgp";
            if (System.IO.File.Exists(EcryptedFileName))
                System.IO.File.Delete(EcryptedFileName);

            System.Diagnostics.Process ProcEncrypt = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo ProcEncryptInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe");
            ProcEncryptInfo.RedirectStandardInput = true;
            ProcEncryptInfo.UseShellExecute = false;
            ProcEncryptInfo.RedirectStandardError = true;

            ProcEncrypt.StartInfo = ProcEncryptInfo;
            ProcEncrypt.Start();


            using (System.IO.StreamWriter ProcEncryptStream = ProcEncrypt.StandardInput)
            {
                //ProcEncryptStream.WriteLine(string.Format(m_CommandImport,Utility.secretPassPhrase,Utility.PublicKey));
                //ProcEncryptStream.WriteLine(string.Format(m_CommandImport,Utility.secretPassPhrase,Utility.PrivateKey));
                // 0 = Recipient , 1 = pgp output file , 2 = passPhrase , 3 = file to encrypt
                //ProcEncryptStream.WriteLine(string.Format(m_CommandEncrypt, "\"" + m_recipient + "\"", p_FileToEncrypt + ".pgp", Utility.secretPassPhrase, p_FileToEncrypt));
                ProcEncryptStream.WriteLine(string.Format(m_CommandEncrypt,Utility.PGPLocation, p_FileToEncrypt, Utility.SendTo, Utility.SendFrom, Utility.secretPassPhrase));
                log.Info(string.Format(m_CommandEncrypt, Utility.PGPLocation, p_FileToEncrypt, Utility.SendTo, Utility.SendFrom, Utility.secretPassPhrase));
                ProcEncryptStream.Close();
            }


            ProcEncrypt.WaitForExit();
            string Error = ProcEncrypt.StandardError.ReadToEnd();
            if (Error == Success)
            {
                ProcEncrypt.Close();
                log.Info(Success);
            }
            else
            {
                ProcEncrypt.Close();
                log.Info(Error);
                //throw new Exception(Error);
            }
        }
        
        private static string HTTPUpLoad(string p_strFile)
        {

            //RIMSBatchTask.clsParameters Parameters = new RIMSBatchTask.clsParameters();
            //Parameters.ArbParm1 = Utility.UploadConfig;


            string Uploadresult;
            //clsStarMineUpload StarMineUload = new clsStarMineUpload();
            //StarMineUload.set_ParameterBlock(Parameters);
            //StarMineUload.ProcessData();
            //StarMineUload = null;
            //UploadToStarmine DoUpload = new UploadToStarmine(log);
            if (Utility.UseXMLHttp)
            {
                UploadToStarmine DoUpload = new UploadToStarmine(log);
                Uploadresult = DoUpload.UploadFileHTTP(Utility.URI, p_strFile + ".pgp", Fields);
            }
            else
            {
                UploadStarmineFiles DoUpload = new UploadStarmineFiles(log);
                Uploadresult = DoUpload.UploadFileHTTP(Utility.URI, p_strFile + ".pgp", Fields);
            }

            System.Threading.Thread.Sleep(Utility.SleepInterval);
            return Uploadresult;
            
        }
        private static string GetStatusXML(string p_XML)
        {

            XmlDocument XDoc = new XmlDocument();
            XDoc.LoadXml(p_XML);
            XPathNavigator XNav = XDoc.CreateNavigator();
            //XPathNodeIterator XNode;
            
            string Result = XNav.SelectSingleNode("/upload/@result").Value;
            return Result;

        }
    }

} 
