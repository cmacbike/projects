using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
using log4net;
using MSXML2;

namespace StarmineXML
{
    class UploadToStarmine
    {
        private ILog m_log;
        private string boundary;
        public UploadToStarmine(ILog p_log)
        {
            m_log = p_log;
        }
        public string UploadFileHTTP(string p_strURL,string p_strFile,Dictionary<string,string> p_Fields)
        {
            string strFormData = "";
            string strData = "";
            boundary = "---------------------------0123456789012";

            strData = "--" + boundary + "\r\n";
            foreach (KeyValuePair<string, string> Pair in p_Fields)
            {

                strData = strData + string.Format("Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}\r\n", Pair.Key, Pair.Value);
                strData = strData + "--" + boundary + "\r\n";
            }

            byte[] binFormData;

            if (!p_strFile.Trim().Equals(""))
            {
                System.IO.FileInfo objFileInfo = new System.IO.FileInfo(p_strFile);
                byte[] FileContents = new byte[objFileInfo.Length];
                System.IO.FileStream theFileStream = objFileInfo.Open(System.IO.FileMode.Open);
                System.IO.BinaryReader bReader = new System.IO.BinaryReader(theFileStream);
                FileContents = bReader.ReadBytes(System.Convert.ToInt32(theFileStream.Length));
                //strFormData = System.Convert.ToBase64String(FileContents);

                string trailer = "\r\n--" + boundary + "--\r\n";

                strData = strData + string.Format("Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: application/upload\r\n\r\n", "userfile", p_strFile);
                //var some64Data = Convert.ToBase64String(Encoding.UTF8.GetBytes(strData)) + strFormData + Convert.ToBase64String(Encoding.UTF8.GetBytes(trailer));
                //strData = strData + strFormData + "\r\n--" + boundary + "--\r\n";
                var some64Databytes = new byte[Encoding.UTF8.GetBytes(strData).Length + FileContents.Length + Encoding.UTF8.GetBytes(trailer).Length];

                Encoding.UTF8.GetBytes(strData).CopyTo(some64Databytes, 0);
                FileContents.CopyTo(some64Databytes, Encoding.UTF8.GetBytes(strData).Length);
                Encoding.UTF8.GetBytes(trailer).CopyTo(some64Databytes, FileContents.Length + Encoding.UTF8.GetBytes(strData).Length);
                var someBase64StringData = Convert.ToBase64String(some64Databytes);
                //strData = strData + "\r\n--" + boundary + "--\r\n";
                binFormData = System.Convert.FromBase64String(someBase64StringData);
            }
            else
            {
                var some64DataString = Convert.ToBase64String(Encoding.UTF8.GetBytes(strData));
                binFormData = System.Convert.FromBase64String(some64DataString);
            }
            m_log.Info(strData);
            
            //var someBytes = Encoding.UTF8.GetBytes(strData);
            //var some64Data = Convert.ToBase64String(someBytes);

            ServerXMLHTTP60 XmlHTTP = new ServerXMLHTTP60();
            XmlHTTP.setOption(SERVERXMLHTTP_OPTION.SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, 13056);

            XmlHTTP.setTimeouts(0, 180000, 180000, 180000);
            XmlHTTP.open("POST", p_strURL, false, Utility.UserName, Utility.UserPass);

            //byte[] binFormData = System.Convert.FromBase64String(someBase64StringData);

            XmlHTTP.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary);
            XmlHTTP.send(binFormData);

            string strResponseText = XmlHTTP.responseText;

            m_log.Info(strResponseText);
            return strResponseText;
            //WebResponse wresp = null;
            //try
            //{
            //    wresp = wr.GetResponse();
            //    System.IO.Stream stream2 = wresp.GetResponseStrem();
            //    System.IO.StreamReader reader2 = new System.IO.StreamReader(stream2);
            //    m_log.Info(reader2.ReadToEnd());
            //    //log.Debug(string.Format("File uploaded, server response is: {0}", reader2.ReadToEnd()));
            //}
            //catch (Exception ex)
            //{
            //    //log.Error("Error uploading file", ex);
            //    if (wresp != null)
            //    {
            //        wresp.Close();
            //        wresp = null;
            //    }
            //}
            //finally
            //{
            //    wr = null;
            //}

        }
    }
}
