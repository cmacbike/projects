using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
using log4net;
namespace StarmineXML
{
    class UploadStarmineFiles
    {
        private ILog m_log;
        private string boundary;
        public UploadStarmineFiles(ILog p_log)
        {
            m_log = p_log;
        }
        public string UploadFileHTTP(string p_strURL,string p_strFile,Dictionary<string,string> p_Fields)
        {
            string formdata;
            boundary = "---------------------------0123456789012";
            formdata = "--" + boundary + "\r\n";
             m_log.Info(formdata);
            byte[] boundarybytes = System.Text.Encoding.UTF8.GetBytes("\r\n--" + boundary + "\r\n");
            var boundary64String = Convert.ToBase64String(boundarybytes);
            boundarybytes = Convert.FromBase64String(boundary64String);
            m_log.Info("\r\n--" + boundary + "\r\n");
            HttpWebRequest wr = (HttpWebRequest)WebRequest.Create(p_strURL);
            wr.ContentType = "multipart/form-data; boundary=" + boundary;
            m_log.Info(wr.ContentType);
            wr.Method = "POST";
            m_log.Info(wr.Method);
            wr.KeepAlive = true;
            wr.Credentials = new NetworkCredential(Utility.UserName, Utility.UserPass);
            //wr.Credentials = System.Net.CredentialCache.DefaultCredentials;
           
            System.IO.Stream rs = wr.GetRequestStream();
            m_log.Info(wr.Headers.ToString());
            

            foreach (KeyValuePair<string,string> Pair in p_Fields)
            {
                
                formdata = string.Format("Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}",Pair.Key,Pair.Value);
                formdata = formdata + "\r\n--" + boundary + "\r\n";
                m_log.Info(formdata);
                byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes(formdata);
                var formBase64string = Convert.ToBase64String(formitembytes);
                formitembytes = Convert.FromBase64String(formBase64string);
                rs.Write(formitembytes, 0, formitembytes.Length);
               
            }

            if (!p_strFile.Equals(""))
            {
                string header = string.Format("Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: application/upload\r\n\r\n", "userfile", p_strFile);
                m_log.Info(header);

                byte[] headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
                var headerbase64String = Convert.ToBase64String(headerbytes);
                headerbytes = Convert.FromBase64String(headerbase64String);
                rs.Write(headerbytes, 0, headerbytes.Length);

                System.IO.FileStream FileToUpload = new System.IO.FileStream(p_strFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                byte[] FileContents = new byte[FileToUpload.Length];
                System.IO.BinaryReader bReader = new System.IO.BinaryReader(FileToUpload);
                FileContents = bReader.ReadBytes(System.Convert.ToInt32(FileToUpload.Length));

                rs.Write(FileContents, 0, FileContents.Length);

                byte[] filebuffer = new byte[4096];
                int bytesRead = 0;
                while ((bytesRead = FileToUpload.Read(filebuffer, 0, filebuffer.Length)) != 0)
                {
                    rs.Write(filebuffer, 0, bytesRead);
                }

                FileToUpload.Close();

                byte[] trailer = System.Text.Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");
                var trailerbase64String = Convert.ToBase64String(trailer);
                trailer = Convert.FromBase64String(trailerbase64String);
                rs.Write(trailer, 0, trailer.Length);
            }
            rs.Close();
           
            
            WebResponse wresp = null;
            
            wresp = wr.GetResponse();
            System.IO.Stream stream2 = wresp.GetResponseStream();
            System.IO.StreamReader reader2 = new System.IO.StreamReader(stream2);
            string Results = reader2.ReadToEnd();
            m_log.Info(Results);
            return Results;
                //log.Debug(string.Format("File uploaded, server response is: {0}", reader2.ReadToEnd()));
            

        }
    }
}
