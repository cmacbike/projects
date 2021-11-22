//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Web;
//using System.Net;
 
//namespace StarmineXML
//{
//    class UploadTemplate
//    {
//        public static void HttpUploadFile(string url, string file, string paramName, string contentType, NameValueCollection nvc)
//        {
            
//            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
//            byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

//            HttpWebRequest wr = (HttpWebRequest)WebRequest.Create(url);
//            wr.ContentType = "multipart/form-data; boundary=" + boundary;
//            wr.Method = "POST";
//            wr.KeepAlive = true;
//            wr.Credentials = System.Net.CredentialCache.DefaultCredentials;

//            System.IO.Stream rs = wr.GetRequestStream();

//            string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
//            foreach (string key in nvc.Keys)
//            {
//                rs.Write(boundarybytes, 0, boundarybytes.Length);
//                string formitem = string.Format(formdataTemplate, key, nvc[key]);
//                byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes(formitem);
//                rs.Write(formitembytes, 0, formitembytes.Length);
//            }
//            rs.Write(boundarybytes, 0, boundarybytes.Length);

//            string headerTemplate = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n";
//            string header = string.Format(headerTemplate, paramName, file, contentType);
//            byte[] headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
//            rs.Write(headerbytes, 0, headerbytes.Length);

//            System.IO.FileStream fileStream = new System.IO.FileStream(file, System.IO.FileMode.Open, System.IO.FileAccess.Read);
//            byte[] buffer = new byte[4096];
//            int bytesRead = 0;
//            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
//            {
//                rs.Write(buffer, 0, bytesRead);
//            }
//            fileStream.Close();

//            byte[] trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
//            rs.Write(trailer, 0, trailer.Length);
//            rs.Close();

//            WebResponse wresp = null;
//            try
//            {
//                wresp = wr.GetResponse();
//                System.IO.Stream stream2 = wresp.GetResponseStream();
//                System.IO.StreamReader reader2 = new System.IO.StreamReader(stream2);
//                //log.Debug(string.Format("File uploaded, server response is: {0}", reader2.ReadToEnd()));
//            }
//            catch (Exception ex)
//            {
//                //log.Error("Error uploading file", ex);
//                if (wresp != null)
//                {
//                    wresp.Close();
//                    wresp = null;
//                }
//            }
//            finally
//            {
//                wr = null;
//            }
//        }
//    }
//}
