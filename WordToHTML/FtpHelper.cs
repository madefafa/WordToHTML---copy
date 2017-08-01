using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WordToHTML
{
  public  class FtpHelper
    {
        private string userIdi;
        private string pwdi;
        private string ftpPathi;

        public static readonly FtpHelper Instance = new FtpHelper();

        /// <summary>  
        /// 取得文件名  
        /// </summary>  
        /// <param name="ftpPath">ftp路径</param>  
        /// <returns></returns>  
        public string[] GetFilePath(string userId, string pwd, string ftpPath)
        {
            userIdi = userId;
            pwdi = pwd;
            ftpPathi = ftpPath;
            string[] downloadFiles;
            StringBuilder result = new StringBuilder();
            FtpWebRequest reqFTP;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPath));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(userId, pwd);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectory;
                reqFTP.UsePassive = true;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }
                result.Remove(result.ToString().LastIndexOf('\n'), 1);
                reader.Close();
                response.Close();
                return result.ToString().Split('\n');
        }
            catch (Exception ex)
            {
                string err = ex.ToString();
        downloadFiles = null;
                return downloadFiles;
            }
}

        //ftp的上传功能  
        public void Upload(string userId, string pwd, string filename, string ftpPath)
        {
            FileInfo fileInf = new FileInfo(filename);
            FtpWebRequest reqFTP;
            // 根据uri创建FtpWebRequest对象   
            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPath +"//"+ fileInf.Name));
            // ftp用户名和密码  
            reqFTP.Credentials = new NetworkCredential(userId, pwd);

            reqFTP.UsePassive = false;
            // 默认为true，连接不会被关闭  
            // 在一个命令之后被执行  
            reqFTP.KeepAlive = false;
            // 指定执行什么命令  
            reqFTP.Method = WebRequestMethods.Ftp.UploadFile;
            // 指定数据传输类型  
            reqFTP.UseBinary = true;
            // 上传文件时通知服务器文件的大小  
            reqFTP.ContentLength = fileInf.Length;
            // 缓冲大小设置为2kb  
            int buffLength = 2048;
            byte[] buff = new byte[buffLength];
            int contentLen;
            // 打开一个文件流 (System.IO.FileStream) 去读上传的文件  
            FileStream fs = fileInf.OpenRead();
            try
            {
                // 把上传的文件写入流  
                Stream strm = reqFTP.GetRequestStream();
                // 每次读文件流的2kb  
                contentLen = fs.Read(buff, 0, buffLength);
                // 流内容没有结束  
                while (contentLen != 0)
                {
                    // 把内容从file stream 写入 upload stream  
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, buffLength);
                }
                // 关闭两个流  
                strm.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                string err = ex.ToString();
            }
        }

        public void Delete(string userId, string pwd, string ftpPath, string fileName)
        {
            FtpWebRequest reqFTP;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPath + fileName));
                reqFTP.Method = WebRequestMethods.Ftp.DeleteFile;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(userId, pwd);
                reqFTP.UsePassive = false;
                FtpWebResponse listResponse = (FtpWebResponse)reqFTP.GetResponse();
                string sStatus = listResponse.StatusDescription;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //从ftp服务器上下载文件的功能  
        public void Download(string userId, string pwd, string ftpPath, string filePath, string fileName)
        {
            FtpWebRequest reqFTP;
            try
            {
                FileStream outputStream = new FileStream(filePath + "\\" + fileName, FileMode.Create);
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPath + fileName));
                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(userId, pwd);
                reqFTP.UsePassive = false;
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                long cl = response.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[bufferSize];
                readCount = ftpStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);
                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                }
                ftpStream.Close();
                outputStream.Close();
                response.Close();


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 创建文件夹
        /// </summary>
        /// <param name="dirName"></param>
        public void MakeDir(string dirName)
        {
            FtpWebRequest reqFTP;
            try
            {
                // dirName = name of the directory to create.
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPathi +"\\"+ dirName));
                reqFTP.Method = WebRequestMethods.Ftp.MakeDirectory;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(userIdi, pwdi);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();

                ftpStream.Close();
                response.Close();
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }


        /// <summary>
        /// 判断当前目录下指定的子目录是否存在
        /// </summary>
        /// <param name="RemoteDirectoryName">指定的目录名</param>
        public bool DirectoryExist(string RemoteDirectoryName)
        {
            string[] dirList = GetDirectoryList();
            foreach (string str in dirList)
            {
                if (str.Trim() == RemoteDirectoryName.Trim())
                {
                    return true;
                }
            }
            return false;
        }


        public string[] GetDirectoryList()
        {
            string[] drectory = GetFilesDetailList();
            if (drectory == null)
            {
                MakeDir("test");
                drectory= GetFilesDetailList();
            }
            string m = string.Empty;
            foreach (string str in drectory)
            {
                int dirPos = str.IndexOf("<DIR>");
                if (dirPos > 0)
                {
                    /*判断 Windows 风格*/
                    m += str.Substring(dirPos + 5).Trim() + "\n";
                }
                else if (str.Trim().Substring(0, 1).ToUpper() == "D")
                {
                    /*判断 Unix 风格*/
                    string dir = str.Substring(54).Trim();
                    if (dir != "." && dir != "..")
                    {
                        m += dir + "\n";
                    }
                }
            }

            char[] n = new char[] { '\n' };
            return m.Split(n);
        }

        /// <summary>
        /// 获取当前目录下明细(包含文件和文件夹)
        /// </summary>
        /// <returns></returns>
        public string[] GetFilesDetailList()
        {
            string[] downloadFiles;
            try
            {
                StringBuilder result = new StringBuilder();
                FtpWebRequest ftp;
                ftp = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPathi));
                ftp.Credentials = new NetworkCredential(userIdi, pwdi);
                ftp.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                WebResponse response = ftp.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);

                //while (reader.Read() > 0)
                //{

                //}
                string line = reader.ReadLine();
                //line = reader.ReadLine();
                //line = reader.ReadLine();

                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }
                result.Remove(result.ToString().LastIndexOf("\n"), 1);
                reader.Close();
                response.Close();
                return result.ToString().Split('\n');
            }
            catch (Exception ex)
            {
                downloadFiles = null;
                //Insert_Standard_ErrorLog.Insert("FtpWeb", "GetFilesDetailList Error --> " + ex.Message);
                return downloadFiles;
            }
        }

    }
}
