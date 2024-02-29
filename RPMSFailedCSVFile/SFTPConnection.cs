using Renci.SshNet;
using RPMSScrapExcel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMSFailedCSVFile
{
     class SFTPConnection
    {
        public bool SFTPFileDownload(string remoteFileName, string remoteDirectory)
        {
            bool result = false;
            string host = Convert.ToString(ConfigurationManager.AppSettings["host"]);
            string password = Convert.ToString(ConfigurationManager.AppSettings["pwd"]);
            string username = Convert.ToString(ConfigurationManager.AppSettings["user"]);
            // string remoteDirectory = Convert.ToString(ConfigurationManager.AppSettings["dir"]);
            int port = Convert.ToInt16(ConfigurationManager.AppSettings["port"]);
            try
            {
                using (SftpClient sftp = new SftpClient(host, port, username, password))
                {
                    sftp.Connect();
                    Log.writelog("Connection Open");
                    //Download Files                
                    using (var file = File.OpenWrite(Convert.ToString(AppDomain.CurrentDomain.BaseDirectory).Replace("bin\\Debug\\", "") + "Excel\\" + remoteFileName))
                    {
                        sftp.DownloadFile(remoteDirectory + remoteFileName, file);
                        Log.writelog(remoteFileName + " File Downloaded Successfully");
                    }

                   

                    sftp.Disconnect();
                    result = true;
                };
            }
            catch (Exception ex)
            {
                result = false;
                Log.writelog("SFTP Error File Not Downloaded " + ex.Message);
                Log.Insert_Log(remoteFileName, "Failure", "SFTP Error File Not Downloaded " + ex.Message);
            }

            return result;
        }

    }
}
