using RPMSScrapExcel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMSFailedCSVFile
{
    public class ReadingCSV
    {
        public void OnStart() {

            string mm = string.Empty;
            int month = DateTime.Now.Month;
            string dd = string.Empty;
            string remoteFileName = string.Empty;
            string remoteDirectory = string.Empty;
            int day = DateTime.Now.Day;
            if (month < 10)
                mm = "0" + Convert.ToString(month);
            else
                mm = Convert.ToString(month);
            if (day < 10)
                dd = "0" + Convert.ToString(day);
            else
                dd = Convert.ToString(day);
            SFTPConnection _con = new SFTPConnection();
            ReadingExcel _ex = new ReadingExcel();
            ImportExcelBAL _bal = new ImportExcelBAL();
            DataTable dt = new DataTable();
            try
            {
                Log.writelog("Job reached try statement..");
                 remoteFileName = "RPMS-2_" + Convert.ToString(DateTime.Now.Year) + mm + dd + ".csv";
                // remoteFileName = Convert.ToString(ConfigurationManager.AppSettings["filenameprefix"]);
                Log.writelog("File name : " + remoteFileName);
                string FullPathFile = Convert.ToString(ConfigurationManager.AppSettings["ActualFilePath"]) + remoteFileName;
                Log.writelog("Actual File path : " + FullPathFile);
                if (System.IO.File.Exists(FullPathFile))
                {
                    Log.writelog("File Exists in the Actual path so no need to Load this file. " + remoteFileName);
                    Log.writelog("Job Stopped");
                }
                else
                {
                    Log.writelog("Failed file loading again " + remoteFileName);
                    remoteDirectory = Convert.ToString(ConfigurationManager.AppSettings["dir"]);
                    
                    bool connectionresult = _con.SFTPFileDownload(remoteFileName, remoteDirectory);
                    Log.writelog("File download from remote directory through FTP - " + connectionresult);

                    dt = _ex.ReadExceldata(remoteFileName);
                    Log.writelog("Getting data from CSV file to Dataset");

                    bool result = _bal.Insert_Excel(dt, remoteFileName);
                    Log.writelog("Final Status " + result.ToString());
                    Log.Insert_Log(remoteFileName, "Success", "");
                }
            }
            catch (Exception ex)
            {
                Log.writelog("Final Status error - " + ex.Message);
                Log.Insert_Log(remoteFileName, "Error", "");
            }

        }
    }
}
