using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace RPMSScrapExcel
{
    public static class Log 
    {
        public static void writelog(string message)
        {
            string log= Convert.ToString(ConfigurationManager.AppSettings["Path"]);
            var path = Convert.ToString(AppDomain.CurrentDomain.BaseDirectory).Replace("bin\\Debug\\", "") + log;
            string sTemp = path + "_" + DateTime.Now.ToString("dd_MM") + ".txt";
            FileStream Fs = new FileStream(sTemp, FileMode.OpenOrCreate | FileMode.Append);
            StreamWriter st = new StreamWriter(Fs);
            string dttemp = DateTime.Now.ToString("[dd:MM:yyyy] [HH:mm:ss:ffff]");
            st.WriteLine(dttemp + "\t" + message);
            st.Close();
        }

        public static ConnectionStringSettings sql_cs = ConfigurationManager.ConnectionStrings["dbConnectionString"];
        public static SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);
        public static void conn()
        {
            try
            {
                if (con.State != ConnectionState.Open)
                {
                    con.Close();
                }
                con.Open();
            }
            catch (Exception)
            {
                if (con.State != ConnectionState.Closed)
                {
                    con.Close();
                    con.Open();
                }
            }
        }
        public static bool ExecuteCmd(SqlCommand sqlCmd, ConnectionStringSettings cs)
        {
            bool _wasSuccessful = false;
            try
            {
                SqlConnection sqlConn = new SqlConnection(cs.ToString());
                sqlCmd.Connection = sqlConn;
                try
                {
                    sqlConn.Open();
                    sqlCmd.ExecuteNonQuery();
                    _wasSuccessful = true;
                }
                catch (SqlException ex)
                {
                    Log.writelog(ex.Message);

                }
                finally
                {
                    sqlConn.Close();
                }
            }
            catch (Exception ex)
            {
                Log.writelog(ex.Message);
            }
            return _wasSuccessful;
        }
        
        public static bool Insert_Log(string filename,string status,string errormessage)
        {
            SqlCommand cmd = new SqlCommand("Asp_ExcelUploadHistroy", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@FileName", filename);
            cmd.Parameters.AddWithValue("@Status", status);
            cmd.Parameters.AddWithValue("@ErrorMessage", errormessage);
            cmd.Parameters.AddWithValue("@UploadedBy", 0);
            return ExecuteCmd(cmd, sql_cs);
        }
    }
}
