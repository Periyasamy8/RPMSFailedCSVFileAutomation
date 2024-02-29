using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using RPMSScrapExcel;

namespace RPMSFailedCSVFile
{
    public class DataComponent
    {
        public DataComponent()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public static ConnectionStringSettings sql_cs = ConfigurationManager.ConnectionStrings["dbConnectionString"];
        public static SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);
        public void conn()
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
            catch(Exception ex)
            {
                Log.writelog(ex.Message);
            }
            return _wasSuccessful;
        }

        public static DataSet SelectCmd(SqlCommand sqlCmd, ConnectionStringSettings cs)
        {
            DataSet dsResults = new DataSet();
            SqlDataAdapter sqlAdapter = new SqlDataAdapter();

            try
            {
                SqlConnection sqlConn = new SqlConnection(cs.ToString());
                sqlCmd.Connection = sqlConn;
                try
                {
                    sqlCmd.CommandTimeout = 2000;
                    sqlAdapter.SelectCommand = sqlCmd;
                    sqlAdapter.Fill(dsResults);
                }
                catch (SqlException ex)
                {
                    Log.writelog(ex.Message);
                    throw ex;
                }

                finally
                {
                    sqlConn.Close();
                    sqlConn.Dispose();
                    sqlCmd.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log.writelog(ex.Message);
            }
            return dsResults;
        }
    }
}