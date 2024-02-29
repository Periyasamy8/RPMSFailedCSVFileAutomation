using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace RPMSFailedCSVFile
{
    class ImportExcelDAL:DataComponent
    {
        public bool Insert_Excel(DataTable dt,string filename)
        {
            SqlCommand cmd = new SqlCommand("InsertRPMSExcelData", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@tblRPMSExcelData", dt);
            cmd.Parameters.AddWithValue("@FileName", filename);   
            return ExecuteCmd(cmd, sql_cs);
        }
    }
}
