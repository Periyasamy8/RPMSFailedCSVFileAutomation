using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace RPMSFailedCSVFile
{
    class ImportExcelBAL
    {
        public bool Insert_Excel(DataTable dt,string filename)
        {
            bool result = false;
            ImportExcelDAL _dal = new ImportExcelDAL();
            result = _dal.Insert_Excel(dt, filename);
            return result;
        }
    }
}
