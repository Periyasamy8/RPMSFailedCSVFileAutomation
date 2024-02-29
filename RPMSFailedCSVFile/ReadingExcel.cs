using LumenWorks.Framework.IO.Csv;
using RPMSScrapExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;




namespace RPMSFailedCSVFile
{
    class ReadingExcel
    {
        public DataTable ReadExceldata(string remoteFileName)
        {
            DataTable dt = DT();
            string FullPathFile = Convert.ToString(AppDomain.CurrentDomain.BaseDirectory).Replace("bin\\Debug\\", "") + "Excel\\" + remoteFileName;
            try
            {

                using (CsvReader csv =
               new CsvReader(new StreamReader(FullPathFile), true))
                {
                    int fieldCount = csv.FieldCount;

                    string[] headers = csv.GetFieldHeaders();
                    while (csv.ReadNextRecord())
                    {
                        DataRow _new = dt.NewRow();
                        for (int i = 0; i < fieldCount; i++)
                        {
                            _new["DataPocketNo"] = Convert.ToString(csv[i]);
                            _new["ENGModel"] = Convert.ToString(csv[i + 1]);
                            _new["TMModel"] = Convert.ToString(csv[i + 2]);
                            _new["DOB"] = Convert.ToString(csv[i + 3]);
                            _new["CC"] = Convert.ToString(csv[i + 4]);
                            _new["InvoiceNo"] = Convert.ToString(csv[i + 5]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["DealerCode"] = Convert.ToString(csv[i + 6]);
                            _new["MaterialDesc"] = Convert.ToString(csv[i + 7]);
                            _new["VinNo"] = Convert.ToString(csv[i + 8]);
                            _new["PQRNO"] = Convert.ToString(csv[i + 9]);
                            _new["ClaimNO"] = Convert.ToString(csv[i + 10]);
                            _new["SafetyIssue"] = Convert.ToString(csv[i + 11]);
                            _new["MaterialPartNo"] = Convert.ToString(csv[i + 12]);
                            _new["DamageCode"] = Convert.ToString(csv[i + 13]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["DistributorName"] = Convert.ToString(csv[i + 14]);
                            _new["EMGTM"] = Convert.ToString(csv[i + 15]);
                            _new["MileageTime"] = Convert.ToString(csv[i + 16]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["DOC"] = Convert.ToString(csv[i + 17]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["FailureDescDealer"] = Convert.ToString(csv[i + 18]);
                            _new["FailureDescMFTBC"] = Convert.ToString(csv[i + 19]);
                            _new["GroupChargeInvestigation"] = Convert.ToString(csv[i + 20]);
                            _new["FailureDamageDesc"] = Convert.ToString(csv[i + 21]);
                            _new["Model"] = Convert.ToString(csv[i + 22]);
                            _new["WPIClaimNo"] = Convert.ToString(csv[i + 23]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["QissueID"] = Convert.ToString(csv[i + 24]);
                            _new["Qissuestatus"] = Convert.ToString(csv[i + 25]);
                            _new["QissueImportancelevel"] = Convert.ToString(csv[i + 26]);
                            _new["Builddate"] = Convert.ToString(csv[i + 27]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["Solddate"] = Convert.ToString(csv[i + 28]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["Failuredate"] = Convert.ToString(csv[i + 29]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["Positioncode"] = Convert.ToString(csv[i + 30]).Replace("\"", string.Empty).Replace("=", string.Empty).Trim();
                            _new["Currentstatus"] = Convert.ToString(csv[i + 31]);



                            break;
                        }
                        dt.Rows.Add(_new);
                    }
                }
            }

            catch (Exception ex)
            {
                Log.writelog("Reading Excel Issue " + ex.Message);
                Log.Insert_Log(remoteFileName, "Failure", "Reading Excel Issue " + ex.Message);

            }
            return dt;
        }

        public DataTable DT()
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("DataPocketNo");
            dt.Columns.Add("DataPocketBranch");
            dt.Columns.Add("DOR");
            dt.Columns.Add("Registrant");
            dt.Columns.Add("ENGModel");
            dt.Columns.Add("TMModel");
            dt.Columns.Add("QuanityINW");
            dt.Columns.Add("StorageLocb4");
            dt.Columns.Add("StorageLocaft4");
            dt.Columns.Add("Return");
            dt.Columns.Add("DOI");
            dt.Columns.Add("DispatchPerson");
            dt.Columns.Add("ReturnDate");
            dt.Columns.Add("QuanityPAR");
            dt.Columns.Add("ReceiptOfReturnParts");
            dt.Columns.Add("CheckResults");
            dt.Columns.Add("POS");
            dt.Columns.Add("DOSI");
            dt.Columns.Add("ROSI");
            dt.Columns.Add("DOB");
            dt.Columns.Add("CC");
            dt.Columns.Add("InvoiceNo");
            dt.Columns.Add("OrderNo");
            dt.Columns.Add("DealerCode");
            dt.Columns.Add("MaterialDesc");
            dt.Columns.Add("VinNo");
            dt.Columns.Add("ENGTMMang");
            dt.Columns.Add("PQRNO");
            dt.Columns.Add("ClaimNO");
            dt.Columns.Add("ReturnReqFromDealer");
            dt.Columns.Add("SafetyIssue");
            dt.Columns.Add("MaterialPartNo");
            dt.Columns.Add("DamageCode");
            dt.Columns.Add("DistributorName");
            dt.Columns.Add("EMGTM");
            dt.Columns.Add("MileageTime");
            dt.Columns.Add("DOC");
            dt.Columns.Add("FailureDescDealer");
            dt.Columns.Add("FailureDescMFTBC");
            dt.Columns.Add("GroupChargeInvestigation");
            dt.Columns.Add("FailureDamageDesc");
            dt.Columns.Add("Model");
            dt.Columns.Add("ApplicationDivision");
            dt.Columns.Add("WPIClaimNo");
            dt.Columns.Add("NeedInvestigation");
            dt.Columns.Add("PIC");
            dt.Columns.Add("SafetyParts");
            //new
            dt.Columns.Add("Location");

            //New
            dt.Columns.Add("QissueID");
            dt.Columns.Add("Qissuestatus");
            dt.Columns.Add("QissueImportancelevel");
            dt.Columns.Add("Builddate");
            dt.Columns.Add("Solddate");
            dt.Columns.Add("Failuredate");
            dt.Columns.Add("Positioncode");
            dt.Columns.Add("Currentstatus");

            return dt;
        }
    }


}
