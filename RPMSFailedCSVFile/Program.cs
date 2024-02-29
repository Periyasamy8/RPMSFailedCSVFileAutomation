using RPMSScrapExcel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMSFailedCSVFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Log.writelog("Job Started..");
            ReadingCSV obj= new ReadingCSV();
            obj.OnStart();
        }
    }
}
