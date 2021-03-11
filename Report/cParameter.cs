using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report
{
    public class cParameter
    {
        public string Name { get; set; }
        public string EMail { get; set; }
        public string Par1 { get; set; }
        public string Par2 { get; set; }
        public string FileName { get; set; }

        public DateTime DateReport { get; set; } = DateTime.Now.Date;
        public string DateFormatFile { get; set; } = "yyyyMMdd";
        //$"{dt:M/d/yyyy}";
        public string strDateReportFile { get { return DateReport.AddDays(-1).ToString(DateFormatFile ); } }
        public string PathMove { get; set; }
        public string CopyFileName { get; set; }
        public cParameter() { }
        public cParameter(cParameter pP ) 
        {
            Name = pP.Name;
            EMail = pP.EMail;
            Par1 = pP.Par1;
            Par2 = pP.Par2;
            FileName = pP.FileName;
            DateReport = pP.DateReport;
            DateFormatFile = pP.DateFormatFile;
            PathMove = pP.PathMove;
        }

    }
}
