using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report
{
    public class cArx
    {
        public int Days { get; set; }
        public string EMail { get; set; }        
        public string DateFormatFile { get; set; } = "yyyyMM";
        public DateTime FirstDayMonth { get {return new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1); } }   //.AddMonths(1).AddDays(-1)

        public string strLastDayMonth { get { return $"{FirstDayMonth:dd.MM.yyyy}"; } }
        public string PathMove { get; set; }
        public int Row { get; set; }
    }
}
