using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices;

namespace Process
{
    class GlobalVar
    {
        public static string varServer = "localhost", varDB = "dw_olap", varCube = null;
        public static string varPrepareSQL = null, varWaitSQL = null, varConectSQL = null;
        public static ProcessType varProcessDimension = ProcessType.ProcessUpdate, varProcessCube = ProcessType.ProcessFull;
        public static bool varIsProcessDimension = false, varIsProcessCube = true;
        public static string varFileLog = null;
        public static string varFileXML = null;
        public static string varKeyErrorLogFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\log\\Error_" + DateTime.Now.ToString("yyyyMMdd") + ".log";
        public static int varStep = 0, varMetod = 0;
        public static int varTimeStart = 0, varTimeEnd = 24;
        public static int varMaxParallel = 0;
        public static int varDayProcess = -1;
        public static DateTime varDateStartProcess = DateTime.Now;
        public static bool varIsArx = false;
        public static DateTime varArxDate = new DateTime(1, 1, 1);
        public static int varDefaultDayProcess = -1;
     
    }
}
