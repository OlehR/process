using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices;
namespace Process
{
    class MyXMLA
    {
        public static void SetProcessTypeDimension(string parStr)
        {
            string varStr = parStr.Trim().ToUpper();
            GlobalVar.varIsProcessDimension = true;
            if (varStr == "NONE")
                GlobalVar.varIsProcessDimension = false;
            else if (varStr == "UPDATE")
                GlobalVar.varProcessDimension = ProcessType.ProcessUpdate;
            else if (varStr == "FULL")
                GlobalVar.varProcessDimension = ProcessType.ProcessFull;
        }
        public static void SetProcessTypeCube(string parStr)
        {
            string varStr = parStr.Trim().ToUpper();
            if (varStr == "NONE")
                GlobalVar.varIsProcessCube = false;
            else if (varStr == "DATA")
                GlobalVar.varProcessCube = ProcessType.ProcessData;
            else if (varStr == "FULL")
                GlobalVar.varProcessCube = ProcessType.ProcessFull;
        }



    }
}