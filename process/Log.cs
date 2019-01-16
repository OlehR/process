using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Process
{
    public static class Log
    {
        public static void StrToFile(string cFileName, string cExpression)
        {
            StrToFile(cFileName, cExpression, FileMode.CreateNew);
        }

        public static void log(string cExpression, string parFile = null)
        {
            if (parFile == null)
                if (GlobalVar.varFileLog == null)
                    parFile = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) +
                                          "\\log\\process_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                else
                    parFile = GlobalVar.varFileLog;
            try
            {
                DateTime now = DateTime.Now;
                StreamWriter sw;
                FileInfo fi = new FileInfo(parFile);
                sw = fi.AppendText();
                sw.WriteLine(now.ToString() + "=>" + cExpression);
                sw.Flush();
                sw.Close();
                Console.WriteLine(now.ToString() + "=>" + cExpression);

                //             StrToFile(Environment.GetEnvironmentVariable("temp") + "\\log_process.txt", "\n" + now.ToString() + "=>" + cExpression, FileMode.OpenOrCreate);
            }
            catch
            {
            };

        }

        static public void StrToFile(string cFileName, string cExpression, System.IO.FileMode parFileMode)
        {
            if ((System.IO.File.Exists(cFileName) == true) && (System.IO.FileMode.OpenOrCreate != parFileMode))
            {
                //If so then Erase the file first as in this case we are overwriting
                System.IO.File.Delete(cFileName);
            }

            //Create the file if it does not exist and open it
            FileStream oFs = new FileStream(cFileName, parFileMode, FileAccess.ReadWrite);

            //Create a writer for the file
            StreamWriter oWriter = new StreamWriter(oFs);

            //Write the contents
            oWriter.Write(cExpression);
            oWriter.Flush();
            oWriter.Close();
            oFs.Close();
        }
    }

}
