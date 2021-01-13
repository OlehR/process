using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "office"
using ExcelApp = Microsoft.Office.Interop.Excel; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "Microsoft.Office.Interop.Excel"

namespace Report
{
    public class Excel
    {
        Mail Mail = new Mail();
        MsSQL MsSQL = new MsSQL();
        public bool ExecuteExcelsMacro(string pSource)
        {
            string[] Files=null;
            bool Result = true;
            
            StringBuilder Success = new StringBuilder($"Start {DateTime.Now} {pSource}{Environment.NewLine}"), Error = new StringBuilder();
            try
            {
                // get the file attributes for file or directory
                FileAttributes attr = File.GetAttributes(pSource);

                if (attr.HasFlag(FileAttributes.Directory))
                {
                    if (CreateResultDirectory(pSource, Success, Error))
                        Files = Directory.GetFiles(pSource, "*.xls*");
                    else
                        Result = false;
                }
                else
                {
                    var receiptFilePath = Path.GetDirectoryName(pSource);
                    CreateResultDirectory(receiptFilePath, Success, Error);
                    Files = new string[] { pSource };
                    //MessageBox.Show("Its a file");
                }
                if (Files != null)
                {
                    foreach (var el in Files)
                        ExecuteExcelMacro(el, Success, Error);
                }
            }
            catch (Exception ex)
            {
                Result = false;
                Error.Append(ex.Message + Environment.NewLine);
                Error.Append(Environment.StackTrace + Environment.NewLine);
            }

            Success.Append($"End {DateTime.Now} {pSource}{Environment.NewLine}");
                if (Error != null && Error.Length > 0)
                    Console.WriteLine(Error.ToString());
                Console.WriteLine(Success.ToString());
            string DT = DateTime.Now.ToString("yyyyMMddHHmmss");
            string FileName = Path.Combine(Path.GetDirectoryName(pSource), "Result", $"Log_{DT}.log");
            File.WriteAllText(FileName, Error.ToString() + Environment.NewLine + Success.ToString());

            return Result;
        }
        bool CreateResultDirectory(string pSourceDirectory, StringBuilder pSuccess, StringBuilder pError)
        {
            bool Result = true;
            try
            {
                var Destination = Path.Combine(pSourceDirectory, "Result");
                if (!Directory.Exists(Destination))
                {
                    Directory.CreateDirectory(Destination);
                    pSuccess.Append($"Create Directory {Destination}");
                }
            }
            catch (Exception ex)
            {
                Result = false;
                pError.Append(ex.Message+ Environment.NewLine);
                pError.Append(Environment.StackTrace+ Environment.NewLine);
            }
            return Result;
        }
        public void ExecuteExcelMacro(string pSourceFile, StringBuilder pSuccess,StringBuilder pError)
        {
            ExcelApp.Application ExcelApp = null;
            ExcelApp.Workbook ExcelWorkBook = null;
            IEnumerable<cParameter> ResPar = null;
            bool Result = true;
            
            try
            {
               
                сRequest ParRequest = null;
                List<сRequest> Requests = new List<сRequest>();
                string Macro = "Main", StartMacro=null;
                string Email = null;               

                ExcelApp = new ExcelApp.Application();
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Visible = false;
                //ExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
                ExcelWorkBook = ExcelApp.Workbooks.Open(pSourceFile);

                ExcelApp.Worksheet worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets["config"];

                var range = worksheet.UsedRange;
                int rows_count = range.Rows.Count;


                for (int i = 1; i <= rows_count; i++)
                {
                    string str = worksheet.Cells[i, 1].value;
                    //string str = String.Format("[{0}] ", s);
                    if (str != null)
                    {
                        if (str.Equals("StartMacro"))
                            StartMacro = worksheet.Cells[i, 2].value;

                        if (str.Equals("Macro"))
                            Macro = worksheet.Cells[i, 2].value;
                        if (str.Equals("e-mail"))
                            Email = worksheet.Cells[i, 2].value;
                        else
                         if (str.Equals("pSQL"))
                            ParRequest = GetRequest(worksheet, i, eClient.MsSql, true);
                        else
                          if (str.Equals("pMDX"))
                            ParRequest = GetRequest(worksheet, i, eClient.MDX, true);
                        else
                        if (str.Equals("SQL"))
                            Requests.Add(GetRequest(worksheet, i, eClient.MsSql));
                        else
                          if (str.Equals("MDX"))
                            Requests.Add(GetRequest(worksheet, i, eClient.MDX));
                    }
                }

                //var r = pSourceFile.Split('.');
                var path = Path.Combine(Path.GetDirectoryName(pSourceFile), "Result");
                var FileName = Path.GetFileNameWithoutExtension(pSourceFile);
                var Extension = Path.GetExtension(pSourceFile);

                if (ParRequest != null)
                    ResPar = MsSQL.RunMsSQL(ParRequest);
                else
                    ResPar = new List<cParameter>() { new cParameter() { EMail = Email,Name="" } };

                foreach (var el in ResPar)
                {
                    if (ParRequest != null)
                    {
                        worksheet.Cells[ParRequest.Row, ParRequest.Column].value = el.Par1;
                        worksheet.Cells[ParRequest.Row, ParRequest.Column + 1].value = el.Name;
                        worksheet.Cells[ParRequest.Row, ParRequest.Column + 2].value = el.EMail;
                        if (!string.IsNullOrEmpty(el.Par2))
                            worksheet.Cells[ParRequest.Row, ParRequest.Column + 3].value = el.Par2;
                    }
                    ExcelApp.Run(Macro);
                    el.FileName = Path.Combine(path, FileName + "_" + el.Name.Trim() + Extension);
                    if (File.Exists(el.FileName))
                        File.Delete(el.FileName);
                    ExcelWorkBook.SaveAs(el.FileName);
                    pSuccess.Append($"{DateTime.Now} Save file {el.FileName}{Environment.NewLine}");
                }
            }
            catch (Exception ex)
            {
                pError.Append(ex.Message + Environment.NewLine);
                pError.Append(Environment.StackTrace + Environment.NewLine);
                Result = false;
            }
            finally
            {
                // Закриваємо ексель.
                if(ExcelWorkBook!=null)
                    ExcelWorkBook.Close(false);
                if (ExcelApp!=null)
                    ExcelApp.Quit();
                if (ExcelWorkBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook); }
                if (ExcelApp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp); }
            }
            //Відправляємо Листи
            if (Result && ResPar != null)
            {
                try
                {
                    foreach (var el in ResPar)
                    {
                        var emails = el.EMail.Split(',');
                        foreach(var email in emails)
                            Mail.SendMail(email, el.FileName, null, null, pSuccess, pError);
                    }
                }
                catch (Exception ex)
                {
                    pError.Append(ex.Message + Environment.NewLine);
                    pError.Append(Environment.StackTrace + Environment.NewLine);
                }
            }

        }

        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
        }

        static private сRequest GetRequest(ExcelApp.Worksheet worksheet, int pInd, eClient pClient = eClient.NotDefine, bool IsPar = false)
        {
            string Request = worksheet.Cells[pInd, 4].value;
            string Sheet = IsPar ? "config" : worksheet.Cells[pInd, 5].value;
            double Column = IsPar ? 5 : worksheet.Cells[pInd, 6].value;
            double Row = IsPar ? pInd : worksheet.Cells[pInd, 7].value;
            return new сRequest() { Client = pClient, Column = Convert.ToInt32(Column), Row = Convert.ToInt32(Row), Request = Request, Sheet = Sheet };
        }

    }
}
