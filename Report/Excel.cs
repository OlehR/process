using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Core; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "office"
//using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "Microsoft.Office.Interop.Excel"

namespace Report
{
    public class Excel
    {       

        Mail Mail;
        MsSQL MsSQL = new MsSQL();
        string EmailError, EmailSuccess;
        public Excel() 
        {
            var CurDir = AppDomain.CurrentDomain.BaseDirectory;
            var AppConfiguration = new ConfigurationBuilder()  
                .SetBasePath(CurDir).AddJsonFile("appsettings.json").Build();
            MailConfig MailConfig=new MailConfig();
            MailConfig.SmtpServer = AppConfiguration.GetSection("Report:Mail:SmtpServer").Value;
            MailConfig.From = AppConfiguration.GetSection("Report:Mail:From").Value;
            MailConfig.Login = AppConfiguration.GetSection("Report:Mail:Login").Value;
            MailConfig.Password = AppConfiguration.GetSection("Report:Mail:Password").Value;
            Mail = new Mail(MailConfig);
            EmailError = AppConfiguration.GetSection("Report:EmailError").Value;
            EmailSuccess = AppConfiguration.GetSection("Report:EmailSuccess").Value;
        }
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
            {
                Console.WriteLine(Error.ToString());                
                Mail.SendMail(EmailSuccess, null , "Помилка формування звітів!!!", Error.ToString());
            }
            else
                Mail.SendMail(EmailSuccess, null, "Звіти успішно зформовано", "Все ОК");

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
        public void ExecuteExcelMacro(string pSourceFile, StringBuilder pSuccess, StringBuilder pError)
        {
            pSuccess.Append($"{DateTime.Now} File {pSourceFile}{Environment.NewLine}");

            ExcelApp.Application ExcelApp = null;
            ExcelApp.Workbook ExcelWorkBook = null;
            IEnumerable<cParameter> ResPar = null;
            bool Result = true, IsSendFile = true;

            string DeletePage = null, HidePage = null, PathCopy = null, MoveOldFile = null;


            try
            {
                сRequest ParRequest = null;
                List<сRequest> Requests = new List<сRequest>();
                string Macro = "Main", StartMacro = null;
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
                            ParRequest = GetRequest(worksheet, ExcelWorkBook, i, eClient.MsSql, true);
                        else
                          if (str.Equals("pMDX"))
                            ParRequest = GetRequest(worksheet, ExcelWorkBook, i, eClient.MDX, true);
                        else
                        if (str.Equals("SQL"))
                            Requests.Add(GetRequest(worksheet, ExcelWorkBook, i, eClient.MsSql));
                        else
                          if (str.Equals("MDX"))
                            Requests.Add(GetRequest(worksheet, ExcelWorkBook, i, eClient.MDX));
                        else
                        if (str.Equals("DeletePage"))
                            DeletePage = worksheet.Cells[i, 2].value;
                        else
                        if (str.Equals("HidePage"))
                            HidePage = worksheet.Cells[i, 2].value;
                        else
                        if (str.Equals("PathCopy"))
                            PathCopy = worksheet.Cells[i, 2].value;
                        else
                        if (str.Equals("MoveOldFile"))
                            MoveOldFile = worksheet.Cells[i, 2].value;
                        if (str.Equals("IsSendFile"))
                            IsSendFile = "true".Equals(worksheet.Cells[i, 2].value);
                    }
                }

                //var r = pSourceFile.Split('.');
                var path = Path.Combine(Path.GetDirectoryName(pSourceFile), "Result");
                var FileName = Path.GetFileNameWithoutExtension(pSourceFile);
                var Extension = Path.GetExtension(pSourceFile);

                if (ParRequest != null)
                    ResPar = MsSQL.RunMsSQL(ParRequest);
                else
                    ResPar = new List<cParameter>() { new cParameter() { EMail = Email, Name = "" } };

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
                    foreach (var r in Requests)
                    {
                        if (r.Client == eClient.MsSql)
                        {
                            pSuccess.Append($"{DateTime.Now} Start SQL = {r}{Environment.NewLine}");
                            MsSQL.Run(r);
                            pSuccess.Append($"{DateTime.Now} End SQL = {r}{Environment.NewLine}");
                        }

                    }
                    pSuccess.Append($"{DateTime.Now} Start Macro = {Macro}{Environment.NewLine}");
                    ExcelApp.Run(Macro);
                    pSuccess.Append($"{DateTime.Now} End Macro = {Macro}{Environment.NewLine}");
                    
                    el.FileName = Path.Combine(path, FileName + "_" + DateTime.Now.ToString("yyyyMMdd")+ (string.IsNullOrEmpty(el.Name)?"":"_"+ el.Name.Trim()) + Extension);
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
                if (ExcelWorkBook != null)
                    ExcelWorkBook.Close(false);

                if (ExcelWorkBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook); }

            }
            
            if (Result && ResPar != null)
            {

                if (!string.IsNullOrEmpty(DeletePage) || !string.IsNullOrEmpty(HidePage))
                {
                    string[] DeletePages = null;
                    string[] HidePages = null;
                    if (!string.IsNullOrEmpty(DeletePage))
                        DeletePages = DeletePage.Split(',');
                    if (!string.IsNullOrEmpty(HidePage))
                        HidePages = HidePage.Split(',');
                    //Видаляємо сторінки
                    foreach (var el in ResPar)
                    {
                        ExcelWorkBook = ExcelApp.Workbooks.Open(el.FileName);
                        if (DeletePages != null)
                            foreach (var page in DeletePages)
                            {
                                try
                                {
                                    ExcelApp.Worksheet worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets[page];
                                    worksheet.Delete();
                                }
                                catch (Exception ex)
                                {
                                    Result = false;
                                    pError.Append($"Page={page} " + ex.Message + Environment.NewLine);
                                    pError.Append(Environment.StackTrace + Environment.NewLine);
                                }
                            }
                        //Ховаєм сторінки
                        if (HidePages != null)
                            foreach (var page in HidePages)
                            {
                                try
                                {
                                    ExcelApp.Worksheet worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets[page];
                                    worksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                                }
                                catch (Exception ex)
                                {
                                    Result = false;
                                    pError.Append($"Page={page} " + ex.Message + Environment.NewLine);
                                    pError.Append(Environment.StackTrace + Environment.NewLine);
                                }

                            }
                        ExcelWorkBook.SaveAs(el.FileName);
                        if (ExcelWorkBook != null)
                            ExcelWorkBook.Close(false);

                    }


                }
                //Відправляємо Листи
                try
                {
                    foreach (var el in ResPar)
                    {
                        //var emails = el.EMail.Split(',');
                        //foreach (var email in emails)
                            Mail.SendMail(el.EMail, el.FileName, null, null, pSuccess, pError);
                    }
                }
                catch (Exception ex)
                {
                    pError.Append(ex.Message + Environment.NewLine);
                    pError.Append(Environment.StackTrace + Environment.NewLine);
                }
            }
            if (ExcelApp != null)
                ExcelApp.Quit();
            if (ExcelApp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp); }

        }

        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
        }

        static private сRequest GetRequest(ExcelApp.Worksheet worksheet, ExcelApp.Workbook pExcelWorkBook, int pInd, eClient pClient = eClient.NotDefine, bool IsPar = false)
        {            
            string Request = worksheet.Cells[pInd, 4].value;
            string Sheet = IsPar ? "config" : worksheet.Cells[pInd, 5].value;
            ExcelApp.Worksheet Worksheet = (ExcelApp.Worksheet) pExcelWorkBook.Worksheets[Sheet];
            double Column = IsPar ? 5 : worksheet.Cells[pInd, 6].value;
            double Row = IsPar ? pInd : worksheet.Cells[pInd, 7].value;
            return new сRequest() { Client = pClient, Column = Convert.ToInt32(Column), Row = Convert.ToInt32(Row), Request = Request, Sheet = Worksheet };
        }

    }
}
