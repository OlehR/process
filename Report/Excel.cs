using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Core; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "office"
using Utils;
using ExcelApp = Microsoft.Office.Interop.Excel; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "Microsoft.Office.Interop.Excel"

namespace Report
{
    public class DeleteSend
    {
        public string Pages { get; set; }
        public string eMails { get; set; }
    }
    public class Excel
    {
        bool Result = true;

        Mail Mail;
        MsSQL MsSQL = new MsSQL();
        string EmailError, EmailSuccess;
        int YearAgoDay, TwoYearAgoDay;

        DateTime FirstDayWeek(DateTime pDT)
        {
            int dd_d = (int)pDT.DayOfWeek;
            if (dd_d == 0)
                dd_d = 7;
            dd_d--;
            return pDT.AddDays(-dd_d);
        }

        public Excel() 
        {
            int Year = DateTime.Now.Year;
            DateTime d0 = new DateTime(Year, 1, 1);
            var fd0 = FirstDayWeek(d0);

            DateTime d1 = new DateTime(Year-1, 1, 1);
            var fd1 = FirstDayWeek(d1);

            DateTime d2 = new DateTime(Year-2, 1, 1);
            var fd2 = FirstDayWeek(d2);
            YearAgoDay = (int)(fd0 - fd1).TotalDays;
            TwoYearAgoDay = (int)(fd0 - fd2).TotalDays;

            var CurDir = AppDomain.CurrentDomain.BaseDirectory;
            var AppConfiguration = new ConfigurationBuilder() .SetBasePath(CurDir).AddJsonFile("appsettings.json").Build();
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
            
           // StringBuilder Success = new StringBuilder($"Start {DateTime.Now} {pSource}{Environment.NewLine}"), Error = new StringBuilder();
            try
            {
                // get the file attributes for file or directory
                FileAttributes attr = File.GetAttributes(pSource);

                if (attr.HasFlag(FileAttributes.Directory))
                {
                    if (CreateResultDirectory(pSource))
                        Files = Directory.GetFiles(pSource, "*.xls*");
                    else
                        Result = false;
                }
                else
                {
                    var receiptFilePath = Path.GetDirectoryName(pSource);
                    CreateResultDirectory(receiptFilePath);

                    Files = new string[] { pSource };
                    //MessageBox.Show("Its a file");
                }
                if (Files != null)
                {
                    foreach (var el in Files)
                        if(!el.StartsWith("~$"))
                        ExecuteExcelMacro(el);
                }
            }
            catch (Exception ex)
            {
                Result = false;
                FileLogger.WriteLogMessage($"ExecuteExcelsMacro Source=> {pSource} Error=> {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");
            }         
            
            //!TMP Якщо навчусь коректно завершувати ексель
            foreach (var process in Process.GetProcessesByName("excel"))
            {
                process.Kill();
                FileLogger.WriteLogMessage($"Kill  {process.ProcessName}{Environment.NewLine}");
            }

            if (!Result)
                Mail.SendMail(EmailSuccess, null, "Помилка формування звітів!!!", FileLogger.GetLog);
            else
                Mail.SendMail(EmailSuccess, null, "Звіти успішно зформовано", "Все ОК");

            FileLogger.WriteLogMessage($"End  {pSource}{Environment.NewLine}");
            return Result;
        }
        
        bool CreateResultDirectory(string pSourceDirectory)
        {
            bool Result = true;
            try
            {
                var Destination = Path.Combine(pSourceDirectory, "Result");
                if (!Directory.Exists(Destination))
                {
                    Directory.CreateDirectory(Destination);
                    FileLogger.WriteLogMessage($"CreateResultDirectory Create Directory {Destination}");
                }

                Destination = Path.Combine(Destination, "Arx");
                if (!Directory.Exists(Destination))
                {
                    Directory.CreateDirectory(Destination);
                    FileLogger.WriteLogMessage($"CreateResultDirectory Create Directory {Destination}");
                }

            }
            catch (Exception ex)
            {
                Result = false;
                FileLogger.WriteLogMessage($"CreateResultDirectory {pSourceDirectory} {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");               
               
            }
            return Result;
        }
        
        public void ExecuteExcelMacro(string pSourceFile)
        {
            FileLogger.WriteLogMessage($"File {pSourceFile}{Environment.NewLine}");

            ExcelApp.Application ExcelApp = null;
            ExcelApp.Workbook ExcelWorkBook = null;
            ExcelApp.Worksheet worksheet = null;

            List<cParameter> ResPar = null;
            bool Result = true, IsSendFile = true, IsMoveOldFile = false;
            DeleteSend DeleteSend=null;

            string DeletePage = null, HidePage = null, PathCopy = null;

            cArx Arx=null;

            try
            {
                сRequest ParRequest = null;
                List<сRequest> Requests = new List<сRequest>();
                string Macro = "Main", StartMacro = null;
                string Email = null, StartMacroEmail=null;
                bool? CalcReport = null;

                ExcelApp = new ExcelApp.Application();
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Visible = true;
                //ExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
                ExcelWorkBook = ExcelApp.Workbooks.Open(pSourceFile);

                worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets["config"];

                var range = worksheet.UsedRange;
                int rows_count = range.Rows.Count;

                for (int i = 1; i <= rows_count; i++)
                {
                    string str = worksheet.Cells[i, 1].value;
                    //string str = String.Format("[{0}] ", s);
                    if (str != null)
                    {
                        if (str.Equals("StartMacro"))
                        {
                            StartMacro = worksheet.Cells[i, 2].value;
                            StartMacroEmail = worksheet.Cells[i, 3].value;
                        }
                        else
                        if (str.Equals("Macro"))
                            Macro = worksheet.Cells[i, 2].value;
                        else
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
                        if (str.Equals("IsMoveOldFile"))
                            IsMoveOldFile = worksheet.Cells[i, 2].value;
                        else
                        if (str.Equals("IsSendFile"))
                            IsSendFile = worksheet.Cells[i, 2].value;
                        else
                        if (str.Equals("DeleteSend"))
                            DeleteSend = new DeleteSend() { Pages = worksheet.Cells[i, 2].value, eMails = worksheet.Cells[i, 3].value };
                        else
                        if (str.Equals("Arx"))
                            Arx = new cArx() { Row = i, Days = Convert.ToInt32(worksheet.Cells[i, 3].value), PathMove = worksheet.Cells[i, 4].value, EMail = worksheet.Cells[i, 5].value, DateFormatFile = worksheet.Cells[i, 6].value };
                        else
                        if (str.Equals("YearAgoDay"))
                            worksheet.Cells[i, 2].value = YearAgoDay;
                        else
                        if (str.Equals("TwoYearAgoDay"))
                            worksheet.Cells[i, 2].value = TwoYearAgoDay;
                        else
                        if (str.Equals("DayMonth"))
                        {
                            int DayMonth = Convert.ToInt32(worksheet.Cells[i, 2].value);
                            CalcReport = (CalcReport == true) || (DateTime.Now.Day == DayMonth);
                            if (DateTime.Now.Day == DayMonth)
                            {
                                string pEmail = worksheet.Cells[i, 3].value;
                                if (!string.IsNullOrEmpty(pEmail))
                                    Email = (string.IsNullOrEmpty(Email) ? "" : Email + ",") + pEmail;
                            }
                        }
                        else
                        if (str.Equals("DayWeek"))
                        {
                            int DayWeek = Convert.ToInt32(worksheet.Cells[i, 2].value);
                            CalcReport = (CalcReport == true) || ((int)DateTime.Now.DayOfWeek == DayWeek);
                            if ((int)DateTime.Now.DayOfWeek == DayWeek)
                            {
                                string pEmail = worksheet.Cells[i, 3].value;
                                if (!string.IsNullOrEmpty(pEmail))
                                    Email = (string.IsNullOrEmpty(Email) ? "" : Email + ",") + pEmail;
                            }
                        }

                    }
                }
                if (CalcReport == false)
                    return;

                //var r = pSourceFile.Split('.');
                var path = Path.Combine(Path.GetDirectoryName(pSourceFile), "Result");
                var FileName = Path.GetFileNameWithoutExtension(pSourceFile);
                var Extension = Path.GetExtension(pSourceFile);

                if (ParRequest != null)
                    ResPar = MsSQL.RunMsSQL(ParRequest).ToList();
                else
                    ResPar = new List<cParameter>() { new cParameter() { EMail = Email, Name = "" } };

                if (Arx != null && Arx.Days>=DateTime.Now.Day) 
                {
                    var ParArx = new List<cParameter>();
                    foreach(var el in ResPar)
                    {
                        var ArxEl = new cParameter(el);
                        if (!string.IsNullOrEmpty(Arx.EMail))
                            ArxEl.EMail = Arx.EMail;
                        if (!string.IsNullOrEmpty(Arx.PathMove))
                            ArxEl.PathMove = Arx.PathMove;
                        if (!string.IsNullOrEmpty(Arx.DateFormatFile))
                            ArxEl.DateFormatFile = Arx.DateFormatFile;
                        ArxEl.DateReport = Arx.FirstDayMonth;
                        ParArx.Add(ArxEl);
                    }
                    ResPar.AddRange(ParArx);
                }

                if(IsMoveOldFile)
                {
                    var ext = ParRequest==null?"":"_*";
                   // var MoveFileName = Path.Combine(path, $"{FileName}_????????{ext}{Extension}");
                    UtilDisc.MoveAllFilesMask(path, $"{FileName}_????????{ext}{Extension}",Path.Combine(path, "Arx"));
                }

                if (StartMacro != null)
                    ResPar.Add(new cParameter() { Macro = StartMacro, EMail = StartMacroEmail });

                ExcelWorkBook.Close();
                foreach (var el in ResPar)
                {
                    var elName = (string.IsNullOrEmpty(el.Name) ? "" : "_" + el.Name.Trim());
                    el.FileName = Path.Combine(path, $"{FileName}_{el.strDateReportFile}{elName}{Extension}");

                    ExcelWorkBook = ExcelApp.Workbooks.Open(pSourceFile);
                    if (File.Exists(el.FileName))
                        File.Delete(el.FileName);
                    ExcelWorkBook.SaveAs(el.FileName);
                    worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets["config"];

                    if (Arx != null)
                        worksheet.Cells[Arx.Row, 2].value = el.DateReport;
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
                            FileLogger.WriteLogMessage($"Start SQL = {r}{Environment.NewLine}");
                            MsSQL.Run(r);
                            FileLogger.WriteLogMessage($"End SQL = {r}{Environment.NewLine}");
                        }
                    }
                    FileLogger.WriteLogMessage($"Start Macro = {Macro} FileName=>{el.FileName} {Environment.NewLine}");
                    ExcelApp.Run(el.Macro??Macro);
                    FileLogger.WriteLogMessage($"End Macro = {Macro}{Environment.NewLine}");
                    
                    
                    
                    ExcelWorkBook.Save();

                    FileLogger.WriteLogMessage($"Save file {el.FileName}{Environment.NewLine}");
                    ExcelWorkBook.Close();
                }         
            }
            catch (Exception ex)
            {
               FileLogger.WriteLogMessage($"ExecuteExcelMacro SourceFile=> {pSourceFile} Error=> {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");                
            }
            finally
            {
                NAR(worksheet);
                worksheet = null;
                if (ExcelWorkBook != null)
                    ExcelWorkBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                NAR(ExcelWorkBook);
                ExcelWorkBook = null;
                
                
            }
            
            if (Result && ResPar != null)
            {
                if (!string.IsNullOrEmpty(DeletePage) || !string.IsNullOrEmpty(HidePage)||!string.IsNullOrEmpty(PathCopy))
                {
                    string[] DeletePages = null;
                    string[] HidePages = null;
                    if (!string.IsNullOrEmpty(DeletePage))
                        DeletePages = DeletePage.Split(',');
                    if (!string.IsNullOrEmpty(HidePage))
                        HidePages = HidePage.Split(',');
                    //Видаляємо сторінки
                    HideDeletePage(ExcelApp, ResPar, DeletePages, HidePages,ref Result,PathCopy);
                }
                //Відправляємо Листи
                SendMail(ResPar, IsSendFile);

                //Готуємо і відправляємо скорочену версію звіта.
                if (DeleteSend != null && !string.IsNullOrEmpty(DeleteSend.Pages) && ResPar.Count == 1)
                {
                    var DeletePages = DeleteSend.Pages.Split(',');
                    HideDeletePage(ExcelApp, ResPar, DeletePages, null, ref Result, PathCopy, true);
                    SendMail(ResPar, IsSendFile);
                }
            }




            // Закриваємо ексель.
            //NAR(ExcelWorkbooks);
            if (ExcelApp != null)
            {
                ExcelApp.Application.Quit();
                ExcelApp.Quit();
                NAR(ExcelApp);
                ExcelApp = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            

        }

        private void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }

        private void SendMail(IEnumerable<cParameter> ResPar, bool pIsSendFile)
        {
            //Відправляємо Листи
            try
            {
                foreach (var el in ResPar)
                {
                  Result &= Mail.SendMail(el.EMail, !pIsSendFile && !string.IsNullOrEmpty(el.CopyFileName)?null: el.FileName, (!pIsSendFile  && !string.IsNullOrEmpty(el.CopyFileName)?$" Звіт Збережено {el.CopyFileName}" : null), null);
                }
            }
            catch (Exception ex)
            {
               FileLogger.WriteLogMessage($"SendMail {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");
            }
        }
       
        private void HideDeletePage(ExcelApp.Application ExcelApp, IEnumerable<cParameter> ResPar, string[] DeletePages, string[] HidePages, ref bool Result, string pPathCopy=null,bool IsShort=false)
        {
            ExcelApp.Workbook ExcelWorkBook;
            foreach (var el in ResPar)
            {

                ExcelWorkBook = ExcelApp.Workbooks.Open(el.FileName);
                if (IsShort)
                {
                   // var s = Path.GetDirectoryName(el.FileName);

                    string FN = Path.Combine(Path.GetDirectoryName(el.FileName), Path.GetFileNameWithoutExtension(el.FileName) + "_short" + Path.GetExtension(el.FileName));
                    el.FileName = FN;
                }
                if (DeletePages != null)
                    foreach (var page in DeletePages)
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(100);
                            ExcelApp.Worksheet worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets[page];
                            worksheet.Delete();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                            worksheet = null;
                        }
                        catch (Exception ex)
                        {
                            Result = false;
                            FileLogger.WriteLogMessage($"Delete Page={page} {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");
                        }
                    }
                //Ховаєм сторінки

                if (HidePages != null)
                    foreach (var page in HidePages)
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(100);
                            ExcelApp.Worksheet worksheet = (ExcelApp.Worksheet)ExcelWorkBook.Worksheets[page];
                            worksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                            worksheet = null;

                        }
                        catch (Exception ex)
                        {
                            Result = false;
                            FileLogger.WriteLogMessage($"Hide Page={page} {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");
                            
                            
                        }

                    }
                try
                {
                    ExcelWorkBook.SaveAs(el.FileName);
                }
                catch (Exception ex)
                {
                    Result = false;
                    FileLogger.WriteLogMessage($"Save={el.FileName} {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");
                }
                if (!string.IsNullOrEmpty(pPathCopy))
                {
                    string NewFile = Path.Combine(pPathCopy, Path.GetFileName(el.FileName));
                    try
                    {
                        if (File.Exists(NewFile))
                            File.Delete(NewFile);
                        ExcelWorkBook.SaveAs(NewFile);
                        el.CopyFileName = NewFile;
                    }
                    catch (Exception ex)
                    {
                        Result = false;
                        FileLogger.WriteLogMessage($"SaveCopy={NewFile} {ex.Message}{Environment.NewLine}{Environment.StackTrace}{Environment.NewLine}");                        
                    }
                }
                if (ExcelWorkBook != null)
                    ExcelWorkBook.Close(false);
                if (ExcelWorkBook != null) 
                { 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook);
                    ExcelWorkBook = null;
                }
            }
        }

       /* private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
        }*/

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
