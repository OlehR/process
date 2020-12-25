
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
/*
using Limilabs.Client.SMTP;
using Limilabs.Mail;
using Limilabs.Mail.Fluent;
*/
namespace Report
{
    public class Mail
    {
        /*       public void SendMail(string pFile, string pTo= "o.rutkovskyj@vopak.uz.ua", StringBuilder pSuccess = null, string pFrom = "reports@vopak.uz.ua",string pPassWord= "tOeD23LCA")
               {
                   IMail email = Limilabs.Mail.Fluent.Mail
           .Html(@"Html with an image: <img src=""cid:lena"" />")
           //.AddVisual(@"c:\lena.jpeg").SetContentId("lena")
           .AddAttachment(pFile) //.SetFileName("document.doc")
           .To(pTo)
           .From(pFrom)
           .Subject("Звіт")
           .Create();
                   using (Smtp smtp = new Smtp())
                   {
                       smtp.Connect("mail.vopak.uz.ua", 25);  // or ConnectSSL for SSL
                       smtp.UseBestLogin(pFrom, pPassWord);//("order@vopak.uz.ua", "JQfzuQtCD");
                       smtp.SendMessage(email);
                       smtp.Close();
                       if (pSuccess != null)
                           pSuccess.Append($"Send Email to {pTo} file {pFile}{Environment.NewLine}");
                   }
               }
       */
        public bool SendMail(string pTo, string pFile, string pSubject = null, string pBody = null, StringBuilder pSuccess = null, StringBuilder pError = null,
                                    string pSmtpServer = "mail.vopak.uz.ua", string pFrom = "reports@vopak.uz.ua", string pLogin = "reports@vopak.uz.ua", string pPassword = "tOeD23LCA")
        {
            try
            {
                SmtpClient Smtp = new SmtpClient(pSmtpServer, 25);
                Smtp.Credentials = new NetworkCredential(pLogin, pPassword);
                MailMessage Message = new MailMessage();
                Message.From = new MailAddress(pFrom);
                Message.To.Add(new MailAddress(pTo));
                Message.Subject = (pSubject == null ? "Send: " + pFile : pSubject);
                Message.Body = (pBody == null ? "Send: " + pFile : pBody);

                if (pFile != null)
                {
                    Attachment attach = new Attachment(pFile, MediaTypeNames.Application.Octet);

                    ContentDisposition disposition = attach.ContentDisposition;
                    disposition.CreationDate = System.IO.File.GetCreationTime(pFile);
                    disposition.ModificationDate = System.IO.File.GetLastWriteTime(pFile);
                    disposition.ReadDate = System.IO.File.GetLastAccessTime(pFile);
                    Message.Attachments.Add(attach);
                }
                Smtp.Send(Message);//отправка
                if (pSuccess != null)
                    pSuccess.Append($"Send Email to {pTo} file {pFile}{Environment.NewLine}");
                return true;
            }
            catch (Exception ex)
            {
                pError.Append(ex.Message + Environment.NewLine);
                pError.Append(Environment.StackTrace + Environment.NewLine);
               
                return false;
            }
        }
    }
}
