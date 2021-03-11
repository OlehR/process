using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace Report
{
    public class MailConfig
    { 
        public string SmtpServer { get; set; }
        public string From { get; set; }        
        public string Login { get; set; }
        public string Password { get; set; }        

    }
    
public class Mail
    {
        

        MailConfig Config = new MailConfig();
        
        public Mail(MailConfig pMailConfig)
        {
            Config = pMailConfig;
        }
        
        public bool SendMail(string pTo, string pFile, string pSubject = null, string pBody = null, StringBuilder pSuccess = null, StringBuilder pError = null)
        {
            try
            {
                SmtpClient Smtp = new SmtpClient(Config.SmtpServer, 25);
                Smtp.Credentials = new NetworkCredential(Config.Login, Config.Password);
                MailMessage Message = new MailMessage();
                Message.From = new MailAddress(Config.From);
                
                var emails = pTo.Split(',');
                foreach (var email in emails)
                    Message.To.Add(new MailAddress(email));

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
