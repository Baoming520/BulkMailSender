using BulkMailSender.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;

namespace BulkMailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            string[][] rows = null;
            var mappings = new Dictionary<string, int>();
            try
            {
                IExcelManager excelMgr = new OXExcelManager();
                var mailQueueFile = Path.Combine(ConfigurationManager.AppSettings["DataDir"], ConfigurationManager.AppSettings["MailConfigTemplate"]);
                string[] fields = excelMgr.ReadFields(mailQueueFile);
                if (fields == null)
                {
                    throw new Exception("邮件队列配置文件路径或名称不正确。");
                }
                
                for (int i = 0; i < fields.Length; i++)
                {
                    mappings.Add(fields[i], i);
                }
                rows = excelMgr.Read(mailQueueFile);
                var currDate = String.Format("{0}月{1}日", DateTime.Now.Month, DateTime.Now.Day);
                var currTime = String.Format("{0}点", DateTime.Now.Hour);
                for (int i = 0; i < rows.Length; i++)
                {
                    rows[i][mappings["Subject"]] = rows[i][mappings["Subject"]].Replace("{$DATE$}", currDate);
                    rows[i][mappings["Body"]] = rows[i][mappings["Body"]].Replace("{$DATETIME$}", currDate + currTime);
                }
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                var eMsg = String.Format("[{0}]邮件队列配置异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(eMsg);
                Console.ResetColor();
                return;
            }

            Console.WriteLine("\r\nMail To Address, Mail Cc Address, Status");
            foreach (var row in rows)
            {
                var mailTo = row[mappings["To"]];
                var mailCC = row[mappings["CC"]];
                var mailSubject = row[mappings["Subject"]];
                var mailBody = row[mappings["Body"]]; 
                var mailAttachments = row[mappings["AttachmentPathes"]];
                
                SendMail(mailTo, mailCC, mailAttachments, mailSubject, mailBody);
            }
        }

        private static void SendMail(string mailTo, string mailCc, string attachments, string subject, string body)
        {
            SendMailBase(
                ConfigurationManager.AppSettings["MailFrom"],
                ConfigurationManager.AppSettings["Password"],
                ConfigurationManager.AppSettings["DisplayName"],
                mailTo,
                mailCc,
                attachments,
                subject,
                body);
        }

        private static void SendMailBase(string mailFrom, string mailPwd, string displayName, string mailTo, string mailCc, string attachments, string subject, string body)
        {
            int loop = Convert.ToInt32(ConfigurationManager.AppSettings["Retry"]);
            var mailSender = new MailSender();
            mailSender.MailFromAddr = mailFrom;
            mailSender.MailFromPwd = mailPwd;
            mailSender.MailFromDisplayName = displayName;
            mailSender.MailTo = mailTo.ParseTextWithSemicolon();
            mailSender.MailCc = mailCc.ParseTextWithSemicolon();
            mailSender.MailSubject = subject;
            mailSender.MailHtmlBody = body;
            var attArr = attachments.ParseTextWithSemicolon();
            mailSender.SetLocalAttachments(attArr);

            bool succ = false;
            while (loop > 0)
            {
                succ = mailSender.Send(ConfigurationManager.AppSettings["SMTPServer"], Convert.ToInt32(ConfigurationManager.AppSettings["SMTPServerPort"]));
                if (succ)
                {
                    break;
                }

                Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["WaitingTime"]));
                loop--;
            }

            if (succ)
            {
                Console.WriteLine(String.Format("{0}, {1}, 已发送！", mailTo, mailCc));
            }
            else
            {
                Console.WriteLine(String.Format("{0}, {1}, 发送失败！", mailTo, mailCc));
            }
        }
    }
}
