using System;
using System.Net;
using System.Net.Mail;

namespace TaskScheduler
{
    struct MailData
    {
        public string smtpServer;
        public string m_user;
        public string m_domain;
        public string from;
        public string m_pass;
        public string mailto;
        public string caption;
        public string message;
        public string attachFile;
    };
    
    class _Net
    {
        /// <summary>
        /// Отправка письма на почтовый ящик C# mail send
        /// </summary>
        /// <param name="smtpServer">Имя SMTP-сервера</param>
        /// <param name="m_user">Пользователь отправитель</param>
        /// <param name="m_domain">Домен отправителя</param>
        /// <param name="from">Адрес отправителя</param>
        /// <param name="m_pass">пароль к почтовому ящику отправителя</param>
        /// <param name="mailto">Адрес получателя</param>
        /// <param name="caption">Тема письма</param>
        /// <param name="message">Сообщение</param>
        /// <param name="attachFile">Присоединенный файл</param>
        public void SendMail(MailData m_maildata)/*string smtpServer, string m_user, string m_domain, string from, string m_pass,
        string mailto, string caption, string message, string attachFile = null)*/
        {
            
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(m_maildata.from);
                mail.To.Add(new MailAddress(m_maildata.mailto));
                mail.Subject = m_maildata.caption;
                mail.Body = m_maildata.message;
                if (!string.IsNullOrEmpty(m_maildata.attachFile))
                    mail.Attachments.Add(new Attachment(m_maildata.attachFile));
                SmtpClient client = new SmtpClient();
                client.Host = m_maildata.smtpServer;
                client.Port = 25;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential(m_maildata.m_user, m_maildata.m_pass, m_maildata.m_domain);//(from.Split('@')[0], pas_smtp);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;                
                client.Send(mail);
                mail.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Mail.Send: " + e.Message);
            }
        }
    }
}
