using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using System.Net.Mail;


namespace Word_test
{
    internal class EmailService
    {
        
        public void SendEmail(string email, string subject, string message)
        {
            var emailMessage = new MimeMessage() ;

            emailMessage.From.Add(new MailboxAddress("Тест Ворда", "VanekTestovik1@yandex.ru"));
            emailMessage.To.Add(new MailboxAddress("", email));
            emailMessage.Subject = subject;
            emailMessage.Body = new TextPart("Plain")
            {
                Text = message
            };
            

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.Connect("smtp.yandex.ru", 465, true);
                client.Authenticate("VanekTestovik1@yandex.ru", "VanekTestovik123");
                client.Send(emailMessage);

                client.Disconnect(true);
            }
        }
    }
}
