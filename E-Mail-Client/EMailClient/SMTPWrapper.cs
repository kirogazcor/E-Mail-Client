using System;
using System.Net.Mail;
using System.Net;

namespace E_Mail_Client.EMailClient
{
    // Оболочка класса Smtp-клиента
    class SMTPWrapper
    {
        SmtpClient Client;      // Smtp-клиент
        MailBox currentBox;     // Почтовый ящик, с которого отправляется почта

        public SMTPWrapper(MailBox box)
        {
            // Создание объекта SMTP-клиента с адресом сервера и портом
            Client = new SmtpClient(box.Settings.SmtpServer, box.Settings.SmtpPort)
            {
                Credentials = new NetworkCredential(box.Settings.UserName, box.Settings.Rassword),
                EnableSsl = box.Settings.Ssl
            };
            // Получение своего почтового ящика
            currentBox = box;
        }

        // Метод отправки письма
        public void SentMessage(object obj)
        {
            try
            {
                MailMessage message = (MailMessage)obj;
                // Добавление собственного адреса, как получателя скрытой копии письма
                message.Bcc.Add(currentBox.MyAddress);
                // Отправка письма                
                Client.Send(message);
                // Освобождение ресурсов
                Client.Dispose();
                // Наступление события завершения отправки письма
                OnSended(currentBox);
            }
            catch (Exception ex)
            {
                OnException(ex);
            }
        }

        #region События
        public event OnExceptionHandler OnException;
        public event OnSendedHandler OnSended;
        #endregion
    }
}
