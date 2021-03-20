using System;
using System.Collections.Generic;
using System.Linq;

namespace E_Mail_Client.EMailClient
{   
    // Оболочка класса IMAP-клиента
    public class IMAPWrapper
    {
        private IMAPClient Client;          // Imap-клиент
        private bool IsConnected = false;   // Флаг подключения к серверу

        public void DragToSent(MailBox box)
        {
            try
            {
                string server = box.Settings.ImapServer;
                int port = box.Settings.ImapPort;
                bool ssl = box.Settings.Ssl;
                string user = box.Settings.UserName;
                string pass = box.Settings.Rassword;
                // Создание соединения с сервером
                Client = new IMAPClient(server, port, ssl);
                // Аутентификация на сервере
                if (Client.Login(user, pass).Contains("OK"))
                {
                    IsConnected = true;
                    int num = 0;
                    if((num = Client.MessInFolder("INBOX")) > 0)
                        if (Client.SelectFolder("INBOX").Contains("OK"))
                        {
                            string pointer = box.Folders.Where(x => x.Type == TYPE_FOLDER.SENT).First().Pointer;
                            if (Client.Copy(pointer, num, num).Contains("OK"))
                            {
                                if (!Client.DeleteMessage(num))
                                    throw new Exception("Ошибка удаления письма. Удалите отправленное письмо из ящика INBOX");
                            }
                            else throw new Exception("Ошибка копирования. Переместите отправленное письмо из ящика INBOX");
                        }
                }
                else throw new Exception("Ошибка подключения. Отправленное письмо сохранено в папке INBOX");                
            }
            catch (Exception ex)
            {
                // Наступление события ошибки
                OnException(ex);
            }
        }

        #region Подключение к серверу
        // Подключение к IMAP-серверу в паралельном потоке
        public void Connect(object obj)
        {
            MailBox box = (MailBox)obj;
            Connect(box);
        }

        // Подключение к IMAP-серверу с аргументом почтового ящика
        public void Connect(MailBox box)
        {
            Connect(box.MyAddress.Address, box.Settings.ImapServer, box.Settings.ImapPort, box.Settings.UserName, box.Settings.Rassword, box.Settings.Ssl);
        }

        // Подключение к IMAP-серверу c параметрами
        private void Connect(string adress, string server, int port, string user, string pass, bool ssl)
        {
            try
            {
                // Создание соединения с сервером
                Client = new IMAPClient(server, port, ssl);
                // Аутентификация на сервере
                if (Client.Login(user, pass).Contains("OK"))
                {
                    Client.ListCompleted += Client_ListCompleted;
                    // Загрузка списка папок
                    Client.List();
                }
                else throw new Exception("Ошибка подключения");
                IsConnected = true;
            }
            catch (Exception ex)
            {
                // Наступление события ошибки
                OnException(ex);
                // Загрузка списка папок из файла
                List<Folder> folders = MailStorage.LoadFolders(adress);
                // Наступление события подключен
                OnConnected(folders);
            }
            
        }

        // Обработка завершения выполнения метода List
        private void Client_ListCompleted(List<Folder> folders)
        {
            // Наступает событие подключен
            OnConnected(folders);
        }
        #endregion

        #region Загрузка списка писем 
        // Загрузка списка писем в параллельном потоке        
        public void LoadMessages(object obj)
        {
            int num = 0;
            string path = (string)obj;
            try
            {
                if (!IsConnected) throw new Exception("Не подключен к серверу");
                // Получение количества писем в папке
                if ((num = Client.MessInFolder(path)) > 0)
                {
                    // Выбор папки
                    if (Client.SelectFolder(path).Contains("OK"))
                    {
                        Client.LoadTitleCompleted += Client_LoadTitleCompleted;
                        // Загрузка заголовков писем
                        Client.LoadTitleMessages(num);
                    }
                    else throw new Exception("Ошибка выбора ящика");
                }
                else OnLoadedMessages(new List<MyMailMessage>());
            }
            catch (Exception ex)
            {
                // Наступление события ошибки
                OnException(ex);
                // Наступление события список писем не загружен с сервера
                OnNotLoadMessages();
            }           
        }
        // Обработка завершения выполнения метода LoadTitleMessages
        private void Client_LoadTitleCompleted(string answer)
        {
            // Создание массива конвертов писем
            string[] headers = answer.Split('*');
            string header = null;

            List<MyMailMessage> messList = new List<MyMailMessage>();
            if (headers.Length > 0)
            {
                for (int i = 1; i < headers.Length; i++)
                {
                    if (headers[i].Contains("\r\n\r\n"))
                    {
                        header += "*" + headers[i].Substring(0, headers[i].IndexOf("\r\n\r\n"));
                        try
                        {
                            // Добавление нового письма в список
                            if (header.Contains("\r\nMIME-Version:"))
                                messList.Add(new MyMailMessage(header));
                        }
                        catch(Exception ex)
                        {
                            OnException(ex);
                        }
                        header = null;
                    }
                    // Приклеивание окончания к конверту с символом '*'
                    else header += "*" + headers[i];
                }
            }
            // Сортировка писем по номеру
            List<MyMailMessage> sorted = messList.OrderByDescending(x => x.Num).ToList();
            // Наступление события завершения загрузки списка писем
            OnLoadedMessages(sorted);
        }
        #endregion

        #region Загрузка письма
        // Метод загрузки письма
        public void LoadMessage(object obj)
        {            
            string num = (string)obj.ToString();
            try
            {
                if (!IsConnected) throw new Exception("Не подключен к серверу");
                Client.LoadMessageCompleted += Client_LoadMessageCompleted;
                Client.LoadMessage(num);
            }
            catch (Exception ex)
            {
                // Наступление события ошибки
                OnException(ex);
                // Наступление событие окончания загрузки письма
                // с параметром того, что с сервера загрузки не было
                OnLoadedMessage("Is_not_loaded_message");
            }
            
        }
        // Обработка завершения загрузки сообщения
        private void Client_LoadMessageCompleted(string answer)
        {
            OnLoadedMessage(answer);
        }
        #endregion

        // Удаление письма
        public bool DeleteMessage(int _num)
        {
            try
            {
                if (!IsConnected) throw new Exception("Не подключен к серверу");
                return Client.DeleteMessage(_num);
            }
            catch (Exception ex)
            {
                // Наступление события ошибки
                OnException(ex);
                return false;
            }
        }
        
        // Отключение от IMAP-сервера        
        public void Disconnect()
        {
            try
            {
                if (IsConnected)
                {
                    Client.Logout();
                    Client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                OnException(ex);
            }
            finally
            {
                IsConnected = false;
            }
        }

        #region События
        public event ListCompleteHandler OnConnected;
        public event LoadListMessHandler OnLoadedMessages;
        public event OnExceptionHandler OnException;
        public event LoadCompleteHandler OnLoadedMessage;
        public event ViewHandler OnNotLoadMessages;
        #endregion
    }
}
