using System;

namespace E_Mail_Client.EMailClient
{
    class IMAPClient : TCPClient
    {
        // Подключение IMAP-клиента в конструкторе
        public IMAPClient(string host, int port, bool ssl) : base(host, port, ssl)
        {
            Read("OK");
        }

        private int lastIndex = 0;              // Последний индекс команды           

        // Команда закрытия соединения
        public void Logout()
        {
            if (Connected)
            {
                string _cmd = "LOGOUT";
                SRComand(_cmd);
            }
        }

        // Команда аутентификации по логину и паролю
        public string Login(string user, string pass)
        {
            string answer = "";
            if (Connected)
            {
                string _cmd = "LOGIN " + user + " " + pass; ;
                answer = SRComand(_cmd);
            }
            else throw new Exception("Нет соединения");
            return answer;
        }

        // Команда загрузки списка папок
        public void List()
        {
            if (Connected)
            {
                string _cmd = "LIST " + "\"\"" + " \"*\"";
                string answer = SRComand(_cmd);
                if (answer.Contains(" OK "))
                {
                    // Получение списка папок из строки ответа сервера
                    // и наступление события завершения выполнения команды LIST
                    ListCompleted(ParserMessage.GetFolders(answer));
                }
                else throw new Exception("Ошибка загрузки списка почтовых ящиков");
            }
            else throw new Exception("Нет соединения");
        }        

        // Команда выбора папки
        public string SelectFolder(string path)
        {
            string answer = "";
            if (Connected)
            {
                string _cmd = "SELECT " + path;
                answer = SRComand(_cmd);
            }
            else throw new Exception("Нет соединения");
            return answer;
        }

        // Команда копирования писем в другую папку
        public string Copy(string path, int first, int last)
        {
            string answer = "";
            if (Connected)
            {
                string _cmd = "COPY " + first.ToString() + ":" + last.ToString() + " " + path;
                answer = SRComand(_cmd);
            }
            else throw new Exception("Нет соединения");
            return answer;
        }

        // Команда загрузки заголовков всех писем,
        // начиная с первого
        public void LoadTitleMessages(int num)
        {
            if (Connected)
            {
                string _cmd = "FETCH 1:" + num.ToString() + " BODY[HEADER]";
                string answer = SRComand(_cmd);
                if (answer.Contains(" OK "))
                {
                    // Наступление события завершения выполнения команды LoadTitleMessages
                    LoadTitleCompleted(answer);
                }
                else throw new Exception("Ошибка загрузки списка писем");
            }
            else throw new Exception("Нет соединения");
        }
        
        // Команда получения количества писем в папке
        public int MessInFolder(string path)
        {
            int num = 0;
            if (Connected)
            {
                string _cmd = "STATUS " + path + " (messages)";
                string answer = SRComand(_cmd);
                if (answer.Contains("OK")) num = ParserMessage.GetNumberMess(answer);
                else throw new Exception("Ошибка загрузки количества писем в почтовом ящике");
            }
            else throw new Exception("Нет соединения");
            return num;
        }

        // Команда загрузки тела письма
        public void LoadMessage(string num)
        {
            if (Connected)
            {
                string _cmd = "FETCH " + num + " BODY[TEXT]";
                string answer = SRComand(_cmd);
                if (answer.Contains(" OK "))
                {
                    // Получение тела письма из строки ответа сервера
                    // и наступление события завершения выполнения команды LoadMessage
                    LoadMessageCompleted(ParserMessage.GetBody(answer));
                }
                else throw new Exception("Ошибка загрузки письма");
            }
            else throw new Exception("Нет соединения");
        }

        // Команда удаления письма
        public bool DeleteMessage(int num)
        {            
            if (Connected)
            {
                string _cmd = "STORE " + num.ToString() + " +FLAGS (\\Deleted)";
                // Отправка команды добавления флага к письму
                string answer = SRComand(_cmd);
                if (answer.Contains(" OK "))
                {
                    _cmd = "EXPUNGE";
                    // Отправка команды на применение флага удаления
                    answer = SRComand(_cmd);
                    // Выдача подтверждения об успешном выполнении удаления письма
                    return answer.Contains(" OK ");                    
                }
                else throw new Exception("Ошибка загрузки письма");
            }
            else throw new Exception("Нет соединения");
        }

        // Метод отправляет запрос на IMAP-сервер и получает с него ответ
        private string SRComand(string cmd)
        {
            if (busyStream) throw new Exception("Поток занят");
            else
            {
                string response = null;
                busyStream = true;
                // Увеличение счетчика команд
                lastIndex++;
                // Ключевое слово для окончания чтения
                string key = "KiR35" + lastIndex.ToString();
                string _cmd = key + " " + cmd + "\r\n";
                Write(_cmd);
                do
                {
                    // Получение ответа и удаление из него лишних символов
                    response += ParserMessage.RemChar(Read(key), "\0");
                } while (!(response.Contains(key)));
                busyStream = false;
                // Ограничение ответа до ключевого слова
                return response.Substring(0, response.IndexOf(key) + key.Length + 4);
            }
        }
        #region События
        public event ListCompleteHandler ListCompleted;
        public event LoadCompleteHandler LoadTitleCompleted;
        public event LoadCompleteHandler LoadMessageCompleted;
        #endregion
    }
}
