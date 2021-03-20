using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;

namespace E_Mail_Client.EMailClient
{
    public delegate void ListCompleteHandler(List<Folder> folders);
    public delegate void OnExceptionHandler(Exception ex);
    public delegate void LoadListMessHandler(List<MyMailMessage> messages);
    public delegate void LoadCompleteHandler(string answer);
    public delegate void ViewMessageHandler(MyMailMessage message);
    public delegate void OnSendedHandler(MailBox box);
    public delegate void ViewHandler();

    // Класс почтового клиента
    public class EmailClient
    {
        private IMAPWrapper inputMail;      // Клиент входящей почты
        private List <MailBox> mailBoxList; // Список почтовых ящиков
        private MailBox currentMailBox;     // Текущий почтовый ящик  
        private SMTPWrapper outputMail;     // Клиент исходящей почты

        #region Публичные свойства

        public List<MailBox> MailBoxList
        {
            set { mailBoxList = value; }
            get { return mailBoxList; }
        }

        public MailBox CurrentMailBox
        {
            set { currentMailBox = value; }
            get { return currentMailBox; }
        }

        #endregion

        // Конструктор почтового клиента
        public EmailClient()
        {
            mailBoxList = new List<MailBox>();
            inputMail = new IMAPWrapper();
            inputMail.OnConnected += InputMail_OnConnected;
            inputMail.OnException += _OnEsception;
            inputMail.OnLoadedMessages += InputMail_OnLoadedMessages;
            inputMail.OnLoadedMessage += InputMail_OnLoadedMessage;
            inputMail.OnNotLoadMessages += InputMail_OnNotLoadMessages;
        }        

        #region Подключение к серверу входящей почты и загрузка списка папок
        // Метод загрузки списка папок
        public void LoadFolderList()
        {

            MailBox current = new MailBox(currentMailBox);
            // Запуск подключения к серверу входящей почты в параллельном потоке
            Thread inConnct = new Thread(new ParameterizedThreadStart(inputMail.Connect));
            inConnct.Start(current);

        }

        // Обработка завершения подключения к серверу входящей почты
        private void InputMail_OnConnected(List<Folder> folders)
        {
            if (currentMailBox != null)
            {
                currentMailBox.Folders = folders;
                // Сохранение списка папок в файл
                MailStorage.SaveFolders(currentMailBox);
            }
        }
        #endregion

        #region Загрузка списка писем из текущей папки
        // Метод загрузки списка писем
        public void LoadMessageList()
        {

            if (currentMailBox.SelectedFolder.Type == TYPE_FOLDER.DRAFTS || currentMailBox.SelectedFolder.Type == TYPE_FOLDER.TEMPLATE)
            {
                // Для папок черновиков и шаблонов
                // загружать список писем из файла
                List<MyMailMessage> messages = MailStorage.LoadHeaders(currentMailBox);
                InputMail_OnLoadedMessages(messages);
            }
            else
            {
                // Запуск загрузки списка писем из папки в параллельном потоке
                Thread inSlt = new Thread(new ParameterizedThreadStart(inputMail.LoadMessages));
                inSlt.Start(currentMailBox.SelectedFolder.Pointer);
            }
        }

        // Обработка завершения загрузки списка писем
        private void InputMail_OnLoadedMessages(List<MyMailMessage> messanges)
        {
            if (currentMailBox != null)
                if (currentMailBox.SelectedFolder != null)
                {
                    currentMailBox.SelectedFolder.Messages = new List<MyMailMessage>(messanges);                    
                    // Сохранение заголовков писем в файл
                    MailStorage.SaveHeaders(currentMailBox);
                }
            // Наступает событие готовности к отображению списка писем
            OnViewListMess();
        }

        // Обработка того, что список писем не загружен с сервера
        private void InputMail_OnNotLoadMessages()
        {
            // Загрузка списка писем из файла
            List<MyMailMessage> messages = MailStorage.LoadHeaders(currentMailBox);
            // Сортировка списка писем по убыванию даты
            List<MyMailMessage> sorted = messages.OrderByDescending(x => x.Date).ToList();
            // Наступление события завершения загрузки списка писем
            InputMail_OnLoadedMessages(sorted);
        }
        #endregion
        
        #region Загрузка выбранного письма
        // Метод загрузки письма
        public void LoadMessage()
        {
            // Запуск загрузки письма в параллельном потоке
            Thread inFtch = new Thread(new ParameterizedThreadStart(inputMail.LoadMessage));
            inFtch.Start(currentMailBox.SelectedFolder.Message.Num);
        }
        // Обработка загрузки письма
        private void InputMail_OnLoadedMessage(string answer)
        {
            if (currentMailBox != null)
                if (currentMailBox.SelectedFolder != null)
                    if (currentMailBox.SelectedFolder.Message != null)
                    {
                        // Выбор адреса получателя, который будет отображаться первым
                        MailAddress firstAddress = null;
                        if (currentMailBox.SelectedFolder.Message.To.Where
                            (e => e.Address.Equals(currentMailBox.MyAddress.Address, 
                            StringComparison.OrdinalIgnoreCase)).Count() > 0)
                            firstAddress = currentMailBox.SelectedFolder.Message.To.Where
                                (e => e.Address.Equals(currentMailBox.MyAddress.Address, 
                                StringComparison.OrdinalIgnoreCase)).First();
                        if (firstAddress != null)
                        {
                            // Перемещение выбранного адреса получателя в начало списка
                            currentMailBox.SelectedFolder.Message.To.Remove(firstAddress);
                            currentMailBox.SelectedFolder.Message.To.Insert(0, firstAddress);
                        }
                        if (answer == "Is_not_loaded_message")
                            // Если загрузить тело письма с сервера не удалось
                            // получить тело письма из файла
                            currentMailBox.SelectedFolder.Message.AddBody(currentMailBox);
                        else // Иначе получить тело письма из ответа сервера
                            currentMailBox.SelectedFolder.Message.AddBody(answer);
                        // Сохранение тела письма в файл
                        MailStorage.SaveMessage(currentMailBox);
                        foreach (MyAttachment attach in currentMailBox.SelectedFolder.Message.MyAttachments)
                            // Очистка памяти от вложений письма
                            attach.Clear();
                    }
            // Наступление события готовности письма к отображению
            OnViewMess();
        }
        #endregion

        #region Удаление выбранного письма
        // Метод удаления письма
        public void DeleteMessage()
        {
            // Если удаление на сервере выполнено успешно перезагрузить список писем
            if (inputMail.DeleteMessage(currentMailBox.SelectedFolder.Message.Num)) LoadMessageList();            
        }
        #endregion

        #region Отправка письма
        public void SentMessage(MailMessage message)
        {
            outputMail = new SMTPWrapper(currentMailBox);
            outputMail.OnException += _OnEsception;
            outputMail.OnSended += OutputMail_OnSended;
            // Запуск подключения к серверу входящей почты в параллельном потоке
            Thread inSent = new Thread(new ParameterizedThreadStart(outputMail.SentMessage));
            inSent.Start(message);
        }
        // Обработка завершения отправки письма
        private void OutputMail_OnSended(MailBox box)
        {
            // Создание дополнительного Imap-клиента для перемещения отправленного письма
            IMAPWrapper imap = new IMAPWrapper();
            imap.OnException += _OnEsception;
            // Перемещение отправленного письма из папки INBOX в папку SENT
            imap.DragToSent(currentMailBox);
            // Отключение дополнительного Imap-клиента от сервера
            imap.Disconnect();
            throw new Exception("Cообщение отправлено");
        }
        #endregion

        // Обработка несинхронных ошибок
        private void _OnEsception(Exception ex)
        {
            OnException(ex);
        }        
               
        // Проверка наличия в списке почтового ящика с указанным адресом
        public bool ConsistAddress(string address)
        {
            bool consist = false;
            foreach (MailBox mb in mailBoxList)
                consist = address == mb.MyAddress.Address;
            return consist;
        }     

        // Метод отключения от сервера входящей почты
        public void Disconnect()
        {
            inputMail.Disconnect();            
        }

        #region События
        public event OnExceptionHandler OnException;
        public event ViewHandler OnViewListMess;
        public event ViewHandler OnViewMess;
        #endregion
    }

    // Класс почтового ящика
    public class MailBox
    {
        #region Конструкторы
        // Создание почтового ящика
        public MailBox()
        {
            name = null;
            myAddress = new MailAddress("user@host.domain");
            settings = new SettingsClass()
            {
                ImapServer = "imap.ru",
                ImapPort = 993,
                SmtpServer = "smtp.ru",
                SmtpPort = 25,
            };
            folders = new List<Folder>();           
        }
        // Создание копии почтового ящика
        public MailBox(MailBox newBox)
        {
            this.name = newBox.name;
            this.myAddress = new MailAddress (newBox.myAddress.Address);
            this.settings = new SettingsClass (newBox.settings);
            this.folders = new List<Folder> (newBox.folders);
        }
        // Создание почтового ящика с параметрами
        public MailBox(string namebox, string address, SettingsClass settingBox)
        {
            name = namebox;
            myAddress = new MailAddress(address);
            settings = new SettingsClass (settingBox);
            folders = new List<Folder>();
        }
        // Создание почтового ящика с параметрами
        public MailBox(string namebox, string address, string user, string pass,
            bool Ssl, string ImapServer, int ImapPort, string SmtpServer, int SmtpPort) : 
            this(namebox,address, new SettingsClass(user, pass, Ssl, ImapServer, ImapPort, SmtpServer, SmtpPort))
        {          
        }
        #endregion

        private string name;            // Имя почтового ящика
        private MailAddress myAddress;  // Адрес почтового ящика
        private List<Folder> folders;   // Список папок ящика
        private Folder selectedFolder;  // Выбранная папка
        private SettingsClass settings; // Настройки почтового ящика

        #region Публичные свойства

        public string Name
        {
            set { name = value; }
            get { return name; }
        }

        public MailAddress MyAddress
        {
            set { myAddress = value; }
            get { return myAddress; }
        }

        public List<Folder> Folders
        {
            set { folders = value; }
            get { return folders; }
        }

        public Folder SelectedFolder
        {
            set { selectedFolder = value; }
            get { return selectedFolder; }
        }
        public SettingsClass Settings
        {
            set { settings = value; }
            get { return settings; }
        }
        #endregion
    }

    // Типы почтовых папок
    public enum TYPE_FOLDER
    {
        INBOX, SENT, DRAFTS, JUNK, TRASH, TEMPLATE, NONE
    }

    // Класс папки
    public class Folder
    {
        #region Конструкторы
        // Создание папки
        public Folder ()
        {
            messages = new List<MyMailMessage>();
        }

        // Создание папки с параметрами
        public Folder(string _pointer, string _name, TYPE_FOLDER _type) : this()
        {
            name = _name;
            type = _type;
            pointer = _pointer;
        }

        // Создание копии папки
        public Folder(Folder fld)
        {
            this.name = fld.name;
            this.type = fld.type;
            this.pointer = fld.pointer;
            this.messages = new List<MyMailMessage>(fld.messages);
        }
        #endregion

        private string pointer;     // Указатель на папку
        private string name;        // Имя папки
        private TYPE_FOLDER type;   // Тип папки
        private List<MyMailMessage> messages;   // Список писем
        private MyMailMessage currentMessage;   // Текщее письмо

        #region Публичные свойства

        public string Name
        {
            set { name = value; }
            get { return name; }
        }

        public string Pointer
        {
            set { pointer = value; }
            get { return pointer; }
        }

        public TYPE_FOLDER Type
        {
            set { type = value; }
            get { return type; }
        }

        public MyMailMessage Message
        {
            set { currentMessage = value; }
            get { return currentMessage; }
        }

        public List<MyMailMessage> Messages
        {
            set { messages = value; }
            get { return messages; }
        }
        #endregion
    }

    // Класс входящего письма
    public class MyMailMessage : MailMessage
    {

        private string id;                      // идентификатор письма
        private int num;                        // номер письма
        private DateTime date;                  // дата письма
        private DateString displayDate;         // показываемая дата
        private ContentType mainContentType;    // Тип содержимого письма
        private bool isHaveAttachments;         // Наличие вложений
        private List<MyAttachment> myAttachments;   // Вложения        

        #region Публичные свойства

        public string ID
        {
            get { return id; }
        }

        public int Num
        {
            get { return num; }
        }

        public DateTime Date
        {
            get { return date; }
        }

        public DateString DisplayDate
        {
            get { return displayDate; }
        }

        public bool IsHaveAttachments
        {
            get { return isHaveAttachments; }
        }

        public List<MyAttachment> MyAttachments
        {
            get { return myAttachments; }
        }
        #endregion

        #region Конструкторы
        // Создание письма с заголовком
        public MyMailMessage(string header) : base()
        {
            Headers.Add("main",header.Substring(header.IndexOf("\r\n")+2));
            num = Convert.ToInt32(ParserMessage.FirstFromTo(header, " ", " "));
            id = ParserMessage.GetID(header);
            Sender = From = ParserMessage.GetFrom(header);
            MailAddressCollection temp = ParserMessage.GetTo(header);
            foreach(MailAddress adr in temp) To.Add(adr);
            Subject = ParserMessage.GetSubject(header);
            temp = ParserMessage.GetReply(header);
            foreach (MailAddress adr in temp) ReplyToList.Add(adr);
            temp = ParserMessage.GetCC(header);
            foreach (MailAddress adr in temp) CC.Add(adr);
            temp = ParserMessage.GetBcc(header);
            foreach (MailAddress adr in temp) Bcc.Add(adr);
            date = ParserMessage.GetDate(header);
            displayDate = new DateString(date);
            mainContentType = ParserMessage.GetContentType(header);
            BodyEncoding = ParserMessage.MyGetEncoding(mainContentType.CharSet);
            BodyTransferEncoding = ParserMessage.GetBodyTransfer(header);
            isHaveAttachments = (mainContentType.MediaType == "multipart/mixed" || mainContentType.MediaType == "multipart/related");
            IsBodyHtml = (mainContentType.MediaType == MediaTypeNames.Text.Html);
            myAttachments = new List<MyAttachment>();            
        }

        // Создание копии письма
        public MyMailMessage(MyMailMessage message)
        {
            Headers.Add("main", message.Headers.GetValues("main")[0]);
            num = message.Num;
            id = message.ID;
            Sender = From = message.From;
            foreach (MailAddress adr in message.To) To.Add(adr);
            Subject = message.Subject;
            foreach (MailAddress adr in message.ReplyToList) ReplyToList.Add(adr);
            foreach (MailAddress adr in message.CC) CC.Add(adr);
            foreach (MailAddress adr in message.Bcc) Bcc.Add(adr);
            date = message.Date;
            displayDate = message.DisplayDate;
            mainContentType = message.mainContentType;
            BodyEncoding = message.BodyEncoding;
            BodyTransferEncoding = message.BodyTransferEncoding;
            isHaveAttachments = message.IsHaveAttachments;
            IsBodyHtml = message.IsBodyHtml;
            myAttachments = new List<MyAttachment>();
            foreach(MyAttachment att in message.MyAttachments) myAttachments.Add(att);
            Body = message.Body;
        }
        #endregion

        #region Получение тела письма
        // Метод добавления тела письма
        public void AddBody(string body)
        {
            Body = body;
            if (isHaveAttachments)
            {
                // Добавление вложений
                AddAttachments(body, mainContentType.Boundary);
                // Если после добавления вложений текст письма не изменился
                // значит письмо имеет пустой текст
                if (Body == body) Body = "";
            }
            if (mainContentType.MediaType == "multipart/alternative")
                // Добавление альтернативных способов представления текста письма
                AddViews(body, mainContentType.Boundary);
            // Декодирование текста письма
            Body = ParserMessage.GetCodeLine(Body, BodyEncoding, BodyTransferEncoding);
        }

        // Метод добавления тела письма  из файла
        public void AddBody(MailBox box)
        {
            MyMailMessage mess = MailStorage.LoadBody(box);
            Body = mess.Body;
            IsBodyHtml = mess.IsBodyHtml;
            myAttachments = new List<MyAttachment>();
            foreach (MyAttachment at in mess.myAttachments)
                myAttachments.Add(new MyAttachment(at));
        }
        #endregion

        #region Получение вложений
        // Метод добавления вложений
        private void AddAttachments(string _body, string _boundary)
        {
            string stringAttach = null;
            int start = _body.IndexOf(_boundary) + _boundary.Length;
            _body = _body.Substring (start, _body.LastIndexOf(_boundary) - 4 - start);
            while (_body.Contains(_boundary + "\r\n"))
            {
                // Строка с одним вложением
                stringAttach = _body.Substring(0, _body.IndexOf(_boundary) - 4);
                // Получить вложение из строки
                GetAttach(stringAttach);
                // Отрезать от строки с вложениями просмотренную строку с вложением
                _body = _body.Substring(_body.IndexOf(_boundary) + _boundary.Length);
            }
            // Получить вложение из строки
            GetAttach(_body);            
        }
        // Метод получения вложения
        private void GetAttach(string text)
        {
            // Получение заголовка вложения
            string header = text.Substring(0, text.IndexOf("\r\n\r\n"));
            // Получение тела вложения
            string body = text.Substring(text.IndexOf("\r\n\r\n") + 4);
            // Тип контента вложения
            ContentType ct = ParserMessage.GetContentType(header);
            // Расположение контента
            ContentDisposition cd = ParserMessage.GetDisposition(header);
            // Получение транспортной кодировки
            TransferEncoding transEncode = ParserMessage.GetBodyTransfer(header);
            // Декодирование строки с именем вложения
            if (ct.Name != null)
                ct.Name = ParserMessage.DecodeString(ct.Name);
            // Добавление вложения внутри вложения
            if (ct.MediaType == "multipart/mixed" || ct.MediaType == "multipart/related")
                AddAttachments(body, ct.Boundary);
            // Получение текста письма в формате html-страницы
            else if (ct.MediaType == MediaTypeNames.Text.Html)
            {
                if (cd == null)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                    IsBodyHtml = true;
                }
                else if (cd.DispositionType != DispositionTypeNames.Attachment)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                    IsBodyHtml = true;
                }
            }
            // Получение текста письма в виде обычного текста
            else if (ct.MediaType == MediaTypeNames.Text.Plain && !IsBodyHtml)
            {
                if (cd == null)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                }
                else if (cd.DispositionType != DispositionTypeNames.Attachment)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                }
            }
            // Добавление альтернативных способов представления текста письма
            else if (ct.MediaType == "multipart/alternative")
            {
                AddViews(body, ct.Boundary);
            }
            // Декодирование тела вложения и добавление в список
            else
            {
                body = ParserMessage.GetCodeLine(body, Encoding.Default, transEncode);
                myAttachments.Add(new MyAttachment(ct, body));
            }
        }
        #endregion

        #region Получение альтернативных способов просмотра письма
        // Метод добавления альтернативных способов просмотра письма
        private void AddViews(string _body, string _boundary)
        {
            string stringView = null;
            int start = _body.IndexOf(_boundary) + _boundary.Length;
            _body = _body.Substring(start, _body.LastIndexOf(_boundary) - 4 - start);
            while (_body.Contains(_boundary))
            {
                // Строка с одним альтернативным способом представления текста
                stringView = _body.Substring(0, _body.IndexOf(_boundary) - 4);
                // Получить альтернативный способ представления текста из строки
                GetView(stringView);
                // Отрезать от строки с альтернативными способами представления текста
                // просмотренную строку с альтернативным способом представления текста
                _body = _body.Substring(_body.IndexOf(_boundary) + _boundary.Length);
            }
            // Получить альтернативный способ представления текста из строки
            GetView(_body);
        }
        // Метод получения альтернативного способа просмотра письма
        private void GetView(string text)
        {
            if (text.Contains("\r\n\r\n"))
            {
                // Заголовок альтернативного способа просмотра письма
                string header = text.Substring(0, text.IndexOf("\r\n\r\n"));
                // Тело альтернативного способа просмотра письма
                string body = text.Substring(text.IndexOf("\r\n\r\n") + 4);
                // Тип контента альтернативного способа просмотра письма
                ContentType ct = ParserMessage.GetContentType(header);
                // Получение транспортной кодировки
                TransferEncoding transEncode = ParserMessage.GetBodyTransfer(header);
                // Получение текста письма в формате html-страницы
                if (ct.MediaType == MediaTypeNames.Text.Html)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                    IsBodyHtml = true;
                }
                // Получение текста письма в виде обычного текста
                else if (ct.MediaType == MediaTypeNames.Text.Plain && !IsBodyHtml)
                {
                    BodyEncoding = ParserMessage.MyGetEncoding(ct.CharSet);
                    BodyTransferEncoding = transEncode;
                    Body = body;
                }
            }
        }
        #endregion
    }

    // Класс отображаемой даты
    public class DateString
    {
        private string shotDate;    // Короткая дата
        private string fullDate;    // Полная дата

        // Конструктор класса отображаемой даты
        public DateString(DateTime date)
        {
            // В полном виде показывать дату и время
            fullDate = date.ToLongDateString() + "  " + date.ToShortTimeString();
            // Если сегодняшняя дата письма в коротком виде показывать только время
            if (date.Date == DateTime.Now.Date) shotDate = date.ToShortTimeString();
            // Если с даты письма не прошло недели в коротком виде показывать день недели
            else if (date.AddDays(7.0) > DateTime.Now.Date) shotDate = date.DayOfWeek.ToString();
            // В остальных случаях в коротком виде показывать дату без времени
            else shotDate = date.ToLongDateString();
        }

        #region Публичные свойства
        public string ShotDate
        {
            get { return shotDate; }
        }

        public string FullDate
        {
            get { return fullDate; }
        }
        #endregion
    }

    // Класс вложения входящего письма
    public class MyAttachment
    {
        private ContentType contentType;    // Тип контента вложения
        private string content;             // Данные файла в виде строки
        private string name;                // Имя вложения

        #region Конструкторы
        // Создание вложения по типу контента и данным файла
        public MyAttachment(ContentType _contentType, string _content)
        {
            contentType = _contentType;
            content = _content;
            name = _contentType.Name;
        }
        // Создание копии вложения
        public MyAttachment(MyAttachment attach)
        {
            contentType = attach.ContentTypeAt;
            content = attach.Content;
            name = attach.Name;
        }
        #endregion
        // Метод очистки памяти от вложения
        public void Clear()
        {
            // Обнуление данных файла
            content = null;
        }

        #region Публичные свойства
        public ContentType ContentTypeAt
        {
            get { return contentType; }
        }

        public string Content
        {
            get { return content; }
        }

        public string Name
        {
            get { return name; }
        }
        #endregion
    }

    // Класс настроек почтового ящика
    public class SettingsClass
    {
        #region Конструкторы
        // Создание настроек почтового ящика
        public SettingsClass()
        {
            
        }

        // Создание настроек почтового ящика с параметрами
        public SettingsClass(string user, string pass, 
            bool Ssl, string ImapServer, int ImapPort, string SmtpServer, int SmtpPort)
        {
            userName = user;
            password = pass;
            ssl = Ssl;
            imapServer = ImapServer;
            imapPort = ImapPort;
            smtpServer = SmtpServer;
            smtpPort = SmtpPort;
        }

        // Создание копии настроек почтового ящика
        public SettingsClass(SettingsClass settings)
        {
            this.userName = settings.userName;
            this.password = settings.password;
            this.ssl = settings.ssl;
            this.imapServer = settings.imapServer;
            this.imapPort = settings.imapPort;
            this.smtpServer = settings.smtpServer;
            this.smtpPort = settings.smtpPort;
        }
        #endregion

        private string userName;    // Логин почтового ящика
        private string password;    // Пароль
        private bool ssl;           // Использование шифрование
        private string imapServer;  // Адрес сервера входящей почты
        private int imapPort;       // Порт сервера входящей почты
        private string smtpServer;  // Адрес сервера исходящей почты
        private int smtpPort;       // Порт сервера исходящей почты

        #region Публичные свойства
        public string UserName
        {
            set { userName = value; }
            get { return userName; }
        }

        public string Rassword
        {
            set { password = value; }
            get { return password; }
        }

        public bool Ssl
        {
            set { ssl = value; }
            get { return ssl; }
        }

        public string ImapServer
        {
            set { imapServer = value; }
            get { return imapServer; }
        }

        public string SmtpServer
        {
            set { smtpServer = value; }
            get { return smtpServer; }
        }

        public int ImapPort
        {
            set { imapPort = value; }
            get { return imapPort; }
        }

        public int SmtpPort
        {
            set { smtpPort = value; }
            get { return smtpPort; }
        }
        #endregion
    }
}
