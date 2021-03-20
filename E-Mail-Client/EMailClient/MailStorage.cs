using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Net.Mime;

namespace E_Mail_Client.EMailClient
{
    // Класс для работы с XML-хранилищем
    public class MailStorage
    {
        // Имя файла для хранения информации
        private static string filename = "mailboxes.stg";
        
        //создание xml файла хранилища
        private static void CreateXML(string filepath)
        {
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.Default);
            xtw.WriteStartDocument();
            xtw.WriteStartElement("MailBoxes");
            xtw.WriteEndDocument();
            xtw.Close();
        }

        // Метод шифрования пароля
        private static string Cript(string pass)
        {
            byte[] mb = Encoding.Unicode.GetBytes(pass);
            byte[] cmb = new byte[mb.Length + 1];
            for (int i = 0; i < mb.Length; i++)
            {
                cmb[i] = (byte)(cmb[i] | (mb[i] >> 5));
                cmb[i + 1] = (byte)(mb[i] << 3);
            }
            string cript = Encoding.GetEncoding(1251).GetString(cmb);
            return cript;
        }

        // Метод дешифрирования пароля
        private static string UnCript(string cript)
        {
            if (cript.Length < 2) return cript;
            else
            {
                byte[] cmb = Encoding.GetEncoding(1251).GetBytes(cript);
                byte[] mb = new byte[cmb.Length - 1];
                for (int i = 0; i < mb.Length; i++)
                {
                    mb[i] = (byte)(cmb[i] << 5 | (cmb[i + 1] >> 3));
                }
                string pass = Encoding.Unicode.GetString(mb);
                return pass;
            }
        }

        // Метод загрузки списка почтовых ящиков
        public static List<MailBox> LoadMailBoxList()
        {
            List<MailBox> boxes = new List<MailBox>();
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.GetElementsByTagName("MailBox");
            foreach (XmlNode xn in list)
            {
                MailBox mb = new MailBox(((XmlElement)xn).GetAttribute("name"),
                                        ((XmlElement)xn).GetAttribute("address"), new SettingsClass());
                boxes.Add(mb);
            }
            return boxes;
        }

        // Метод загрузки настроек почтового ящика
        public static SettingsClass LoadSettingBox(string address)
        {
            SettingsClass sc;
            XmlElement setting = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == address)
                {
                    setting = (XmlElement)node.FirstChild;
                    break;
                }
            string user, password, ssl, imapHost, imapPort, smtpHost, smtpPort;
            if (setting != null)
            {
                user = setting.Attributes["user"].Value;
                password = UnCript(setting.Attributes["password"].Value);
                ssl = setting.Attributes["ssl"].Value;
                XmlElement imap = (XmlElement)setting.ChildNodes[0];
                XmlElement smtp = (XmlElement)setting.ChildNodes[1];
                imapHost = imap.Attributes["host"].Value;
                imapPort = imap.Attributes["port"].Value;
                smtpHost = smtp.Attributes["host"].Value;
                smtpPort = smtp.Attributes["port"].Value;
                sc = new SettingsClass(user, password, Convert.ToBoolean(ssl),
                                        imapHost, Convert.ToInt32(imapPort), smtpHost, Convert.ToInt32(smtpPort));
            }
            else sc = new SettingsClass();
            return sc;
        }

        // Метод сохранения нового почтового ящика
        public static void SaveNewMailBox(MailBox box)
        {
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);

            XmlElement mailBox = xd.CreateElement("MailBox");
            mailBox.SetAttribute("name", box.Name);
            mailBox.SetAttribute("address", box.MyAddress.Address);
            XmlElement settings = xd.CreateElement("settings");
            settings.SetAttribute("user", box.Settings.UserName);
            settings.SetAttribute("password", Cript(box.Settings.Rassword));
            settings.SetAttribute("ssl", box.Settings.Ssl.ToString());
            XmlElement imapserver = xd.CreateElement("imapserver");
            imapserver.SetAttribute("host", box.Settings.ImapServer);
            imapserver.SetAttribute("port", box.Settings.ImapPort.ToString());
            XmlElement smtpserver = xd.CreateElement("smtpserver");
            smtpserver.SetAttribute("host", box.Settings.SmtpServer);
            smtpserver.SetAttribute("port", box.Settings.SmtpPort.ToString());
            settings.AppendChild(imapserver);
            settings.AppendChild(smtpserver);
            mailBox.AppendChild(settings);
            XmlElement folders = xd.CreateElement("folders");
            mailBox.AppendChild(folders);
            xd.DocumentElement.AppendChild(mailBox);
            xd.Save(filename);
        }

        // Метод сохранения настроек почтового ящика
        public static void SaveSettings(MailBox box)
        {
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    // Ввод настроек
                    XmlElement settings = xd.CreateElement("settings");
                    settings.SetAttribute("user", box.Settings.UserName);
                    settings.SetAttribute("password", Cript(box.Settings.Rassword));
                    settings.SetAttribute("ssl", box.Settings.Ssl.ToString());
                    XmlElement imapserver = xd.CreateElement("imapserver");
                    imapserver.SetAttribute("host", box.Settings.ImapServer);
                    imapserver.SetAttribute("port", box.Settings.ImapPort.ToString());
                    XmlElement smtpserver = xd.CreateElement("smtpserver");
                    smtpserver.SetAttribute("host", box.Settings.SmtpServer);
                    smtpserver.SetAttribute("port", box.Settings.SmtpPort.ToString());
                    settings.AppendChild(imapserver);
                    settings.AppendChild(smtpserver);
                    // Замена старых настроек на новые
                    node.ReplaceChild(settings, node.FirstChild);
                }
            xd.Save(filename);
        }

        // Удаление почтового ящика из файла по адресу
        public static void RemoveMailBox(string address)
        {
            if (File.Exists(filename))
            {
                XmlDocument xd = new XmlDocument();
                xd.Load(filename);
                XmlNodeList list = xd.DocumentElement.ChildNodes;
                foreach (XmlNode node in list)
                    if (node.Attributes["address"].Value == address)
                        xd.DocumentElement.RemoveChild(node);
                xd.Save(filename);
            }
        }

        // Метод сохранения новых загруженных папок в файл
        // и удаление из файла отсутствующих на сервере
        public static void SaveFolders(MailBox box)
        {
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    XmlElement folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Удаление из файла папок отсутствующих на сервере
                    foreach (XmlNode fld in fList)
                    {
                        if (!box.Folders.Exists(x => x.Pointer == fld.Attributes["pointer"].Value))
                            folders.RemoveChild(fld);
                    }
                    foreach (Folder fldr in box.Folders)
                    {
                        // Проверка по указателю отсутствие загруженной папки в файле
                        bool isNotFolder = true;
                        foreach (XmlNode fld in fList)
                        {
                            if (fldr.Pointer == fld.Attributes["pointer"].Value)
                                isNotFolder = false;
                        }
                        // Загрузка новой папки в файл
                        if (isNotFolder)
                        {
                            XmlElement folder = xd.CreateElement("folder");
                            folder.SetAttribute("pointer", fldr.Pointer);
                            folder.SetAttribute("name", fldr.Name);
                            folder.SetAttribute("type", fldr.Type.ToString());
                            XmlElement messages = xd.CreateElement("messages");
                            folder.AppendChild(messages);
                            folders.AppendChild(folder);
                        }
                    }
                    break;
                }
            xd.Save(filename);
        }

        // Метод загрузки списка папок из файла
        public static List<Folder> LoadFolders(string adress)
        {
            List<Folder> flds = new List<Folder>();
            XmlElement folders = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == adress)
                {
                    folders = (XmlElement)node.LastChild;
                    break;
                }
            if (folders != null)
            {
                XmlNodeList listFld = folders.ChildNodes;
                foreach (XmlElement fld in listFld)
                {
                    string name = fld.GetAttribute("name");
                    string pointer = fld.GetAttribute("pointer");
                    TYPE_FOLDER type = (TYPE_FOLDER)Enum.Parse(typeof(TYPE_FOLDER), fld.GetAttribute("type"));
                    flds.Add(new Folder(pointer, name, type));
                }
            }
            return flds;
        }

        // Метод сохранения списка новых загруженных заголовков в файл
        // и удаление из файла отсутствующих на сервере писем
        public static void SaveHeaders(MailBox box)
        {
            XmlElement folders = null;
            XmlElement messages = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Выбор папки по указателю
                    foreach (XmlNode fld in fList)
                        if (fld.Attributes["pointer"].Value == box.SelectedFolder.Pointer)
                        {
                            messages = (XmlElement)fld.FirstChild;
                            XmlNodeList messList = messages.ChildNodes;
                            // Удаление из файла писем отсутствующих на сервере 
                            foreach (XmlNode mess in messList)
                            {
                                if (!box.SelectedFolder.Messages.Exists(x => x.ID == mess.Attributes["id"].Value))
                                    messages.RemoveChild(mess);
                            }
                            // 
                            foreach (MyMailMessage mssg in box.SelectedFolder.Messages)
                            {
                                XmlElement message = xd.CreateElement("message");
                                // Проверка по идентификатору отсутствие загруженного заголовка письма в файле
                                bool isNotMessage = true;
                                foreach (XmlNode msg in messList)
                                {
                                    if (mssg.ID == msg.Attributes["id"].Value)
                                        isNotMessage = false;
                                }
                                // Загрузка нового письма в файл
                                if (isNotMessage)
                                {
                                    message.SetAttribute("id", mssg.ID);
                                    message.SetAttribute("header", mssg.Headers.Get("main"));
                                    XmlElement attachments = xd.CreateElement("attachments");
                                    message.AppendChild(attachments);
                                    messages.AppendChild(message);
                                }
                                message.SetAttribute("num", mssg.Num.ToString());
                            }
                            break;
                        }
                    break;
                }
            xd.Save(filename);
        }               

        // Метод загрузки списка заголовков писем из файла
        public static List<MyMailMessage> LoadHeaders(MailBox box)
        {
            List<MyMailMessage> mmm = new List<MyMailMessage>();
            XmlElement folders = null;
            XmlElement messages = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Выбор папки по указателю
                    foreach (XmlNode fld in fList)
                        if (fld.Attributes["pointer"].Value == box.SelectedFolder.Pointer)
                        {
                            messages = (XmlElement)fld.FirstChild;
                            XmlNodeList messList = messages.ChildNodes;
                            foreach (XmlElement mess in messList)
                            {
                                string num = mess.GetAttribute("num");
                                string header = mess.GetAttribute("header");
                                mmm.Add(new MyMailMessage(" " + num + " " + header));
                            }
                            break;
                        }
                    break;
                }
            return mmm;
        }

        // Метод сохранения текста загруженного письма с вложенными файлами
        public static void SaveMessage(MailBox box)
        {
            XmlElement folders = null;
            XmlElement messages = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Выбор папки по указателю
                    foreach (XmlNode fld in fList)
                        if (fld.Attributes["pointer"].Value == box.SelectedFolder.Pointer)
                        {
                            messages = (XmlElement)fld.FirstChild;
                            XmlNodeList messList = messages.ChildNodes;
                            // Выбор письма по идентификатору
                            foreach (XmlElement message in messList)
                                if (message.Attributes["id"].Value == box.SelectedFolder.Message.ID)
                                {
                                    // Добавление тела письма
                                    MyMailMessage msg = box.SelectedFolder.Message;
                                    message.SetAttribute("body", msg.Body);
                                    message.SetAttribute("isHtml", msg.IsBodyHtml.ToString());
                                    XmlElement attachments = (XmlElement)message.LastChild;
                                    XmlNodeList attachList = attachments.ChildNodes;
                                    if (attachList.Count == 0)
                                        // Добавление вложений
                                        foreach (MyAttachment attach in msg.MyAttachments)
                                        {
                                            XmlElement attachment = xd.CreateElement("attachment");
                                            attachment.SetAttribute("contenttype", attach.ContentTypeAt.ToString());
                                            attachment.SetAttribute("content", attach.Content);
                                            attachment.SetAttribute("name", attach.Name);
                                            attachments.AppendChild(attachment);
                                        }
                                    break;
                                }
                            break;
                        }
                    break;
                }
            xd.Save(filename);
        }
        
        // Метод загрузки текста
        public static MyMailMessage LoadBody(MailBox box)
        {
            MyMailMessage outMess = new MyMailMessage(box.SelectedFolder.Message);
            XmlElement folders = null;
            XmlElement messages = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Выбор папки по указателю
                    foreach (XmlNode fld in fList)
                        if (fld.Attributes["pointer"].Value == box.SelectedFolder.Pointer)
                        {
                            messages = (XmlElement)fld.FirstChild;
                            XmlNodeList messList = messages.ChildNodes;
                            // Выбор письма по идентификатору
                            foreach (XmlElement message in messList)
                                if (message.Attributes["id"].Value == box.SelectedFolder.Message.ID)
                                {
                                    outMess.Body = message.GetAttribute("body");
                                    string isHtml = message.GetAttribute("isHtml") == "" ? "False" : message.GetAttribute("isHtml");
                                    outMess.IsBodyHtml = Convert.ToBoolean(isHtml);
                                    XmlElement attachments = (XmlElement)message.LastChild;
                                    XmlNodeList attachList = attachments.ChildNodes;
                                    foreach (XmlElement attachment in attachList)
                                    {
                                        ContentType ct = new ContentType(attachment.GetAttribute("contenttype"));
                                        outMess.MyAttachments.Add(new MyAttachment(ct, null));
                                    }
                                    break;
                                }
                            break;
                        }
                    break;
                }
            return outMess;
        }

        // Загрузка вложения из сохраненного письма в виде массива байтов
        public static byte[] LoadAttach(MailBox box, string attachName)
        {
            string content = null;
            XmlElement folders = null;
            XmlElement messages = null;
            XmlElement attachments = null;
            if (!File.Exists(filename)) CreateXML(filename);
            XmlDocument xd = new XmlDocument();
            xd.Load(filename);
            XmlNodeList list = xd.DocumentElement.ChildNodes;
            // Выбор почтового ящика по адресу
            foreach (XmlNode node in list)
                if (node.Attributes["address"].Value == box.MyAddress.Address)
                {
                    folders = (XmlElement)node.LastChild;
                    XmlNodeList fList = folders.ChildNodes;
                    // Выбор папки по указателю
                    foreach (XmlNode fld in fList)
                        if (fld.Attributes["pointer"].Value == box.SelectedFolder.Pointer)
                        {
                            messages = (XmlElement)fld.FirstChild;
                            XmlNodeList messList = messages.ChildNodes;
                            // Выбор письма по идентификатору
                            foreach (XmlNode mess in messList)
                                if (mess.Attributes["id"].Value == box.SelectedFolder.Message.ID)
                                {
                                    attachments = (XmlElement)mess.LastChild;
                                    XmlNodeList attachList = attachments.ChildNodes;
                                    // Выбор вложения по имени
                                    foreach (XmlNode attach in attachList)
                                        if (attach.Attributes["name"].Value == attachName)                                        
                                        {
                                            XmlElement attachment = (XmlElement)attach;
                                            content = attachment.GetAttribute("content");
                                            break;
                                        }
                                    break;
                                }
                            break;
                        }
                    break;
                }
            byte[] byteContent;
            return byteContent = Encoding.Default.GetBytes(content);
        }    
    }       
}
