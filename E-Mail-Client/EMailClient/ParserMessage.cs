using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;

namespace E_Mail_Client.EMailClient
{
    // Класс обработки ответов почтовых серверов
    public class ParserMessage
    {
        // Метод получения тела сообщения
        public static string GetBody(string letter)
        {            
            string textMsg = letter.Substring(letter.IndexOf('{') + 1);
            string sizeS = textMsg.Substring(0, textMsg.IndexOf('}'));
            int sizeI = Convert.ToInt32(sizeS);
            return textMsg.Substring(sizeS.Length + 3, sizeI);
        }

        #region Методы работы со строками
        // Метод удаления символов chars из строки
        public static string RemChar(string text, string chars)
        {
            int indexStart, indexFinish;
            int LenCh = chars.Length;
            indexStart = indexFinish = text.IndexOf(chars);
            while (indexStart != -1)
            {
                if (indexFinish < text.Length - LenCh)
                    while (text.Substring(indexFinish, LenCh) == chars && indexFinish < text.Length - LenCh)
                        indexFinish += LenCh;
                int count = indexFinish == indexStart ? LenCh : indexFinish - indexStart;
                text = text.Remove(indexStart, count);
                indexStart = indexFinish = text.IndexOf(chars);                
            }
            return text;
        }

        // Метод выделения первого фрагмента выделенного строками
        public static string FirstFromTo(string text, string from, string to)
        {
            if (text.IndexOf(from) < 0 || text.LastIndexOf(to) < 0 || (text.IndexOf(from) == text.LastIndexOf(to)))
                return "";
            else
            {
                string textMsg = text.Substring(text.IndexOf(from) + from.Length);
                return textMsg.Substring(0, textMsg.IndexOf(to));
            }
        }

        // Метод выделения первого фрагмента выделенного строками
        // после ключевого слова
        public static string FirstFromToByKey(string text, string from, string to, string key)
        {
            string textMsg;
            if (text.Contains(key))
            {
                textMsg = text.Substring(text.IndexOf(key) + key.Length);
                return FirstFromTo(textMsg, from, to);
            }
            return "";
        }

        // Метод выделения последнего фрагмента выделенного строками
        public static string LastFromTo(string text, string from, string to)
        {
            if (text.IndexOf(from) == -1 || text.LastIndexOf(to) == -1 || (text.IndexOf(from) == text.LastIndexOf(to)))
                return null;
            else
            {
                string textMsg = text.Substring(0, text.LastIndexOf(to));
                return textMsg.Substring(textMsg.LastIndexOf(from) + from.Length);
            }
        }
        #endregion

        #region Обработка ответов сервера
        // Получение из ответа LIST списка папок
        public static List<Folder> GetFolders(string answer)
        {
            List<Folder> folders = new List<Folder>();
            string[] tabFolders = answer.Split('*');
            foreach (string fld in tabFolders)
            {
                if (fld != "")
                {
                    if (fld.Contains("INBOX") || fld.Contains("Inbox") || fld.Contains("inbox"))
                        folders.Insert(0, new Folder("INBOX", "Входящие", TYPE_FOLDER.INBOX));
                    else if (fld.Contains("Template") || fld.Contains("template") || fld.Contains("TEMPLATE"))
                        folders.Add(new Folder("TEMPLATE", "Шаблоны", TYPE_FOLDER.TEMPLATE));
                    else
                    {
                        string _name = LastFromTo(fld, "\"", "\"");
                        string point = _name;
                        string outname = null;
                        while (_name.IndexOf("&") >= 0)
                        {
                            outname += _name.Substring(0, _name.IndexOf("&"));
                            _name = _name.Substring(_name.IndexOf("&"));
                            if (_name.IndexOf("-") > 0)
                            {
                                string rus = FirstFromTo(_name, "&", "-");
                                if (rus == "") outname += "&";
                                else
                                {
                                    outname += Decode_IMAP_UTF7_String(rus);
                                }
                                if (_name.IndexOf("-") == _name.Length - 1)
                                    _name = "";
                                else _name = _name.Substring(_name.IndexOf("-") + 1);
                            }

                        }
                        outname += _name;
                        if (fld.Contains("Sent") || fld.Contains("sent") || fld.Contains("SENT"))
                            folders.Add(new Folder(point, outname, TYPE_FOLDER.SENT));
                        else if (fld.Contains("Junk") || fld.Contains("junk") || fld.Contains("JUNK"))
                            folders.Add(new Folder(point, outname, TYPE_FOLDER.JUNK));
                        else if (fld.Contains("Drafts") || fld.Contains("drafts") || fld.Contains("DRAFTS"))
                            folders.Add(new Folder(point, outname, TYPE_FOLDER.DRAFTS));
                        else if (fld.Contains("Trash") || fld.Contains("trash") || fld.Contains("TRASH"))
                            folders.Add(new Folder(point, outname, TYPE_FOLDER.TRASH));
                        else folders.Add(new Folder(point, outname, TYPE_FOLDER.NONE));

                    }
                }
            }
            return folders;
        }

        // Получение из ответа STATUS количества писем в папке
        public static int GetNumberMess(string ans)
        {
            ans = ans.Substring(ans.IndexOf("MESSAGES"));
            return Convert.ToInt32(FirstFromTo(ans, " ", ")"));
        }
        #endregion

        #region Обработка заголовков
        // Получение адреса MailAddress из строки с почтовым адресом
        private static MailAddress GetAddress(string fullAddress)
        {
            MailAddress Address;
            try
            {
                if (fullAddress.Contains("<"))
                {
                    // Получение попочтового адреса
                    string address = FirstFromTo(fullAddress, "<", ">");

                    // Получение отображаемого имени и его кодировки
                    string displayName = "";
                    string dn = fullAddress.Substring(0, fullAddress.IndexOf("<"));
                    Encoding encDN = Encoding.Default;
                    string transferCoding = "";
                    while (dn.IndexOf("=?") >= 0)
                    {
                        if (dn.Substring(dn.LastIndexOf("=?")).Contains("?="))
                        {
                            string temp = LastFromTo(dn, "=?", "?=");
                            displayName = temp.Substring(temp.LastIndexOf("?") + 1) + displayName;
                            encDN = MyGetEncoding(temp.Substring(0, temp.IndexOf("?")));
                            transferCoding = FirstFromTo(temp, "?", "?");
                        }
                        dn = dn.Substring(0, dn.LastIndexOf("=?"));
                    }
                    displayName = GetCodeLine(displayName, encDN, GetTransferEncoding(transferCoding));
                    // Создание объекта почтового адреса с ототбражаемым именем
                    Address = new MailAddress(address, displayName, encDN);
                }
                // Создание объекта почтового адреса без отображаемого имени
                else Address = new MailAddress(fullAddress);
            }
            catch (Exception ex)
            {
                return null;
            }
            return Address;
        }

        // Получение коллекции адресов из строки
        private static MailAddressCollection GetAddressColection(string fullAddress)
        {
            string[] addresses = fullAddress.Split(',');
            MailAddressCollection adresCol = new MailAddressCollection();
            foreach (string adress in addresses)
            {
                MailAddress temp = GetAddress(adress);
                if (temp != null)
                    adresCol.Add(temp);
            }
            return adresCol;
        }

        #region Декодирование заголовков
        // Получение строки после ключевого слова
        private static string GetFullString(string text, string key)
        {
            string fullString;
            if ((text.IndexOf(key) + key.Length) < text.Length)
            {
                if (text.Substring(text.IndexOf(key) + key.Length).Contains(':'))
                {
                    fullString = FirstFromToByKey(text, " ", ": ", key);
                    if (fullString.Contains("\r\n"))
                    {
                        fullString = fullString.Substring(0, fullString.LastIndexOf("\r\n"));
                        return RemChar(fullString, "\r\n");
                    }
                    else return "";
                }
                else
                {
                    fullString = text.Substring(text.IndexOf(key) + key.Length + 1);
                    return RemChar(fullString, "\r\n");
                }
            }
            else return "";
        }        
        
        // Получение транспортной кодировки заголовка
        public static TransferEncoding GetTransferEncoding(string encStr)
        {
            TransferEncoding te;
            switch(encStr)
            {
                case "B":
                    te = TransferEncoding.Base64;
                    break;
                case "b":
                    te = TransferEncoding.Base64;
                    break;
                case "Q":
                    te = TransferEncoding.QuotedPrintable;
                    break;
                case "q":
                    te = TransferEncoding.QuotedPrintable;
                    break;
                case "quoted-printable":
                    te = TransferEncoding.QuotedPrintable;
                    break;
                case "base64":
                    te = TransferEncoding.Base64;
                    break;
                case "7bit":
                    te = TransferEncoding.SevenBit;
                    break;
                default:
                    te = TransferEncoding.EightBit;
                    break;
            }
            return te;

        }

        // Получение типа кодировки из принятой строки
        public static Encoding MyGetEncoding(string encStr)
        {
            Encoding enc;
            switch (encStr)
            {
                case "UTF-8":
                    enc = Encoding.UTF8;
                    break;
                case "utf-8":
                    enc = Encoding.UTF8;
                    break;
                case "koi8-r":
                    enc = Encoding.GetEncoding("koi8r");
                    break;
                case "windows-1251":
                    enc = Encoding.GetEncoding(1251);
                    break;
                default:
                    enc = Encoding.GetEncoding("koi8r");
                    break;
            }
            return enc;
        }

        // Выбор между типами транспортного кодирования
        // и декодирование строки
        public static string GetCodeLine(string line, Encoding enc, TransferEncoding transferCoding)
        {
            string[] lines;
            string temp;
            switch (transferCoding)
            {
                // Декодирование Base64
                case TransferEncoding.Base64:
                    lines = line.Split('=');
                    temp = "";
                    if (lines.Length > 1)
                    {
                        foreach (string l in lines)
                            temp += enc.GetString(Base64DecodeEx(enc.GetBytes(l + "="), null));
                        line = temp;
                    }
                    else line = enc.GetString(Base64DecodeEx(enc.GetBytes(line), null));
                    break;                
                // Декодирование Quoted_Printable
                case TransferEncoding.QuotedPrintable:
                    line = DecodeQuotedPrintable(line, enc);
                    break;
                case TransferEncoding.SevenBit:
                    if(enc == Encoding.GetEncoding("koi8r"))
                        line = enc.GetString(Encoding.Default.GetBytes(line));
                    else line = Decode_IMAP_UTF7_String(line);
                    break;
                case TransferEncoding.EightBit:
                    line = enc.GetString(Encoding.Default.GetBytes(line));
                    break;
            }
            return line;
        }        
        
        // Получение полной декодированной строки
        public static string DecodeString (string text)
        {
            string transferCoding="";
            string decode = "";
            string line = "";
            Encoding enc = Encoding.Default;
            if (text.Contains("=?"))
            {
                while (text.IndexOf("=?") >= 0)
                {
                    if (text.Substring(text.LastIndexOf("=?")).Contains("?="))
                    {
                         
                        string temp = LastFromTo(text, "=?", "?=");
                        string newTransferCoding = FirstFromTo(temp, "?", "?");
                        enc = MyGetEncoding(temp.Substring(0, temp.IndexOf("?")));
                        if (newTransferCoding == transferCoding || line == "")
                        {
                            line = temp.Substring(temp.LastIndexOf("?") + 1) + line;
                            transferCoding = newTransferCoding;
                        }
                        else
                        {
                            decode = GetCodeLine(line, enc, GetTransferEncoding(transferCoding)) + decode;
                            line = temp.Substring(temp.LastIndexOf("?") + 1);
                        }
                    }
                    text = text.Substring(0, text.LastIndexOf("=?"));
                }
                return GetCodeLine(line, enc, GetTransferEncoding(transferCoding)) + decode;
            }
            else return text;
        }
        #endregion
        #endregion

        #region Методы декодирования
        // Метод декодироавния из кодировки "Quoted-Printable"
        private static string DecodeQuotedPrintable(string str, Encoding enc)
        {
            var result = new List<byte>();
            str = str.Replace("=\r\n", "");
            for (int i = 0; i < str.Length; i++)
            {
                char ch = str[i];
                if (ch == '=')
                {
                    result.Add(Convert.ToByte(str.Substring(i + 1, 2), 16));
                    i += 2;
                }
                else result.Add((byte)ch);
            }
            return enc.GetString(result.ToArray());
        }

        // Метод декодирования из модифицированного UTF-7
        private static string Decode_IMAP_UTF7_String(string text)
        {
            // Алфавит модифицированного UTF-7
            char[] base64Chars = new char[]{
                'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
                '0','1','2','3','4','5','6','7','8','9','+',','
            };
            // Исходный текст в массив байтов            
            byte[] encodedBlock = Encoding.Default.GetBytes(text);
            // Декодирование из Base64 с использованием алфавита модифицированного UTF-7						
            byte[] decodedData = Base64DecodeEx(encodedBlock, base64Chars);
            char[] decodedChars = new char[decodedData.Length / 2];
            // Преобразование массива байтов в массив символов юникода
            for (int iC = 0; iC < decodedChars.Length; iC++)
                decodedChars[iC] = (char)(decodedData[iC * 2] << 8 | decodedData[(iC * 2) + 1]);
            string retVal = new string(decodedChars);
            return retVal;
        }

        // Метод декодирования из Base64
        private static byte[] Base64DecodeEx(byte[] base64Data, char[] base64Chars)
        {
            if (base64Chars == null)
            {
                // Алфавит Base64
                base64Chars = new char[]{
                    'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                    'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
                    '0','1','2','3','4','5','6','7','8','9','+','/'
                };
            }

            // Создание таблицы кодировки
            byte[] decodeTable = new byte[128];
            for (int i = 0; i < 128; i++)
            {
                int mappingIndex = -1;
                for (int bc = 0; bc < base64Chars.Length; bc++)
                {
                    if (i == base64Chars[bc])
                    {
                        mappingIndex = bc;
                        break;
                    }
                }
                if (mappingIndex > -1)
                {
                    decodeTable[i] = (byte)mappingIndex;
                }
                else
                {
                    decodeTable[i] = 0xFF;
                }
            }

            byte[] decodedDataBuffer = new byte[((base64Data.Length * 6) / 8) + 4]; // Таблица хранения декодированных байтов
            int decodedBytesCount = 0;      // Количество декодированных байтов
            int nByteInBase64Block = 0;     // Количество исходных байтов в блоке base64Block
            byte[] decodedBlock = new byte[3];  // Трехбайтовый блок для временого хранения декодированных байтов
            byte[] base64Block = new byte[4];   // Четырехбайтовый блок для временого хранения исходных байтов

            for (int i = 0; i < base64Data.Length; i++)
            {
                byte b = base64Data[i];

                // Заполнение блока исходных байтов только символами Base64
                if (b == '=')
                {
                    base64Block[nByteInBase64Block] = 0xFF;
                }
                else
                {
                    byte decodeByte = decodeTable[b & 0x7F];
                    if (decodeByte != 0xFF)
                    {
                        base64Block[nByteInBase64Block] = decodeByte;
                        nByteInBase64Block++;
                    }
                }

                int encodedBytesCount = -1;
                // Количество декодированных байтов в блоке
                if (nByteInBase64Block == 4)
                {
                    encodedBytesCount = 3;
                }

                else if (i == base64Data.Length - 1)
                {
                    if (nByteInBase64Block == 1)
                        encodedBytesCount = 0;
                    else if (nByteInBase64Block == 2)
                        encodedBytesCount = 1;
                    else if (nByteInBase64Block == 3)
                        encodedBytesCount = 2;
                }

                if (encodedBytesCount > -1)
                {
                    // Заполнение декодированного блока
                    decodedDataBuffer[decodedBytesCount + 0] = (byte)((int)base64Block[0] << 2 | (int)base64Block[1] >> 4);
                    decodedDataBuffer[decodedBytesCount + 1] = (byte)(((int)base64Block[1] & 0xF) << 4 | (int)base64Block[2] >> 2);
                    decodedDataBuffer[decodedBytesCount + 2] = (byte)(((int)base64Block[2] & 0x3) << 6 | (int)base64Block[3] >> 0);
                    // Добавление количества декодированных байтов в блоке к общему количеству
                    // декодированных байтов
                    decodedBytesCount += encodedBytesCount;
                    nByteInBase64Block = 0;
                }
            }
            // Отсечение лишних декодированных байтов
            if (decodedBytesCount > -1)
            {
                byte[] retVal = new byte[decodedBytesCount];
                Array.Copy(decodedDataBuffer, 0, retVal, 0, decodedBytesCount);
                return retVal;
            }
            else
            {
                return new byte[0];
            }
        }
        #endregion

        #region Получение полей заголовков
        // Получение темы письма
        public static string GetSubject(string header)
        {
            string fullSubject = GetFullString(header, "\r\nSubject:");
            return DecodeString(fullSubject);
        }

        // Получение транспортрой кодировки тела письма
        public static TransferEncoding GetBodyTransfer(string text)
        {
            string transEnc = GetFullString(text, "\r\nContent-Transfer-Encoding:");
            return GetTransferEncoding(transEnc);
        }
        
        
        // Получение из ответа адреса отправителя
        public static MailAddress GetFrom(string text)
        {
            string fullAddress = GetFullString(text, "\r\nFrom:");
            return GetAddress(fullAddress);
        }

        // Получение из конверта коллекции адресов получателя
        public static MailAddressCollection GetTo(string text)
        {
            string fullAddress = GetFullString(text, "\r\nTo:");
            return GetAddressColection(fullAddress);
        }

        // Получение из конверта коллекции адресов для ответа
        public static MailAddressCollection GetReply(string text)
        {
            string fullAddress = GetFullString(text, "\r\nReply-To:");
            return GetAddressColection(fullAddress);
        }

        // Получение из конверта коллекции адресов копий
        public static MailAddressCollection GetCC(string text)
        {
            string fullAddress = GetFullString(text, "\r\nCc:");
            return GetAddressColection(fullAddress);
        }

        // Получение из конверта коллекции адресов скрытых копий
        public static MailAddressCollection GetBcc(string text)
        {
            string fullAddress = GetFullString(text, "\r\nBCC:");
            return GetAddressColection(fullAddress);
        }

        // Получение типа содержимого письма
        public static ContentType GetContentType(string text)
        {
            ContentType ct;
            string contType = GetFullString(text, "\r\nContent-Type:");
            if (contType.Length != 0)
            {
                ct = new ContentType(contType);
                if (ct.MediaType == "multipart/related")
                    if (contType.Contains("type="))
                        ct.MediaType = FirstFromToByKey(contType, "\"", "\"", "type=");                
            }
            else ct = new ContentType();
            return ct;
        }

        // Получение идентификатора письма
        public static string GetID(string text)
        {
            if (text.Contains("\r\nMessage-ID")) return FirstFromToByKey(text, "<", ">", "\r\nMessage-ID:");
            else return FirstFromToByKey(text, "<", ">", "\r\nMessage-Id:");
        }

        public static ContentDisposition GetDisposition(string text)
        {
            string disposition = FirstFromToByKey(text, " ", ";", "Content-Disposition:");
            if (disposition == "") return null;
            else return new ContentDisposition(disposition);            
        }
        // Получение из конверта даты письма
        public static DateTime GetDate(string text)
        {
            string date = GetFullString(text, "\r\nDate:");
            try
            {
                if (date.Contains("+")) date = date.Substring(0, date.IndexOf('+') + 5);
                return Convert.ToDateTime(date);
            }
            catch(Exception ex)
            {
                return new DateTime();
            }
        }
        #endregion
    }
}
