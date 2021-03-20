using System;
using System.Text;
using System.Net.Sockets;
using System.Net.Security;
using System.IO;

namespace E_Mail_Client.EMailClient
{
    // Класс TCP-клиента
    class TCPClient : TcpClient
    {

        private int bufferSize = 1024;
        private Stream _tcpStream = null;
        protected bool busyStream = false;        

        // Конструктор с параметрами
        public TCPClient(string host, int port, bool ssl) : base(host, port)
        {
            bufferSize = 1024;
            try
            {
                if (ssl)
                {
                    _tcpStream = new SslStream(GetStream());
                    ((SslStream)_tcpStream).AuthenticateAsClient(host);
                }
                else _tcpStream = GetStream();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Отключение от TCP-сервера
        public void Disconnect()
        {
            if (_tcpStream != null)
                _tcpStream.Dispose();
            if (Connected)
                Close();
        }

        // Отправка запроса на сервер
        protected void Write(string request)
        {
            try
            {
                byte[] sBytes = Encoding.Default.GetBytes(request);
                _tcpStream.Write(sBytes, 0, sBytes.Length);
                _tcpStream.Flush();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Прием ответа с сервера
        protected string Read(string key)
        {
            string s = "";
            StringBuilder sb = new StringBuilder();
            int bytesRead = 0;
            byte[] buffer = new byte[bufferSize];
            try
            {
                int i = 0;
                do
                {
                    // Считывание с сервера строк, пока в строке не появится ключевое слово
                    // или длина строки не достигнет миллиона символов
                    i++;
                    bytesRead = _tcpStream.Read(buffer, 0, bufferSize);
                    sb.Append(Encoding.Default.GetString(buffer));
                    s = sb.ToString();
                } while (!(s.Contains(key) || s.Length > 1000000));
                if(s.Contains(key))_tcpStream.Flush();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return sb.ToString();
        }
    }
}
