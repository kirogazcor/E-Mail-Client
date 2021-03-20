using System.Collections.Generic;
using E_Mail_Client.EMailClient;
using System.Net.Mail;
using System.Windows;

namespace E_Mail_Client
{
    public class MyViewControl:DependencyObject
    {
        // Символ отображающийся на кнопки максимального/ нормального размера окна
        public char MaxNorm
        {
            get { return (char)GetValue(MaxNormProperty); }
            set { SetValue(MaxNormProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства MaxNorm
        public static readonly DependencyProperty MaxNormProperty =
            DependencyProperty.Register("MaxNorm", typeof(char), typeof(MyViewControl), new PropertyMetadata('⬜'));

        // Прозрачность фона панели просмотра писем
        public double OpMessBox
        {
            get { return (double)GetValue(OpMessBoxProperty); }
            set { SetValue(OpMessBoxProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства OpMessBox
        public static readonly DependencyProperty OpMessBoxProperty =
            DependencyProperty.Register("OpMessBox", typeof(double), typeof(MyViewControl), new PropertyMetadata(0.0));

        // Заголовок главного окна
        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства Title
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(MyViewControl), new PropertyMetadata("E-mail клиент"));

        // Заголовок окна настроек почтового ящика
        public string SettingTitle
        {
            get { return (string)GetValue(SettingTitleProperty); }
            set { SetValue(SettingTitleProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства SettingTitle
        public static readonly DependencyProperty SettingTitleProperty =
            DependencyProperty.Register("SettingTitle", typeof(string), typeof(MyViewControl), new PropertyMetadata("Настройки ящика"));

        // Список почтовых ящиков
        public List<MailBox> MailBoxList
        {
            get { return (List<MailBox>)GetValue(MailBoxListProperty); }
            set { SetValue(MailBoxListProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства MailBoxList
        public static readonly DependencyProperty MailBoxListProperty =
            DependencyProperty.Register("MailBoxList", typeof(List<MailBox>), typeof(MyViewControl), new PropertyMetadata(null));

        
        // Текущий почтовый ящик
        public MailBox CurrentBoxNum
        {
            get { return (MailBox)GetValue(CurrentBoxNumProperty); }
            set { SetValue(CurrentBoxNumProperty, value); }
        }

        // Использование DependencyProperty в качестве резервного хранилища для свойства CurrentBoxNum
        public static readonly DependencyProperty CurrentBoxNumProperty =
            DependencyProperty.Register("CurrentBoxNum", typeof(MailBox), typeof(MyViewControl), new PropertyMetadata(null));
        
        // Имя выбранного каталога
        public Folder SelFolder
        {
            get { return (Folder)GetValue(SelFolderProperty); }
            set { SetValue(SelFolderProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства SelFolder
        public static readonly DependencyProperty SelFolderProperty =
            DependencyProperty.Register("SelFolder", typeof(Folder), typeof(MyViewControl), new PropertyMetadata(null));

        // Отображаемое имя почтового ящика
        public string NameBox
        {
            get { return (string)GetValue(NameBoxProperty); }
            set { SetValue(NameBoxProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства NameBox
        public static readonly DependencyProperty NameBoxProperty =
            DependencyProperty.Register("NameBox", typeof(string), typeof(MyViewControl), new PropertyMetadata(null));

        // Адрес текущего или создаваемого почтового ящика
        public string Address
        {
            get { return (string)GetValue(AddressProperty); }
            set { SetValue(AddressProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства Address
        public static readonly DependencyProperty AddressProperty =
            DependencyProperty.Register("Address", typeof(string), typeof(MyViewControl), new PropertyMetadata(null));
        
        // Настройки текущего или создаваемого почтового ящика
        public SettingsClass Settings
        {
            get { return (SettingsClass)GetValue(SettingsProperty); }
            set { SetValue(SettingsProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства Settings
        public static readonly DependencyProperty SettingsProperty =
            DependencyProperty.Register("Settings", typeof(SettingsClass), typeof(MyViewControl), new PropertyMetadata(null));

        // Выбранное письмо для просмотра
        public MyMailMessage Message
        {
            get { return (MyMailMessage)GetValue(MessageProperty); }
            set { SetValue(MessageProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства Message
        public static readonly DependencyProperty MessageProperty =
            DependencyProperty.Register("Message", typeof(MyMailMessage), typeof(MyViewControl), new PropertyMetadata(null));
        
        // Путь к просматриваемому html письму
        public string HtmlPath
        {
            get { return (string)GetValue(HtmlPathProperty); }
            set { SetValue(HtmlPathProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства HtmlPath
        public static readonly DependencyProperty HtmlPathProperty =
            DependencyProperty.Register("HtmlPath", typeof(string), typeof(MyViewControl), new PropertyMetadata(null));

        // Основной получатель
        public string HeaderTo
        {
            get { return (string)GetValue(HeaderToProperty); }
            set { SetValue(HeaderToProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства HtmlPath
        public static readonly DependencyProperty HeaderToProperty =
            DependencyProperty.Register("HeaderTo", typeof(string), typeof(MyViewControl), new PropertyMetadata(null));

        // Письмо для отправки
        public MailMessage SendMessage
        {
            get { return (MailMessage)GetValue(SendMessageProperty); }
            set { SetValue(SendMessageProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства SendMessage
        public static readonly DependencyProperty SendMessageProperty =
            DependencyProperty.Register("SendMessage", typeof(MailMessage), typeof(MyViewControl), new PropertyMetadata(null, SendMessageChanged));
        // Метод вызываемый при изменении свойства SendMessage
        private static void SendMessageChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            MyViewControl current = d as MyViewControl;
            if(current != null)
            {
                current.To = current.SendMessage.To.ToString();
            }
        }

        // Список получателей письма получатель
        public string To
        {
            get { return (string)GetValue(ToProperty); }
            set { SetValue(ToProperty, value); }
        }
        // Использование DependencyProperty в качестве резервного хранилища для свойства To
        public static readonly DependencyProperty ToProperty =
            DependencyProperty.Register("To", typeof(string), typeof(MyViewControl), new PropertyMetadata(null));

    }
    
}
