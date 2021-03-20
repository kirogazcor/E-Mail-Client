using System;
using System.Windows;
using E_Mail_Client.EMailClient;
using System.Net.Mail;

namespace E_Mail_Client
{
    /// <summary>
    /// Логика взаимодействия для CreatBoxWindow.xaml
    /// </summary>
    public partial class CreatBoxWindow : Window
    {
        private MailBox myBox;  // Выбранный или новый почтовый ящик

        // Конструктор окна с параметрами названия окна и почтового ящика
        public CreatBoxWindow(string title, MailBox currentBox)
        {
            InitializeComponent();            
            DataContext = new MyViewControl();
            myBox = new MailBox(currentBox);
            ((MyViewControl)DataContext).NameBox = myBox.Name;
            ((MyViewControl)DataContext).SettingTitle = title + " " + myBox.Name;
            ((MyViewControl)DataContext).Address = myBox.MyAddress.Address;
            ((MyViewControl)DataContext).Settings = myBox.Settings;
            password.Password = myBox.Settings.Rassword;            
        }

        // Обработчик нажатия кнопки Сохранить
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка наличия значений в полях ввода
            if (((MyViewControl)DataContext).NameBox == null)
            {
                DialogResult = false;
                MessageBox.Show("Введите название учетной записи");
            }
            else if (((MyViewControl)DataContext).Address == null)
            {
                DialogResult = false;
                MessageBox.Show("Введите адрес учетной записи");
            }
            else if (password.Password == null)
            {
                DialogResult = false;
                MessageBox.Show("Введите пароль");
            }
            else if (((MyViewControl)DataContext).Settings.ImapServer == null)
            {
                DialogResult = false;
                MessageBox.Show("Введите адрес сервера входящей почты");
            }
            else if (((MyViewControl)DataContext).Settings.SmtpServer == null)
            {
                DialogResult = false;
                MessageBox.Show("Введите адрес сервера исходящей почты");
            }
            else
            {
                try
                {
                    // Запись настроек из полей ввода в объект почтового ящика
                    myBox.MyAddress = new MailAddress(((MyViewControl)DataContext).Address);
                    myBox.Settings.Rassword = password.Password;
                    myBox.Name = ((MyViewControl)DataContext).NameBox;
                    DialogResult = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Введите корректные настройки. " + ex.Message);
                }                
            }
        }

        // Свойство почтовый ящик
        public MailBox MyBox
        { get { return myBox; } }
    }
}
