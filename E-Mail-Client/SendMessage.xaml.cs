using System;
using System.Windows;
using System.Windows.Controls;
using E_Mail_Client.EMailClient;
using System.Net.Mail;
using Microsoft.Win32;

namespace E_Mail_Client
{
    /// <summary>
    /// Логика взаимодействия для SendMessage.xaml
    /// </summary>
    public partial class SendMessage : Window
    {
        private MailMessage myMessage;  // Отправляемое письмо

        public MailMessage MyMessage
        {
            get { return myMessage; }
        }

        #region Конструкторы
        public SendMessage()
        {
            InitializeComponent();
        }
        public SendMessage(MailBox currentBox) :this()
        {
            myMessage = new MailMessage(currentBox.MyAddress, currentBox.MyAddress);
            DataContext = new MyViewControl();
            ((MyViewControl)DataContext).SendMessage = myMessage;
        }
        #endregion

        // Обработка нажатия кнопки отправить
        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Формирование письма для отправки
                myMessage = ((MyViewControl)DataContext).SendMessage;
                myMessage.To.Clear();
                myMessage.To.Add(((MyViewControl)DataContext).To);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Введите корректные настройки. " + ex.Message);
                DialogResult = false;
            }
        }

        // Обработка нажатия кнопки добавления файла
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            if(openFile.ShowDialog() == true)
            {
                // Прикрепление файла
                ((MyViewControl)DataContext).
                    SendMessage.Attachments.Add(new Attachment(openFile.FileName));                
                listBox.Items.Refresh();
            }            
        }

        // Обработка щелчка мыши по значку вложения
        // (удаление вложения из письма)
        private void AttachBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((ListBox)sender).SelectedIndex >= 0)
            {
                // Удаление вложения из списка
                ((MyViewControl)DataContext).SendMessage.Attachments.
                    RemoveAt(((ListBox)sender).SelectedIndex);
                listBox.Items.Refresh();
                // Снятие выделения с элементов ListBox
                ((ListBox)sender).SelectedIndex = -1;
            }
        }
    }
}
