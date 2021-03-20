using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using E_Mail_Client.EMailClient;
using System.IO;
using Microsoft.Win32;

namespace E_Mail_Client
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private EmailClient Client; // Почтовый клиент
        
        // Конструктор окна
        public MainWindow()
        {
            InitializeComponent();
            // Регистрация класса свойств зависимости
            DataContext = new MyViewControl();
            // Создание файла для временного хранения содержимого писем в формате html
            using (FileStream fStream = new FileStream("temp.html", FileMode.Create, FileAccess.Write))
            {
                // Получение адреса файла для временного хранения содержимого писем в формате html
                ((MyViewControl)DataContext).HtmlPath = "file:///" + Environment.CurrentDirectory + "\\temp.html";
            }
            // Создание объекта почтового клиента
            Client = new EmailClient();
            // Регистрация событий клиента
            Client.OnException += Client_OnException;
            Client.OnViewListMess += Client_OnViewListMess;
            Client.OnViewMess += Client_OnViewMess;
        }        

        #region Открытие и закрытие окна
        // Обработка загрузки окна
        private void EmailWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Загрузка списка почтовых ящиков из файла
            Client.MailBoxList = MailStorage.LoadMailBoxList();
            ((MyViewControl)DataContext).MailBoxList = Client.MailBoxList;
        }

        // Обработка закрытия окна
        private void EmailWindow_Closed(object sender, EventArgs e)
        {
            // Отключение от сервера
            Client.Disconnect();
        }
        #endregion

        #region Изменение размеров и перетаскивание окна
        // Обработка нажатия кнопки закрытия окна
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Обработка нажатия кнопки максимизации и нормализации размера окна
        private void MaxNormButton_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
                ((MyViewControl)DataContext).MaxNorm = '⬜';
            }
            else
            {
                WindowState = WindowState.Maximized;
                ((MyViewControl)DataContext).MaxNorm = '⧉';
            }
        }

        // Обработка нажатия кнопки сворачивания окна
        private void MinButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;

        }

        // Обработка перетаскивания окна
        private void EmailWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        #endregion
        
        // Обработка создания почтового ящика
        private void CreatBoxButton_Click(object sender, RoutedEventArgs e)
        {
            // Запуск диалогового окна для ввода настроек почтового ящика
            CreatBoxWindow NewWindow = new CreatBoxWindow("Новая учетная запись", new MailBox())
            {
                Owner = this
            };
            if (NewWindow.ShowDialog() == true)
            {
                // Проверка, что ящик с таким адресом ещё не существует
                string newAddress = NewWindow.MyBox.MyAddress.Address;
                if (!Client.ConsistAddress(newAddress))
                {
                    try
                    {
                        Client.MailBoxList.Add(NewWindow.MyBox);
                        // Добавление нового почтового ящика в список
                        MailBox mb = Client.MailBoxList.Last();
                        // Визуализация добавленного ящика
                        if (mb != null) ((MyViewControl)DataContext).Title = mb.Name;
                        ((MyViewControl)DataContext).CurrentBoxNum = mb;
                        // Изменение текущего почтового ящика
                        Client.CurrentMailBox = mb;
                        // Добавление нового ящика в файл
                        MailStorage.SaveNewMailBox(mb);
                        listMailBox.Items.Refresh();
                        // Очистка списка писем
                        ((MyViewControl)DataContext).SelFolder = null;
                        // Очистка окна просмотра письма
                        ((MyViewControl)DataContext).Message = null;
                        ((MyViewControl)DataContext).OpMessBox = 0;
                        // Загрузить список папок
                        Client.LoadFolderList();                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else MessageBox.Show("Учетная запись с адресом " + newAddress + " уже существует");
            }
        }

        // Обработка удаления почтового ящика
        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
           
            if (Client.CurrentMailBox != null)
            {
                if (MessageBox.Show("Удалить учетную запись " + Client.CurrentMailBox.Name + "?","Внимание!!!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    // Удаление ящика из файла
                    MailStorage.RemoveMailBox(Client.CurrentMailBox.MyAddress.Address);
                    // Удаление ящика из списка в оперативной памяти
                    Client.MailBoxList.Remove(Client.CurrentMailBox);
                    // Обнуление текущего ящика
                    Client.CurrentMailBox = null;
                    // Визуализация удаления
                    ((MyViewControl)DataContext).CurrentBoxNum = null;
                    ((MyViewControl)DataContext).Title = "E-mail клиент";
                    listMailBox.Items.Refresh();
                    // Очистка списка писем
                    ((MyViewControl)DataContext).SelFolder = null;
                    // Очистка окна просмотра письма
                    ((MyViewControl)DataContext).Message = null;
                    ((MyViewControl)DataContext).OpMessBox = 0;
                }
            }
            else MessageBox.Show("Выберите учетную запись");
        }

        // Обработка вызова окна настроек почтового ящика 
        private void SettingButton_Click(object sender, RoutedEventArgs e)
        {
            if (Client.CurrentMailBox != null)
            {
                // Запуск диалогового окна для ввода настроек почтового ящика
                CreatBoxWindow NewWindow = new CreatBoxWindow("Настройки учетной записи", Client.CurrentMailBox)
                {
                    Owner = this
                };
                if (NewWindow.ShowDialog() == true)
                {
                    try
                    {
                        // Замена текущего ящика на ящик с новыми настройками
                        MailBox newMb = new MailBox(NewWindow.MyBox);
                        Client.MailBoxList[Client.MailBoxList.IndexOf(Client.CurrentMailBox)] = newMb;
                        // Визуализация ящика
                        ((MyViewControl)DataContext).CurrentBoxNum = newMb;
                        Client.CurrentMailBox = newMb;
                        // Сохранение настроек почтового ящика в файл
                        MailStorage.SaveSettings(newMb);
                        listMailBox.Items.Refresh();
                        // Очистка списка писем
                        ((MyViewControl)DataContext).SelFolder = null;
                        // Очистка окна просмотра письма
                        ((MyViewControl)DataContext).Message = null;
                        ((MyViewControl)DataContext).OpMessBox = 0;
                        // Загрузить список папок
                        Client.LoadFolderList();                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else MessageBox.Show("Выберите учетную запись");
        }

        // Обработка выбора почтового ящика
        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((ListBox)sender).SelectedIndex >= 0)
            {
                // Сворачивание расширителя
                MailBoxExp.IsExpanded = false;
                try
                {
                    // Отключить соединение
                    Client.Disconnect();
                    MailBox mb = Client.MailBoxList[((ListBox)sender).SelectedIndex];
                    if (mb != null)
                    {
                        // Загрузка настроек почтового ящика из файла
                        mb.Settings = MailStorage.LoadSettingBox(mb.MyAddress.Address);
                        // Визуализация выбранного ящика
                        ((MyViewControl)DataContext).Title = mb.Name;
                        ((MyViewControl)DataContext).CurrentBoxNum = mb;
                        Client.CurrentMailBox = mb;
                        // Очистка списка писем
                        ((MyViewControl)DataContext).SelFolder = null;
                        // Очистка окна просмотра письма
                        ((MyViewControl)DataContext).Message = null;
                        ((MyViewControl)DataContext).OpMessBox = 0;
                        // Загрузить список папок
                        Client.LoadFolderList();
                        
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        // Обработка разворачивания списка папок
        private void FolderExp_Expanded(object sender, RoutedEventArgs e)
        {
            if (Client.CurrentMailBox != null)
            {
                // Визуализация папок текущего ящика
                ((MyViewControl)DataContext).CurrentBoxNum = new MailBox(Client.CurrentMailBox);
            }
        }

        #region Выбор папки
        // Обработка выбора папки
        private void FolderListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((ListBox)sender).SelectedIndex >= 0)
            {
                // Сворачивание расширителя
                FolderExp.IsExpanded = false;
                try
                {
                    MailBox mb = new MailBox(Client.CurrentMailBox);
                    Folder f = new Folder(Client.CurrentMailBox.Folders[((ListBox)sender).SelectedIndex]);
                    // Визуализация выбранной папки
                    if (f != null && mb != null) ((MyViewControl)DataContext).Title = mb.Name + " - " + f.Name;
                    ((MyViewControl)DataContext).SelFolder = f;
                    Client.CurrentMailBox.SelectedFolder = f;
                    // Загрузка списка писем
                    Client.LoadMessageList();
                    // Очистка окна просмотра письма
                    ((MyViewControl)DataContext).Message = null;
                    ((MyViewControl)DataContext).OpMessBox = 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        // Обработка отображения списка писем
        private void Client_OnViewListMess()
        {
            Dispatcher.Invoke(() =>
            {
                // Очистка окна просмотра письма
                ((MyViewControl)DataContext).Message = null;
                ((MyViewControl)DataContext).OpMessBox = 0;
                // Показ списка писем
                ((MyViewControl)DataContext).SelFolder = new Folder(Client.CurrentMailBox.SelectedFolder);
                listMessage.Items.Refresh();
            });
        }
        #endregion

        #region Выбор письма
        // Выбор письма для просмотра
        private void ListMessage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((ListBox)sender).SelectedIndex >= 0)
            {                
                try
                {
                    MyMailMessage mess = new MyMailMessage(" " + Client.CurrentMailBox.SelectedFolder.Messages[((ListBox)sender).SelectedIndex].Num.ToString()
                                                            + " " + Client.CurrentMailBox.SelectedFolder.Messages[((ListBox)sender).SelectedIndex].Headers.Get("main"));
                    // Визуализация выбранной папки                    
                    Client.CurrentMailBox.SelectedFolder.Message = mess;
                    // Очистка окна просмотра письма
                    ((MyViewControl)DataContext).Message = null;
                    ((MyViewControl)DataContext).OpMessBox = 0;
                    // Загрузка письма с сервера
                    Client.LoadMessage();
                    listMessage.Items.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }                
            }
        }

        // Обработка отображения письма
        private void Client_OnViewMess()
        {
            Dispatcher.Invoke(() =>
            {
                if (((MyViewControl)DataContext).Message == null)
                    ((MyViewControl)DataContext).Message = new MyMailMessage(Client.CurrentMailBox.SelectedFolder.Message);
                ((MyViewControl)DataContext).HeaderTo = ((MyViewControl)DataContext).Message.To.First().ToString();
                if (Client.CurrentMailBox.SelectedFolder.Message.IsBodyHtml)
                {
                    // Поучение ссылки на окно просмотра html
                    Label labelBox = (Label)labelMessage.Template.FindName("labelBox", labelMessage);
                    Frame frm = (Frame)labelBox.Content;
                    // Сохранение идентификатора источника окна html
                    // на время изменения файла источника
                    Uri temp = frm.Source;
                    frm.Source = null;
                    try
                    {
                        // Загрузка в файл содержимое письма в формате html
                        using (StreamWriter sw = new StreamWriter("temp.html", false, Encoding.Unicode))
                        {
                            sw.Write(Client.CurrentMailBox.SelectedFolder.Message.Body);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    // Изменение прозрачности панели просмотра писем
                    frm.Source = temp;
                    // Обновление окна просмотра  html
                    frm.Refresh();
                    listMessage.Items.Refresh();
                }
                ((MyViewControl)DataContext).OpMessBox = 1;
            });
        }
        #endregion

        // Выбор вложения для сохранения
        private void ListBoxAttach_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((ListBox)sender).SelectedIndex >= 0)
            {
                // Выбор имени файла для сохранения
                string filename = ((MyViewControl)DataContext).Message.MyAttachments[((ListBox)sender).SelectedIndex].Name;
                SaveFileDialog saveFileDialog = new SaveFileDialog()
                {
                    FileName = filename
                };
                if (saveFileDialog.ShowDialog() == true)
                {
                    FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                    byte[] content = MailStorage.LoadAttach(Client.CurrentMailBox, filename);
                    if (content.Length > 0)
                        stream.Write(content, 0, content.Length);
                    // Очистка памяти после скачивания
                    content = null;
                    stream.Dispose();
                }
            // Снятие выделения элемента списка
            ((ListBox)sender).SelectedIndex = -1;
            }
        }               

        // Обработка нажатия кнопки удаления письма
        private void RemoveMessage_Click(object sender, RoutedEventArgs e)
        {
            Client.DeleteMessage();
        }

        // Обновление состояния почтового ящика
        private void Update_Click(object sender, RoutedEventArgs e)
        {
            if (Client.CurrentMailBox != null)
            {
                try
                {
                    ((MyViewControl)DataContext).Title = Client.CurrentMailBox.Name;
                    // Очистка списка писем
                    ((MyViewControl)DataContext).SelFolder = null;
                    // Очистка окна просмотра письма
                    ((MyViewControl)DataContext).Message = null;
                    ((MyViewControl)DataContext).OpMessBox = 0;
                    // Загрузить список папок
                    Client.LoadFolderList();                    
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else MessageBox.Show("Выберите учетную запись");
        }

        // Обработка вызова окна создания нового письма
        private void Write_Click(object sender, RoutedEventArgs e)
        {
            if (Client.CurrentMailBox != null)
            {
                // Запуск диалогового окна создания исходящего письма
                SendMessage NewWindow = new SendMessage(Client.CurrentMailBox)
                {
                    Owner = this
                };                
                try
                {
                    if (NewWindow.ShowDialog() == true)
                    {
                        // Отправка письма
                        Client.SentMessage(NewWindow.MyMessage);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else MessageBox.Show("Выберите учетную запись");
        }

        // Обработка возникновения асинхронных ошибок в работе клиента
        private void Client_OnException(Exception ex)
        {
            MessageBox.Show(ex.Message);
        }        
    }
}
