using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration.Install;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Data.Linq;
using System.Threading;
using System.IO.MemoryMappedFiles;
using System.Data.SqlClient;
using System.Data.Linq.SqlClient;
using System.Globalization;
using Microsoft.Win32;

namespace DataServiceApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private OpenFileDialog openFile = new OpenFileDialog();
        private OleDbConnection connection = new OleDbConnection();
        private OleDbConnection dataConnection = new OleDbConnection();
        private OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
        private OleDbDataAdapter dataTemperatureAdapter = new OleDbDataAdapter();
        private DataSet dataTemperatureSet = new DataSet();
        private OleDbCommand command = new OleDbCommand();
        private DataContext db = null;
        private static DataSet dataSet = new DataSet();
        private static List<int> indexes = new List<int>();
        public static Dictionary<string, DataTable> dataTables = new Dictionary<string, DataTable>();
        private int currentIndex = 0;
        System.Windows.Forms.NotifyIcon icon = new System.Windows.Forms.NotifyIcon();
        System.Windows.Forms.ContextMenu m_menu = new System.Windows.Forms.ContextMenu();
        bool windowsStart = false;
        bool serviceStart = false;
        bool trayMinimize = false;

        static string[] path = new string[2];

        public MainWindow()
        {
            InitializeComponent();

            icon.Icon = new System.Drawing.Icon("kumz.ico");
            icon.Click += Icon_Click;
            m_menu.MenuItems.Add(0,
                new System.Windows.Forms.MenuItem("Показать", new System.EventHandler(Show_Click)));
            m_menu.MenuItems.Add(1,
                new System.Windows.Forms.MenuItem("Запустить службу", new System.EventHandler(Start_Click)));
            m_menu.MenuItems.Add(2,
                new System.Windows.Forms.MenuItem("Остановить службу", new System.EventHandler(Pause_Click)));
            m_menu.MenuItems.Add(3,
                new System.Windows.Forms.MenuItem("Параметры запуска", new System.EventHandler(Settings_Click)));
            m_menu.MenuItems.Add(4,
                new System.Windows.Forms.MenuItem("Выход", new System.EventHandler(Exit_Click)));
            icon.ContextMenu = m_menu;
            this.ShowInTaskbar = true;



            //Открытие последнего путя до папки с таблицами
            try
            {
                StreamReader sr = new StreamReader("DatabasePath");
                StreamReader srf = new StreamReader("Frequency");
                path[0] = sr.ReadLine();
                path[1] = srf.ReadLine();
                sr.Close();
                srf.Close();
                srf.Dispose();
                sr.Dispose();
                tbPath.Text = path[0];
                tbFreq.Text = path[1];
            }
            catch(Exception ex)
            {

            }

            try
            {
                StreamReader srTray = new StreamReader("trayMinimize");
                StreamReader srServicestart = new StreamReader("serviceStart");
                StreamReader srWindowsstart = new StreamReader("windowsStart");
                trayMinimize = Convert.ToBoolean(srTray.ReadLine());
                serviceStart = Convert.ToBoolean(srServicestart.ReadLine());
                windowsStart = Convert.ToBoolean(srWindowsstart.ReadLine());
                srTray.Close();
                srServicestart.Close();
                srWindowsstart.Close();
                srTray.Dispose();
                srServicestart.Dispose();
                srWindowsstart.Dispose();
                cbTrayMinimize.IsChecked = trayMinimize;
                cbServiceStart.IsChecked = serviceStart;
                cbAppWindows.IsChecked = windowsStart;
            }
            catch
            {

            }

            if (trayMinimize)
            {
                icon.Visible = true;
            }
            else
            {
                icon.Visible = false;
            }

            if (windowsStart)
            {
                SetAutorunValue(true);
            }
            else
            {
                SetAutorunValue(false);
            }
            if (serviceStart)
            {
                try
                {
                    StartService("Служба интеграции");
                }
                catch
                {
                    MessageBox.Show("Возникла ошибка при запуске службы. Проверьте параметры запуска.");
                }
            }
            openFile.FileName = "../../../DataService/bin/Debug/DataService.exe";
            var serviceExists = ServiceController.GetServices().Any(s => s.DisplayName == "Служба интеграции");
            ServiceController service = ServiceController.GetServices().FirstOrDefault(s => s.DisplayName == "Служба интеграции");
            if (!serviceExists)
            {
                btnDelete.IsEnabled = false;
                btnInstall.IsEnabled = true;
                btnPause.IsEnabled = false;
                btnReStart.IsEnabled = false;
                btnStart.IsEnabled = false;
                stateSettingsTab.Visibility = Visibility.Hidden;
                txtCurrentState.Text = "Текущее состояние службы:\n Служба не установлена";
                txtCurrentState.Foreground = Brushes.Red;
                m_menu.MenuItems[3].Enabled=false;
                
            }
            else
            {
                stateSettingsTab.Visibility = Visibility.Visible;
                m_menu.MenuItems[3].Enabled = true;
                if (service.Status == ServiceControllerStatus.Running)
                {
                    txtCurrentState.Text = "Текущее состояние службы:\n Служба запущена";
                    txtCurrentState.Foreground = Brushes.Green;
                    btnStart.IsEnabled = false;
                }

                if (service.Status == ServiceControllerStatus.Stopped)
                {
                    txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                    txtCurrentState.Foreground = Brushes.Yellow;
                    btnPause.IsEnabled = false;
                }

                btnDelete.IsEnabled = true;
                btnInstall.IsEnabled = false;
                service.Dispose();
            }
            
        }

        public bool SetAutorunValue(bool autorun)
        {
            string ExePath = System.Windows.Forms.Application.ExecutablePath;
            RegistryKey reg;
            reg = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run\\");
            try
            {
                if (autorun)
                    reg.SetValue("DataServiceApp", ExePath);
                else
                    reg.DeleteValue("DataServiceApp");

                reg.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }

        //Выбор папки с температурами по индексу оборудования
        public string selectFolder(int index)
        {
            string hex = index.ToString("X8");
            string pathToFolder = $"{path[0]}\\D0000\\DT{hex}";

            return pathToFolder;
        }


        //Подсоединение к базе
        private OleDbConnection OpenConnection(OleDbConnection connection, string _path)
        {
            var builder = new OleDbConnectionStringBuilder();

            builder.Add("Provider", "Microsoft.Jet.OLEDB.4.0");
            builder.Add("Data Source", _path);
            builder.Add("Persist Security Info", "True");
            builder.Add("Extended properties", "Paradox 7.x; HDR=YES");
            builder.Add("Jet OLEDB:Database Password", "jIGGAe");

            connection.ConnectionString = builder.ToString();

            try
            {
                connection.Open();
                return connection;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }

        //Выбор папки с таблицами
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog();

            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbPath.Text = folderDialog.SelectedPath;
            }
        }

        //Запись путя до папки в файл
        private void btnApply_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            if (service.Status != ServiceControllerStatus.Running)
            {
                try
                {
                    path[0] = tbPath.Text;
                    if (connection != null)
                    {
                        connection.Close();
                    }
                    dataTables.Clear();
                    indexes.Clear();
                    File.WriteAllText("DatabasePath", path[0]);
                    StopService("Служба интеграции");
                    GC.Collect();
                    txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                    txtCurrentState.Foreground = Brushes.Yellow;
                    btnStart.IsEnabled = true;
                    btnPause.IsEnabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла ошибка, возможно, данного пути не существует. \nПодробно:\n" + ex.Message.ToString());
                }
            }
            else
            {
                if (MessageBox.Show("Чтобы выбрать новый путь, необходимо остановить службу. Остановить службу и поменять путь до базы данных?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        path[0] = tbPath.Text;
                        if (connection != null)
                        {
                            connection.Close();
                        }
                        dataTables.Clear();
                        indexes.Clear();
                        File.WriteAllText("DatabasePath", path[0]);
                        StopService("Служба интеграции");
                        GC.Collect();
                        txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                        txtCurrentState.Foreground = Brushes.Yellow;
                        btnStart.IsEnabled = true;
                        btnPause.IsEnabled = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Возникла ошибка, возможно, данного пути не существует. \nПодробно:\n" + ex.Message.ToString());
                    }
                }
            }

        }

        //Запись частоты обновления в файл
        private void btnApplyFreq_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            if (service.Status != ServiceControllerStatus.Running)
            {
                try
                {
                    path[1] = tbFreq.Text;
                    if (connection != null)
                    {
                        connection.Close();
                    }
                    dataTables.Clear();
                    indexes.Clear();
                    File.WriteAllText("Frequency", path[1]);
                    StopService("Служба интеграции");
                    GC.Collect();
                    txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                    txtCurrentState.Foreground = Brushes.Yellow;
                    btnStart.IsEnabled = true;
                    btnPause.IsEnabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла ошибка, возможно, вы ввели неподходящие данные. \nПодробно:\n" + ex.Message.ToString());
                }
            }
            else
            {
                if (MessageBox.Show("Чтобы выбрать новую частоту, необходимо остановить службу. Остановить службу и поменять частоту?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        path[1] = tbFreq.Text;
                        if (connection != null)
                        {
                            connection.Close();
                        }
                        dataTables.Clear();
                        indexes.Clear();
                        File.WriteAllText("Frequency", path[1]);
                        StopService("Служба интеграции");
                        GC.Collect();
                        txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                        txtCurrentState.Foreground = Brushes.Yellow;
                        btnStart.IsEnabled = true;
                        btnPause.IsEnabled = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Возникла ошибка, возможно, вы ввели неподходящие данные. \nПодробно:\n" + ex.Message.ToString());
                    }
                }
            }
        }

        //Установка службы
        private void btnInstall_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            try
            {
                ManagedInstallerClass.InstallHelper(new[] {@"/", openFile.FileName });
            }
           catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            MessageBox.Show("Установка службы выполнена!");
            txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
            txtCurrentState.Foreground = Brushes.Yellow;
            btnDelete.IsEnabled = true;
            btnInstall.IsEnabled = false;
            btnPause.IsEnabled = false;
            btnReStart.IsEnabled = true;
            btnStart.IsEnabled = true;
            stateSettingsTab.Visibility = Visibility.Visible;
            m_menu.MenuItems[3].Enabled = true;
        }

        //Удаление службы
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы точно хотите удалить службу?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                try
                {
                    ManagedInstallerClass.InstallHelper(new[] { @"/u", openFile.FileName });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return;
                }
                MessageBox.Show("Служба удалена успешно!");
                txtCurrentState.Text = "Текущее состояние службы:\n Служба не установлена";
                txtCurrentState.Foreground = Brushes.Red;
                btnInstall.IsEnabled = true;
                btnDelete.IsEnabled = false;
                btnPause.IsEnabled = false;
                btnReStart.IsEnabled = false;
                btnStart.IsEnabled = false;
                stateSettingsTab.Visibility = Visibility.Hidden;
                m_menu.MenuItems[3].Enabled = false;
            }
            else
            {
                return;
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                StartService("Служба интеграции");
                txtCurrentState.Text = "Текущее состояние службы:\n Служба запущена";
                txtCurrentState.Foreground = Brushes.Green;
                btnStart.IsEnabled = false;
                btnPause.IsEnabled = true;
                MessageBox.Show("Служба была успешно запущена!");
            }
            catch
            {
                MessageBox.Show("Возникла ошибка. Возможно, вы не указали параметры запуска либо служба не установлена на компьютере.");
                Environment.Exit(1);
            }
            
        }

        private void btnPause_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            if (service.Status == ServiceControllerStatus.Stopped)
            {
                System.Windows.Forms.MessageBox.Show("Служба уже остановлена!");
            }
            else
            {
                StopService("Служба интеграции");
                txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                txtCurrentState.Foreground = Brushes.Yellow;
                btnStart.IsEnabled = true;
                btnPause.IsEnabled = false;
                MessageBox.Show("Служба была успешно остановлена!");
            }
        }

        private void btnReStart_Click(object sender, RoutedEventArgs e)
        {
            RestartService("Служба интеграции");
            btnStart.IsEnabled = false;
            btnPause.IsEnabled = true;
        }

        public static async void StartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Проверяем не запущена ли служба
            if (service.Status != ServiceControllerStatus.Running)
            {
                // Запускаем службу
                service.MachineName = @"DESKTOP-KORPNPP";
                
                service.Start(path);

                // В течении минуты ждём статус от службы
                service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromMinutes(1));
                
                
            }
            else
            {
                MessageBox.Show("Служба уже запущена!");
            }
        }

        // Останавливаем службу
        public static void StopService(string serviceName)
        {
            
            try
            {
                ServiceController service = new ServiceController(serviceName);
                // Если служба не остановлена
                if (service.Status != ServiceControllerStatus.Stopped)
                {
                    // Останавливаем службу
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMinutes(1));

                }
            }
            catch(Exception)
            {

            }
            
        }

        // Перезапуск службы
        public void RestartService(string serviceName)
        {

            ServiceController service = new ServiceController(serviceName);
            TimeSpan timeout = TimeSpan.FromMinutes(1);
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                txtCurrentState.Text = "Текущее состояние службы:\n Перезапуск службы. Останавливаем службу...";
                txtCurrentState.Foreground = Brushes.Orange;
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
            }
            if (service.Status != ServiceControllerStatus.Running)
            {
                txtCurrentState.Text = "Текущее состояние службы:\n Перезапуск службы. Запускаем службу...";
                txtCurrentState.Foreground = Brushes.Yellow;
                // Запускаем службу
                service.Start(path);
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                txtCurrentState.Text = "Текущее состояние службы:\n Служба запущена";
                txtCurrentState.Foreground = Brushes.Green;
                MessageBox.Show("Служба была успешно перезапущена!");
            }
        }

        private void Icon_Click(object sender, EventArgs e)
        {
            if(trayMinimize)
            if (this.WindowState == WindowState.Minimized)
            {
                this.WindowState = WindowState.Normal;
                this.ShowInTaskbar = false;
                this.Activate();
            }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if(trayMinimize)
            {
                if (this.WindowState == WindowState.Minimized)
                {
                    icon.Visible = true;
                }
            }
            
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController service = new ServiceController("Служба интеграции");
                if (service.Status != ServiceControllerStatus.Running)
                {
                    Environment.Exit(1);
                }
                else
                {
                    if (MessageBox.Show("При выходе из программы, служба автоматически прекратит свою работу! Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        StopService("Служба интеграции");
                        Environment.Exit(1);
                    }
                }
            }
            catch(Exception ex)
            {

            }
            
            
        }

        private void Settings_Click(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.ShowInTaskbar = false;
            this.Activate();
            stateSettingsTab.IsSelected = true;
            
        }

        private void Start_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController service = new ServiceController("Служба интеграции");
                // Проверяем не запущена ли служба
                if (service.Status != ServiceControllerStatus.Running)
                {
                    // Запускаем службу
                    service.MachineName = @"DESKTOP-KORPNPP";

                    service.Start(path);

                    // В течении минуты ждём статус от службы
                    service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromMinutes(1));


                }
                else
                {
                    MessageBox.Show("Служба уже запущена!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Служба не установлена");
            }
        }

        private void Pause_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController service = new ServiceController("Служба интеграции");
                if (service.Status == ServiceControllerStatus.Stopped)
                {
                    System.Windows.Forms.MessageBox.Show("Служба уже остановлена!");
                }
                else
                {
                    StopService("Служба интеграции");
                    txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
                    txtCurrentState.Foreground = Brushes.Yellow;
                    btnStart.IsEnabled = true;
                    btnPause.IsEnabled = false;
                    MessageBox.Show("Служба была успешно остановлена!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Служба не установлена");
            }
        }

        private void Show_Click(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.ShowInTaskbar = true;
            this.Activate();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            StopService("Служба интеграции");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(trayMinimize)
            {
                e.Cancel = true;
                this.WindowState = WindowState.Minimized;
            }
            else
            {
                ServiceController service = new ServiceController("Служба интеграции");
                if (service.Status != ServiceControllerStatus.Running)
                {
                    Environment.Exit(1);
                }
                else
                {
                    e.Cancel = true;
                    if (MessageBox.Show("При выходе из программы, служба автоматически прекратит свою работу! Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        StopService("Служба интеграции");
                        Environment.Exit(1);
                    }
                }
            }
        }

        private void cbAppWindows_Checked(object sender, RoutedEventArgs e)
        {
            windowsStart = true;
            SetAutorunValue(true);
            File.WriteAllText("windowsStart", windowsStart.ToString());
        }

        private void cbAppWindows_Unchecked(object sender, RoutedEventArgs e)
        {
            windowsStart = false;
            SetAutorunValue(false);
            File.WriteAllText("windowsStart", windowsStart.ToString());
        }

        private void cbServiceStart_Checked(object sender, RoutedEventArgs e)
        {
            serviceStart = true;
            File.WriteAllText("serviceStart", serviceStart.ToString());
        }

        private void cbServiceStart_Unchecked(object sender, RoutedEventArgs e)
        {
            serviceStart = false;
            File.WriteAllText("serviceStart", serviceStart.ToString());
        }

        private void cbTrayMinimize_Checked(object sender, RoutedEventArgs e)
        {
            trayMinimize = true;
            icon.Visible = true;
            File.WriteAllText("trayMinimize", trayMinimize.ToString());
        }

        private void cbTrayMinimize_Unchecked(object sender, RoutedEventArgs e)
        {
            trayMinimize = false;
            icon.Visible = false;
            File.WriteAllText("trayMinimize", trayMinimize.ToString());
        }
    }
}


