using ClassLibrary;
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

namespace DataServiceApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        OpenFileDialog openFile = new OpenFileDialog();
        OleDbConnection connection = new OleDbConnection();
        OleDbConnection dataConnection = new OleDbConnection();
        OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
        OleDbCommand command = new OleDbCommand();
        public static DataSet dataSet = new DataSet();
        public static List<int> indexes = new List<int>();
        public static Dictionary<DataSet,int> dataSets = new Dictionary<DataSet, int>();

        string path = "";

        public MainWindow()
        {
            InitializeComponent();

            //Открытие последнего путя до папки с таблицами
            StreamReader sr = new StreamReader("DatabasePath");
            path = sr.ReadLine();
            sr.Close();
            sr.Dispose();

            tbPath.Text = path;

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
                txtCurrentState.Text = "Текущее состояние службы:\n Служба не установлена";
                txtCurrentState.Foreground = Brushes.Red;
            }
            else
            {
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
            }
            service.Dispose();
        }

        //Выбор папки с температурами по индексу оборудования
        public string selectFolder(int index)
        {
            string hex = index.ToString("X8");
            string pathToFolder = $"{path}\\D0000\\DT{hex}";

            return pathToFolder;
        }

        //Обновление данных
        public void dataUpdate()
        {
            connection = new OleDbConnection();
            connection = OpenConnection(connection, path);
            try
            {
                //Получение таблицы с именами и индексами оборудования
                dataSet.Reset();
                var commandSelectUnits = new OleDbCommand("SELECT UnitName,Index FROM DBMAIN", connection);
                dataAdapter.SelectCommand = commandSelectUnits;
                using (OleDbDataReader dataReader = commandSelectUnits.ExecuteReader())
                {
                    dataSet.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                }
                dataAdapter.Fill(dataSet);
                //Декодирование
                for (int j = 0; j < dataSet.Tables[0].Rows.Count; j++)
                {
                    String stringField = dataSet.Tables[0].Rows[j].ItemArray[0].ToString();
                    Encoding enc = Encoding.GetEncoding(1252);
                    Encoding enc2 = Encoding.GetEncoding(1251);
                    string result = enc2.GetString(enc.GetBytes(stringField));
                    dataSet.Tables[0].Rows[j].BeginEdit();
                    dataSet.Tables[0].Rows[j][0] = result;
                    dataSet.Tables[0].Rows[j].EndEdit();
                    dataSet.Tables[0].Rows[j].AcceptChanges();
                    dataSet.AcceptChanges();
                }
                connection.Close();
                string pathFolder = "";
                //Отбор всех индексов
                for (var i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    indexes.Add(Convert.ToInt32(dataSet.Tables[0].Rows[i][1]));
                }
                //Создание таблиц с температурой для каждой ед. оборудования
                for (int i = 0; i < indexes.Count; i++)
                {
                    //Путь до папки с таблицами температур
                    pathFolder = selectFolder(indexes[i]);
                    if (Directory.Exists(pathFolder))
                    {
                        dataConnection = new OleDbConnection();
                        dataConnection = OpenConnection(dataConnection, pathFolder);

                        //Выбор последней созданной таблицы с температурой
                        string hex = indexes[i].ToString("X4");
                        string[] allfiles = Directory.GetFiles(pathFolder, $"ANLGI{hex}_0_???.DB");
                        string currentTable = allfiles.Last();
                        string currentTableNumber = currentTable.Substring(currentTable.Length - 6);
                        currentTableNumber = currentTableNumber.Substring(0, 3);
                        string tableName = $"ANLGI{hex}_0_{currentTableNumber}";

                        var commandSelectTemperatures = new OleDbCommand
                            ($"SELECT TOP 10 InstValue, DateTime FROM {tableName} ORDER BY 'DateTime' DESC", dataConnection);
                        commandSelectTemperatures.Prepare();

                        OleDbDataAdapter dataTemperatureAdapter = new OleDbDataAdapter();
                        DataSet dataTemperatureSet = new DataSet();
                        dataTemperatureAdapter.SelectCommand = commandSelectTemperatures;

                        //ЗДЕСЬ БЕСКОНЕЧНАЯ ЗАГРУЗКА
                        using (OleDbDataReader reader = commandSelectTemperatures.ExecuteReader())
                        {
                            dataTemperatureSet.Load(reader, LoadOption.Upsert, dataConnection.DataSource);
                        }
                        dataTemperatureAdapter.Fill(dataTemperatureSet);
                        dataSets.Add(dataTemperatureSet, indexes[i]);
                        dgTemperatures.ItemsSource = dataTemperatureSet.Tables[0].DefaultView;
                        dataConnection.Close();
                    }
                }
                dgUnits.ItemsSource = dataSet.Tables[0].DefaultView;
                Manager.dataSet = dataSet;
                dataSet.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
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
            try
            {
                path = tbPath.Text;
                if (connection != null)
                {
                    connection.Close();
                }
                File.WriteAllText("DatabasePath", path);
                dataUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Возникла ошибка, возможно, данного пути не существует. \nПодробно:\n" + ex.Message.ToString());
            }

        }

        private void btnInstall_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            try
            {
                ManagedInstallerClass.InstallHelper(new[] { openFile.FileName });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            MessageBox.Show("Установка сервиса выполнена!");
            txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
            txtCurrentState.Foreground = Brushes.Yellow;
            btnDelete.IsEnabled = true;
            btnInstall.IsEnabled = false;
            btnPause.IsEnabled = false;
            btnReStart.IsEnabled = true;
            btnStart.IsEnabled = true;
        }

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
                MessageBox.Show("Сервис удален успешно!");
                txtCurrentState.Text = "Текущее состояние службы:\n Служба не установлена";
                txtCurrentState.Foreground = Brushes.Red;
                btnInstall.IsEnabled = true;
                btnDelete.IsEnabled = false;
                btnPause.IsEnabled = false;
                btnReStart.IsEnabled = false;
                btnStart.IsEnabled = false;
            }
            else
            {
                return;
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            StartService("Служба интеграции");
            txtCurrentState.Text = "Текущее состояние службы:\n Служба запущена";
            txtCurrentState.Foreground = Brushes.Green;
            btnStart.IsEnabled = false;
            btnPause.IsEnabled = true;
        }

        private void btnPause_Click(object sender, RoutedEventArgs e)
        {
            StopService("Служба интеграции");
            txtCurrentState.Text = "Текущее состояние службы:\n Служба приостановлена";
            txtCurrentState.Foreground = Brushes.Yellow;
            btnStart.IsEnabled = true;
            btnPause.IsEnabled = false;
        }

        private void btnReStart_Click(object sender, RoutedEventArgs e)
        {
            RestartService("Служба интеграции");
            btnStart.IsEnabled = false;
            btnPause.IsEnabled = true;
        }

        public static void StartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Проверяем не запущена ли служба
            if (service.Status != ServiceControllerStatus.Running)
            {
                // Запускаем службу
                service.Start();

                // В течении минуты ждём статус от службы
                service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromMinutes(1));
                MessageBox.Show("Служба была успешно запущена!");
            }
            else
            {
                MessageBox.Show("Служба уже запущена!");
            }
        }

        // Останавливаем службу
        public static void StopService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Если служба не остановлена
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMinutes(1));
                MessageBox.Show("Служба была успешно остановлена!");
            }
            else
            {
                MessageBox.Show("Служба уже остановлена!");
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
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                txtCurrentState.Text = "Текущее состояние службы:\n Служба запущена";
                txtCurrentState.Foreground = Brushes.Green;
                MessageBox.Show("Служба была успешно перезапущена!");
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataUpdate();
        }
    }
}


