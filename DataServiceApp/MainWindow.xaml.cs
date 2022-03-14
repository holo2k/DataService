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
using Topshelf;

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

        static string[] path = new string[1];

        public MainWindow()
        {
            InitializeComponent();

            //Открытие последнего путя до папки с таблицами
            try
            {
                StreamReader sr = new StreamReader("DatabasePath");
                path[0] = sr.ReadLine();
                sr.Close();
                sr.Dispose();
                tbPath.Text = path[0];
            }
            catch(Exception ex)
            {

            }

           //var exitCode = HostFactory.Run(x =>
           //{
           //   x.Service<DataUpdate>(s => 
           //   {
           //       s.ConstructUsing(dataupdate => new DataUpdate(path[0]));
           //       s.WhenStarted(dataupdate => dataupdate.Start());
           //       s.WhenStarted(dataupdate => dataupdate.Stop());
           //   });
           //
           //    x.RunAsLocalSystem();
           //    
           //
           //    x.SetServiceName("DataUpdateService");
           //    x.SetDisplayName("Обновление базы данных температур");
           //    x.SetDescription("Служба, осуществляющая интеграцию локальной базы данных температур оборудования в SQL-Server");
           //
           //});
           //int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
           //Environment.ExitCode = exitCodeValue;

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
                service.Dispose();
            }
            
        }

        //Выбор папки с температурами по индексу оборудования
        public string selectFolder(int index)
        {
            string hex = index.ToString("X8");
            string pathToFolder = $"{path[0]}\\D0000\\DT{hex}";

            return pathToFolder;
        }

        //Обновление данных
        public void dataUpdate()
        {
            connection = new OleDbConnection();
            connection = OpenConnection(connection, path[0]);
            try
            {
                //Получение таблицы с именами и индексами оборудования
                dataSet.Reset();
                var commandSelectUnits = new OleDbCommand("SELECT UnitName,Index FROM DBMAIN", connection);
                dataAdapter.SelectCommand = commandSelectUnits;
                using (OleDbDataReader dataReader = commandSelectUnits.ExecuteReader())
                {
                    dataSet.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                    dataReader.Close();
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
                dataAdapter.Dispose();
                connection.Close();
                string pathFolder = "";
                //Отбор всех индексов
               //for (var i = 0; i < dataSet.Tables[0].Rows.Count; i++)
               //{
               //    indexes.Add(Convert.ToInt32(dataSet.Tables[0].Rows[i][1]));
               //}
               ////Создание таблиц с температурой для каждой ед. оборудования
               //for (int i = 0; i < indexes.Count; i++)
               //{
               //    //Путь до папки с таблицами температур
               //    pathFolder = selectFolder(indexes[i]);
               //    if (Directory.Exists(pathFolder))
               //    {
               //        dataConnection = new OleDbConnection();
               //        dataConnection = OpenConnection(dataConnection, pathFolder);
               //
               //        //Выбор последней созданной таблицы с температурой
               //        string hex = indexes[i].ToString("X4");
               //        string[] allfiles = Directory.GetFiles(pathFolder, $"ANLGI{hex}_0_???.DB");
               //        string currentTable = allfiles.Last();
               //        string currentTableNumber = currentTable.Substring(currentTable.Length - 6);
               //        currentTableNumber = currentTableNumber.Substring(0, 3);
               //        string tableName = $"ANLGI{hex}_0_{currentTableNumber}";
               //
               //        var commandSelectTemperatures = new OleDbCommand($@"SELECT TOP 10 InstValue, DateTime FROM [{tableName}]", dataConnection);
               //        commandSelectTemperatures.Prepare();
               //        commandSelectTemperatures.CommandTimeout = 1000;
               //        commandSelectTemperatures.CommandType = CommandType.Text;
               //        dataTemperatureAdapter.SelectCommand = commandSelectTemperatures;
               //        
               //        using (OleDbDataReader reader = commandSelectTemperatures.ExecuteReader())
               //        {
               //            dataTemperatureSet.Load(reader, LoadOption.Upsert, dataConnection.DataSource);
               //             reader.Close();
               //        }
               //        dataTemperatureAdapter.Fill(dataTemperatureSet);
               //        currentIndex = indexes[i];
               //
               //        // Presuming the DataTable has a column named Date.
               //        string expression;
               //        expression = $"Index = {currentIndex}";
               //        DataRow[] foundRows;
               //
               //        // Use the Select method to find all rows matching the filter.
               //        foundRows = dataSet.Tables[0].Select(expression);
               //
               //        var currentEquipment = foundRows[0][0].ToString();
                       
                        //SqlConnection conn = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=EquipmentTemperatures;" + "Integrated Security=true;");
                        //conn.Open();
                        //SqlCommand cmdSelectTableNames = new SqlCommand("SELECT TABLE_NAME FROM EquipmentTemperatures.INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME", conn);
                        //SqlDataAdapter sqlAdapter = new SqlDataAdapter();
                        //DataSet tableNames = new DataSet();
                        //sqlAdapter.SelectCommand = cmdSelectTableNames;
                        //using (SqlDataReader dataReader = cmdSelectTableNames.ExecuteReader())
                        //{
                        //    tableNames.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                        //    dataReader.Close();
                        //}
                        //sqlAdapter.Fill(tableNames);
                        //tableNames.Tables[0].Columns[0].Unique = true;
                        //tableNames.Tables[0].PrimaryKey = new DataColumn[] { tableNames.Tables[0].Columns["TABLE_NAME"] };
                        //bool hasName = false;
                        //for (int j = 0; j < tableNames.Tables[0].Rows.Count; j++)
                        //{
                        //    if (tableNames.Tables[0].Rows[j].ItemArray[0].ToString() == currentEquipment) { hasName = true; }
                        //
                        //}
                        //if (hasName)
                        //{
                        //    try
                        //    {
                        //         SqlCommand cmdSelectTemp = new SqlCommand($"SELECT * FROM [{currentEquipment}]", conn);
                        //         SqlDataAdapter sqlAdapterTemp = new SqlDataAdapter();
                        //         DataSet tableTemp = new DataSet();
                        //         sqlAdapter.SelectCommand = cmdSelectTableNames;
                        //         using (SqlDataReader dataReader = cmdSelectTemp.ExecuteReader())
                        //         {
                        //             tableTemp.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                        //             dataReader.Close();
                        //         }
                        //         foreach (DataRow row in dataTemperatureSet.Tables[0].Rows)
                        //         {
                        //             bool hasDate = false;
                        //             decimal instValue = Convert.ToDecimal(row[0].ToString());
                        //             string value = instValue.ToString().Replace(',', '.');
                        //             string dateTime = Convert.ToString(row[1]);
                        //             for (int j = 0; j < tableTemp.Tables[0].Rows.Count; j++)
                        //             {
                        //                 if (tableTemp.Tables[0].Rows[j].ItemArray[1].ToString() == dateTime) { hasDate = true; }
                        //             }
                        //             if (!hasDate)
                        //             {
                        //                 using (var command = new SqlCommand($"INSERT INTO [{currentEquipment}] (InstValue, DateTime) VALUES ({value}, CAST('{dateTime}' AS DateTime))", conn))
                        //                 {
                        //                     command.ExecuteNonQuery();
                        //                     File.AppendAllText(@"C:\Users\nneke\Desktop\values.txt", value + " " + dateTime + "\n");
                        //                 }
                        //             }
                        //         }
                        //
                        //     }
                        //    catch (Exception ex)
                        //    {
                        //        System.Windows.Forms.MessageBox.Show(ex.Message.ToString());
                        //    }
                        //
                        //}
                        //else
                        //{
                        //    try
                        //    {
                        //        using (var command = new SqlCommand($"CREATE TABLE [{currentEquipment}] (InstValue float, DateTime dateTime)", conn))
                        //        {
                        //            command.ExecuteNonQuery();
                        //        }
                        //
                        //        foreach (DataRow row in dataTemperatureSet.Tables[0].Rows)
                        //        {
                        //            decimal instValue = Convert.ToDecimal(row[0].ToString());
                        //            string value = instValue.ToString().Replace(',', '.');
                        //            string dateTime = Convert.ToString(row[1]);
                        //            using (var command = new SqlCommand($"INSERT INTO [{currentEquipment}] (InstValue, DateTime) VALUES ({value}, CAST('{dateTime}' AS DateTime))", conn))
                        //            {
                        //                command.ExecuteNonQuery();
                        //            }
                        //        }
                        //
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        System.Windows.Forms.MessageBox.Show(ex.Message.ToString());
                        //    }
                        //}
                        dataTemperatureSet.Reset();
                   // }
               // }
                dgUnits.ItemsSource = dataSet.Tables[0].DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            
                GC.Collect();
            dataConnection.Close();

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
                    dataUpdate();
                    GC.Collect();
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
                        dataUpdate();
                        GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Возникла ошибка, возможно, данного пути не существует. \nПодробно:\n" + ex.Message.ToString());
                    }
                }
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
            MessageBox.Show("Служба была успешно запущена!");
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

        public static void StartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            // Проверяем не запущена ли служба
            if (service.Status != ServiceControllerStatus.Running)
            {
                // Запускаем службу
                
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
            ServiceController service = new ServiceController(serviceName);
            // Если служба не остановлена
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMinutes(1));
                
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataUpdate();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopService("Служба интеграции");
        }
    }
}


