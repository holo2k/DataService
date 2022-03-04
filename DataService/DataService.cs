using ClassLibrary;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using System.IO.MemoryMappedFiles;
using System.Data.OleDb;
using System;
using System.Windows.Forms;
using System.Data.Linq;
using System.Text;
using System.Linq;
using System.Diagnostics;

namespace DataService
{
    public partial class DataService : ServiceBase
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
        public static Dictionary<int, DataTable> dataTables = new Dictionary<int, DataTable>();
        static EventLog myLog;
        public string tableName = "";
        private int currentIndex = 0;
        public string path { get; set; }
        public DataService()
        {
            InitializeComponent();
            this.CanStop = true;
            this.CanPauseAndContinue = true;
            this.AutoLog = true;
           
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
           
            path = args[0];
            dataUpdate();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            DataTable value = null;
            foreach (var item in dataTables.Keys)
            {
                dataTables.TryGetValue(item, out value);
                File.AppendAllText(@"C:\Users\nneke\Desktop\NewFile.txt", value.Rows[0][0].ToString());
            }
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //myLog = new EventLog("D:\\templog.txt");
            //myLog.WriteEntry(ex.Message.ToString());

        }

        protected override void OnStop()
        {
            // Update the service state to Stop Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update the service state to Stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

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
                        tableName = $"ANLGI{hex}_0_{currentTableNumber}";

                        var commandSelectTemperatures = new OleDbCommand($@"SELECT TOP 10 InstValue, DateTime FROM [{tableName}]", dataConnection);
                        commandSelectTemperatures.Prepare();
                        commandSelectTemperatures.CommandTimeout = 1000;
                        commandSelectTemperatures.CommandType = CommandType.Text;
                        dataTemperatureAdapter.SelectCommand = commandSelectTemperatures;

                        using (OleDbDataReader reader = commandSelectTemperatures.ExecuteReader())
                        {
                            dataTemperatureSet.Load(reader, LoadOption.Upsert, dataConnection.DataSource);
                            reader.Close();
                        }
                        dataTemperatureAdapter.Fill(dataTemperatureSet);
                        currentIndex = indexes[i];
                        dataTables.Add(currentIndex, dataTemperatureSet.Tables[0].DefaultView.Table);
                        dataTemperatureSet.Reset();
                        dataConnection.Close();
                    }
                    
                }

            }
            catch (Exception ex)
            {
                
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
                return null;
            }
        }
    }
}
