using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Linq;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace DataServiceApp
{
    class DataUpdate
    {
        private readonly Timer _timer;
        string path = "";
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
        static EventLog myLog;
        public string tableName = "";
        private int currentIndex = 0;

        public DataUpdate(string _path)
        {
            _timer = new Timer(10000) { AutoReset = true };
            _timer.Elapsed += timer_Elapsed;
            if (_path.Length > 0)
            {
                path = _path;
            }
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
           //dataUpdate();
            File.AppendAllText(@"C:\Users\nneke\Desktop\fuckyou.txt", "fuck you" + "\n");
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
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
                        // Presuming the DataTable has a column named Date.
                        string expression;
                        expression = $"Index = {currentIndex}";
                        DataRow[] foundRows;

                        // Use the Select method to find all rows matching the filter.
                        foundRows = dataSet.Tables[0].Select(expression);

                        var currentEquipment = foundRows[0][0].ToString();
                        //До сюда норм доходит

                        SqlConnection conn = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=EquipmentTemperatures;" + "Integrated Security=true;");
                        conn.Open();
                        File.AppendAllText(@"C:\Users\nneke\Desktop\conninit.txt", conn.State + "\n");
                        SqlCommand cmdSelectTableNames = new SqlCommand("SELECT TABLE_NAME FROM EquipmentTemperatures.INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME", conn);
                        SqlDataAdapter sqlAdapter = new SqlDataAdapter();
                        DataSet tableNames = new DataSet();
                        sqlAdapter.SelectCommand = cmdSelectTableNames;
                        using (SqlDataReader dataReader = cmdSelectTableNames.ExecuteReader())
                        {
                            tableNames.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                            dataReader.Close();
                        }
                        sqlAdapter.Fill(tableNames);
                        tableNames.Tables[0].Columns[0].Unique = true;
                        tableNames.Tables[0].PrimaryKey = new DataColumn[] { tableNames.Tables[0].Columns["TABLE_NAME"] };
                        bool hasName = false;
                        for (int j = 0; j < tableNames.Tables[0].Rows.Count; j++)
                        {
                            if (tableNames.Tables[0].Rows[j].ItemArray[0].ToString() == currentEquipment) { hasName = true; }
                            File.AppendAllText(@"C:\Users\nneke\Desktop\tableNames.txt", tableNames.Tables[0].Rows[j].ItemArray[0].ToString() + "\n");

                        }
                        if (hasName)
                        {
                            try
                            {
                                SqlCommand cmdSelectTemp = new SqlCommand($"SELECT * FROM [{currentEquipment}]", conn);
                                SqlDataAdapter sqlAdapterTemp = new SqlDataAdapter();
                                DataSet tableTemp = new DataSet();
                                sqlAdapter.SelectCommand = cmdSelectTableNames;
                                using (SqlDataReader dataReader = cmdSelectTemp.ExecuteReader())
                                {
                                    tableTemp.Load(dataReader, LoadOption.Upsert, connection.DataSource);
                                    dataReader.Close();
                                }
                                foreach (DataRow row in dataTemperatureSet.Tables[0].Rows)
                                {
                                    bool hasDate = false;
                                    decimal instValue = Convert.ToDecimal(row[0].ToString());
                                    string value = instValue.ToString().Replace(',', '.');
                                    string dateTime = Convert.ToString(row[1]);
                                    for (int j = 0; j < tableTemp.Tables[0].Rows.Count; j++)
                                    {
                                        if (tableTemp.Tables[0].Rows[j].ItemArray[1].ToString() == dateTime) { hasDate = true; }
                                    }
                                    if (!hasDate)
                                    {
                                        using (var command = new SqlCommand($"INSERT INTO [{currentEquipment}] (InstValue, DateTime) VALUES ({value}, CAST('{dateTime}' AS DateTime))", conn))
                                        {
                                            command.ExecuteNonQuery();
                                            File.AppendAllText(@"C:\Users\nneke\Desktop\values.txt", value + " " + dateTime + "\n");
                                        }
                                    }
                                }

                            }
                            catch (Exception ex)
                            {

                            }

                        }
                        else
                        {
                            try
                            {
                                using (var command = new SqlCommand($"CREATE TABLE [{currentEquipment}] (InstValue float, DateTime dateTime)", conn))
                                {
                                    command.ExecuteNonQuery();
                                }

                                foreach (DataRow row in dataTemperatureSet.Tables[0].Rows)
                                {
                                    decimal instValue = Convert.ToDecimal(row[0].ToString());
                                    string value = instValue.ToString().Replace(',', '.');
                                    string dateTime = Convert.ToString(row[1]);
                                    using (var command = new SqlCommand($"INSERT INTO [{currentEquipment}] (InstValue, DateTime) VALUES ({value}, CAST('{dateTime}' AS DateTime))", conn))
                                    {
                                        command.ExecuteNonQuery();
                                    }
                                }

                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        dataTemperatureSet.Reset();
                    }
                }

            }
            catch (Exception ex)
            {

            }
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
                return null;
            }
        }
    }
}
