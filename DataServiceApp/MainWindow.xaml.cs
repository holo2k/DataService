using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DataServiceApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        OpenFileDialog openFile = new OpenFileDialog();
        
        public MainWindow()
        {
            InitializeComponent();
            openFile.FileName = "../../../DataService/bin/Debug/DataService.exe";
        }

        private void btnInstall_Click(object sender, RoutedEventArgs e)
        {
            ServiceController service = new ServiceController("Служба интеграции");
            try
            {
                ManagedInstallerClass.InstallHelper(new[] { openFile.FileName });
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            MessageBox.Show("Установка сервиса выполнена!");
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ManagedInstallerClass.InstallHelper(new[] {@"/u", openFile.FileName });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            MessageBox.Show("Сервис удален успешно!");
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            StartService("Служба интеграции");
        }

        private void btnPause_Click(object sender, RoutedEventArgs e)
        {
            StopService("Служба интеграции");
        }

        private void btnReStart_Click(object sender, RoutedEventArgs e)
        {
            RestartService("Служба интеграции");
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
        public static void RestartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            TimeSpan timeout = TimeSpan.FromMinutes(1);
            if (service.Status != ServiceControllerStatus.Stopped)
            {
                Console.WriteLine("Перезапуск службы. Останавливаем службу...");
                // Останавливаем службу
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
            }
            if (service.Status != ServiceControllerStatus.Running)
            {
                Console.WriteLine("Перезапуск службы. Запускаем службу...");
                // Запускаем службу
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                MessageBox.Show("Служба была успешно перезапущена!");
            }
        }

        
    }
}
  

