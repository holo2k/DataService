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
                if(service.Status == ServiceControllerStatus.Running)
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
            if(MessageBox.Show("Вы точно хотите удалить службу?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
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

        
    }
}
  

