
namespace DataService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ServiceProcess.ServiceProcessInstaller serviceProcessInstaller1;
            this.serviceInstaller1 = new System.ServiceProcess.ServiceInstaller();
            serviceProcessInstaller1 = new System.ServiceProcess.ServiceProcessInstaller();
            // 
            // serviceProcessInstaller1
            // 
            serviceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.NetworkService;
            serviceProcessInstaller1.Password = "";
            serviceProcessInstaller1.Username = "DESKTOP-KORPNPP\\DataService";
            // 
            // serviceInstaller1
            // 
            this.serviceInstaller1.DelayedAutoStart = true;
            this.serviceInstaller1.DisplayName = "Служба интеграции";
            this.serviceInstaller1.ServiceName = "DataService";
            this.serviceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceInstaller1,
            serviceProcessInstaller1});

        }

        #endregion
        public System.ServiceProcess.ServiceInstaller serviceInstaller1;
    }
}