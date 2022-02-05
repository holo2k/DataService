using System.ServiceProcess;

namespace DataService
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        public static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new DataService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
