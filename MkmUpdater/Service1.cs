using System;
using System.IO;
using System.ServiceProcess;
using System.Timers;

namespace MkmUpdater
{
    public partial class Service1 : ServiceBase
    {
        private readonly Timer timer;

        public Service1()
        {
            timer = new Timer();
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile($"Service has started at {DateTime.Now}");
            timer.Elapsed  += OnElapsedTime;
            timer.Interval =  5000;
            timer.Enabled  =  true;
        }

        protected override void OnStop()
        {
            WriteToFile($"Service has stopped at {DateTime.Now}");
        }

        private void OnElapsedTime(object sender, ElapsedEventArgs e)
        {
            WriteToFile($"Service is recall at {DateTime.Now}");
        }

        private void WriteToFile(string message)
        {
            var path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            var filePath = path + "\\ServiceLog_" + DateTime.Now.ToShortDateString();

            if (!File.Exists(filePath))
                using (var writer = File.CreateText(filePath))
                {
                    writer.WriteLine(message);
                }
            else
                using (var writer = File.AppendText(filePath))
                {
                    writer.WriteLine(message);
                }
        }
    }
}