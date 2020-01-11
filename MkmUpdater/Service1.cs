using System;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using Provider;

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

        public void WriteToFile(string message)
        {
            if (!Directory.Exists(Constants.Path))
                Directory.CreateDirectory(Constants.Path);

            var filePath = Constants.Path + "\\ServiceLog_" + DateTime.Now.DayOfWeek + ".txt";

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