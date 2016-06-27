using System;
using Core;
using System.IO;
using System.ServiceProcess;

namespace ASS
{
    public partial class Service : ServiceBase
    {
        private FileSystemWatcher _watcher;
        private Main _main;

        public Service()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            _watcher = new FileSystemWatcher
            {
                Path = @"C:\inetpub\wwwroot\aspnet_client\ad2016\completed",
                NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName,
                Filter = "*.xlsx",
                EnableRaisingEvents = true,
            };
            _watcher.Created += OnCreated;
            _main = new Main();
            _main.Log("Service started");
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            _main.Log("File: " + e.FullPath + " Processing...");
            try
            {
                _main.Process(e.FullPath);
                _main.Log("File: " + e.FullPath + " Completed :)");
            }
            catch (Exception ex)
            {
                _main.Log("File: " + e.FullPath + " Error: " + ex);
            }
        }

        protected override void OnStop()
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Dispose();
            _main.Log("Service stoped");
        }
    }
}