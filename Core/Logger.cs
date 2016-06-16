using System;
using System.IO;

namespace Core
{
    public class Logger
    {
        private readonly string _file;
        private readonly object _o = new object();

        public Logger(string file)
        {
            _file = AppDomain.CurrentDomain.BaseDirectory + file;
        }

        public void Write(string msg)
        {
            var sw = new StreamWriter(_file, true);
            sw.WriteLine(DateTime.Now + ": " + msg);
            sw.Flush();
            sw.Close();
        }
    }
}