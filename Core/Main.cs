using System;
using MySql.Data.MySqlClient;

namespace Core
{
    public class Main
    {
        private readonly Logger _log;

        public Main()
        {
            _log = new Logger("log.txt");
        }
        public void Process(string file)
        {
            //const string conStr = @"server=27.254.41.2;user id=anuchap;persistsecurityinfo=True;database=adsurvey;password=morningM00n;";
            const string conStr = @"server=localhost;user id=root;persistsecurityinfo=True;database=adsurvey;password=morningM00n;";
            using (var con = new MySqlConnection(conStr))
            {
                con.Open();
                var tr = con.BeginTransaction();
                var reader = new ExcelReader(new DbCommander(con), ExcelConfiguration.GetConfiguration());
                try
                {
                    reader.Read(file);
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    _log.Write(ex.ToString());
                }
            }
        }

        public void Log(string msg)
        {
            _log.Write(msg);
        }
    }
}