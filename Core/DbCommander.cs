using MySql.Data.MySqlClient;

namespace Core
{
    public class DbCommander
    {
        private readonly MySqlConnection _con;

        public DbCommander(MySqlConnection con)
        {
            _con = con;
        }

        public long InsertDiscipline(string agencyId, string name, string categoryName, int sheet, double? value)
        {
            var cmd = _con.CreateCommand();
            cmd.CommandText = @"insert into discipline(agency_id,name,category_name,sheet,value) values";
            cmd.CommandText += string.Format(@"('{0}','{1}','{2}',{3},{4})", agencyId, name, categoryName, sheet, value);
            cmd.ExecuteNonQuery();
            return cmd.LastInsertedId;
        }

        public long InsertSubDiscipline(long disciplineId, string name, double? percent)
        {
            var cmd = _con.CreateCommand();
            cmd.CommandText = @"insert into sub_discipline(discipline_id,name,percent) values";
            cmd.CommandText += string.Format(@"({0},'{1}',{2})", disciplineId, name, percent);
            cmd.ExecuteNonQuery();
            return cmd.LastInsertedId;
        }
    }
}