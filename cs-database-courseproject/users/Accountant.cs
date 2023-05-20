using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cs_database_courseproject
{
    internal class Accountant
    {
      
       
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public string getPassword()
        {
            connection.Open();
            string s = "SELECT Users.Password FROM Users WHERE ID_User = 2";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();

            }
            reader.Close();
            connection.Close();

            return s;

        }


        public Accountant()
        { }
        

    }
}
