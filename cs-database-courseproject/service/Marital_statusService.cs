using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;

namespace cs_database_courseproject.service
{
    internal class Marital_statusService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public Marital_statusService() { }
        public void ShowStatus(bool sort,DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                DataTable dt = new DataTable();
                string s = "SELECT ID_Ms, Name AS Наименование FROM Marital_status "; 
                if (sort) { s += "ORDER BY Marital_status.Name"; }
                adapter = new SqlDataAdapter(s, connection);
                adapter.Fill(dt);
                dataGrid.DataSource = dt;
                connection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void truncateTable(bool sort, DataGridView dataGrid)
        {
            try
            {
                cmd = new SqlCommand("TRUNCATE TABLE Marital_status", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowStatus(sort,dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void DeleteStatus(string IdStatus, bool sort, DataGridView dataGrid)
        {
            try
            {
                if (IdStatus != "")
                {
                    cmd = new SqlCommand("DELETE FROM Marital_status WHERE ID_Ms = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(IdStatus));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Семейное положение удалено");
                    Console.WriteLine("Successful");
                    ShowStatus(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбрано семейное положение");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void ChangeMs(string name, string id, bool sort, DataGridView dataGrid)
        {
            try
            {
                if (name != ""  && id != "")
                {
                    cmd = new SqlCommand("UPDATE Marital_status SET Name = @name WHERE @id = ID_Ms",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@name", name);

                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Наименование обновлено");
                    Console.WriteLine("Successful");
                    ShowStatus(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddMs(string name, bool sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "" )
                {
                    cmd = new SqlCommand("INSERT INTO Marital_status (Name)" +
                        " VALUES (@name)", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@name", name);
        
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Семейное положение добавлено");
                    Console.WriteLine("Successful");
                    ShowStatus(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
    }
}
