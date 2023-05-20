using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

namespace cs_database_courseproject.service
{
    internal class Type_of_accuralsService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public Type_of_accuralsService() { }

        public void ShowTypes(bool sort,DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                DataTable dt = new DataTable();
                string s = "SELECT ID_tpaccr, Accurals AS Начисление FROM Type_of_accural ";
                if (sort) { s += "ORDER BY Type_of_accural.Accurals";}
                    adapter = new SqlDataAdapter(s, connection);
                adapter.Fill(dt);
                dataGrid.DataSource = dt;
                connection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void TruncateTable(bool sort, DataGridView dataGrid)
        {
            try
            {
                cmd = new SqlCommand("TRUNCATE TABLE Type_of_accural", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowTypes(sort,dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void DeleteTypes(string Idtp, bool sort,DataGridView dataGrid)
        {
            try
            {
                if (Idtp != "")
                {
                    cmd = new SqlCommand("DELETE FROM Type_of_accural WHERE ID_tpaccr = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(Idtp));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Вид начисления удален");
                    Console.WriteLine("Successful");
                    ShowTypes(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбран вид начисления");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddTypes(string name, bool sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "")
                {
                    cmd = new SqlCommand("INSERT INTO Type_of_accural (Accurals)" +
                        " VALUES (@name)", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@name", name);

                    connection.Close();
                    MessageBox.Show("Вид начисления добавлен");
                    Console.WriteLine("Successful");
                    ShowTypes(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void ChangeTypes(string name, string id, bool sort,DataGridView dataGrid)
        {
            try
            {
                if (name != "" && id != "")
                {
                    cmd = new SqlCommand("UPDATE Type_of_accural SET Accurals = @name WHERE @id = ID_tpaccr",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@name", name);

                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Вид начисления обновлен");
                    Console.WriteLine("Successful");
                    ShowTypes(sort,dataGrid);
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
