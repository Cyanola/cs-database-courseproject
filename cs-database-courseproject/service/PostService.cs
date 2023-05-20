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
    internal class PostService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public PostService() { }
        public void ShowPost(System.Windows.Forms.ComboBox sort,DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                DataTable dt = new DataTable();
            string s = "SELECT ID_Post, Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор," +
                    " (SELECT COUNT(*) FROM Workers WHERE Workers.ID_Post = Post.ID_Post) AS [Количество работников] FROM Post ";
                if (sort.Text != "")
                {
                    switch (sort.Text)
                    {
                        case "Должность": s += $"ORDER BY Post.Name"; break;
                        case "Оклад": s += $"ORDER BY Post.Salary"; break;
                        case "Директор": s += $"ORDER BY Post.Director"; break;
                        case "Количество работников": s += $"ORDER BY [Количество работников] "; break;
                    }
                }
                
                adapter = new SqlDataAdapter(s, connection);
                adapter.Fill(dt);
                dataGrid.DataSource = dt;
                connection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void TruncateTable(System.Windows.Forms.ComboBox sort,DataGridView dataGrid)
        {
            try
            {
                cmd = new SqlCommand("TRUNCATE TABLE Post", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowPost(sort,dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void DeletePost(System.Windows.Forms.ComboBox sort, string Idpst, DataGridView dataGrid)
        {
            try
            {
                if (Idpst != "")
                {
                    cmd = new SqlCommand("DELETE FROM Post WHERE ID_Post = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(Idpst));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Должность удалена");
                    Console.WriteLine("Successful");
                    ShowPost(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбрана должность");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void AddPost(string name, string salary, string director, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "" && salary != "" && director != "")
                {
                    cmd = new SqlCommand("INSERT INTO Post (Name, Salary, Director)" +
                        " VALUES (@name, @salary, @director)", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@salary", salary);
                    cmd.Parameters.AddWithValue("@director", director);
          
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Должность добавлена");
                    Console.WriteLine("Successful");
                    ShowPost(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void ChangePost(string name, string salary, string director, string id, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "" && salary != "" && director != "" && id != "")
                {
                    cmd = new SqlCommand("UPDATE Post SET Name = @name, Salary = @salary," +
                        "Director = @director WHERE @id = ID_Post",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@salary", salary);
                    cmd.Parameters.AddWithValue("@director", director);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Должность обновлена");
                    Console.WriteLine("Successful");
                    ShowPost(sort, dataGrid);
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
