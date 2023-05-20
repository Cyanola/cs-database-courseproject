using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace cs_database_courseproject.service
{
    internal class WorkerService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public WorkerService() { }

        public void ShowWorkers(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                DataTable dt = new DataTable();
                string s = "SELECT ID_wrk, Surname AS Фамилия, Workers.Name AS Имя, Patronymic AS Отчество, " +
                    "Sex AS Пол, [Children count] AS [Количество детей], Birthdate AS [Дата рождения]," +
                    "Phone_number AS [Номер телефона], Adress AS Адрес, Tabel_numb AS [Табельный номер]," +
                    " email AS [Электронная почта], Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение] FROM Workers " +
                    "JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                    "JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms ";
                if (sort.Text != "")
                {
                    switch (sort.Text)
                    {
                        case "Фамилия":
                            {
                                s += $"ORDER BY Surname";
                                break;
                            }
                        case "Имя":
                            { s += $"ORDER BY Workers.Name"; break; }
                        case "Отчество":
                            {
                                s += $"ORDER BY Patronymic"; break;
                            }
                        case "Оклад":
                            {
                                s += $"ORDER BY Post.Salary"; break;
                            }
                        case "Пол":
                            {
                                s += $"ORDER BY Sex"; break;
                            }
                        case "Количество детей":
                            {
                                s += $"ORDER BY Workers.[Children count]"; break;
                            }
                        case "Должность":
                            {
                                s += $"ORDER BY Post.Name"; break;
                            }
                        case "Семейное положение":
                            {
                                s += $"ORDER BY Marital_status.Name"; break;
                            }
                        case "ID сотрудника":
                            {
                                s += $"ORDER BY Workers.ID_wrk"; break;
                            }
                        case "Табельный номер":
                            {
                                s += "ORDER BY Workers.Tabel_numb"; break;
                            }
                    }
                }
                adapter = new SqlDataAdapter(s, connection);                    
               
                adapter.Fill(dt);
                dataGrid.DataSource = dt;
                connection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void ChangeWorker(string name, string lastName, string patronymic, string sex, string child_count, 
            string dateOfBirth, string phone,
       string adres, string tabel, string email, string post, string marital_st, string id, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "" && lastName != "" && patronymic != "" && child_count != "" &&
                    dateOfBirth != "" && phone != "" && sex !=
               "" && adres != "" && tabel != "" && email != "" && post != "" && marital_st != "" &&id!="")
                {
                    cmd = new SqlCommand("UPDATE Workers SET Name = @name, Surname = @lastName," +
                        "Patronymic = @patronymic, Sex = @sex, [Children count] = @child_count, " +
                        "Birthdate = @dateOfBirth, Phone_number = @phone," +
                        "Adress = @adres, Tabel_numb = @tabel, email = @email, " +
                        "ID_Post = (SELECT ID_Post FROM Post WHERE Post.Name = @post), " +
                        "ID_Ms = (SELECT Marital_status.ID_Ms FROM Marital_status WHERE Marital_status.Name = @marital_st)" +
                        " WHERE @id = ID_wrk",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@lastName", lastName);
                    cmd.Parameters.AddWithValue("@patronymic", patronymic);
                    cmd.Parameters.AddWithValue("@sex", sex);
                    cmd.Parameters.AddWithValue("@child_count", child_count);
                    cmd.Parameters.AddWithValue("@dateOfBirth", dateOfBirth);
                    cmd.Parameters.AddWithValue("@phone", phone);
                    cmd.Parameters.AddWithValue("@adres", adres);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@post", post);
                    cmd.Parameters.AddWithValue("@marital_st", marital_st);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Cотрудник обновлен");
                    Console.WriteLine("Successful");
                    ShowWorkers(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddWorker(string name, string lastName, string patronymic, string sex, string child_count,
            string dateOfBirth, string phone, string adres, string tabel,
            string email,string post, string marital_st, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (name != "" && lastName != "" && patronymic != "" &&child_count!=""&&
                    dateOfBirth != "" && phone != "" && sex !=
               "" && adres != "" && tabel != "" && email != "" &&post!="" &&marital_st!="")
                {
                    cmd = new SqlCommand("INSERT INTO Workers (Name, Surname, Patronymic, Sex, [Children count], " +
                        "Birthdate, Phone_number, Adress, Tabel_numb, email, ID_Post, ID_Ms)" +
                        " VALUES (@name, @lastName, @patronymic, @sex, @child_count, @dateOfBirth, @phone, @adres, @tabel, @email, @post, @marital_st)", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@lastName", lastName);
                    cmd.Parameters.AddWithValue("@patronymic", patronymic);
                    cmd.Parameters.AddWithValue("@sex", sex);
                    cmd.Parameters.AddWithValue("@child_count", child_count);
                    cmd.Parameters.AddWithValue("@dateOfBirth", dateOfBirth);
                    cmd.Parameters.AddWithValue("@phone", phone);
                    cmd.Parameters.AddWithValue("@adres", adres);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@post", post);
                    cmd.Parameters.AddWithValue("@marital_st", marital_st);
                    cmd.ExecuteNonQuery();
 
                    connection.Close();
                    MessageBox.Show("Сотрудник добавлен");
                    Console.WriteLine("Successful");
                    ShowWorkers(sort, dataGrid);

                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void truncateTable(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                cmd = new SqlCommand("TRUNCATE TABLE Workers", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowWorkers(sort, dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }

        public void DeleteWorker(System.Windows.Forms.ComboBox sort, string Idwrk, DataGridView dataGrid)
        {
            try
            {
                if (Idwrk != "")
                {
                    cmd = new SqlCommand("DELETE FROM Workers WHERE ID_wrk = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(Idwrk));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Сотрудник удален");
                    Console.WriteLine("Successful");
                    ShowWorkers(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбран сотрудник");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
     
    }
}
