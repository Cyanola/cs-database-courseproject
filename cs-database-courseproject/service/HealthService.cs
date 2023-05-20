using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace cs_database_courseproject.service
{
    internal class HealthService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public HealthService() { }
        public void ShowHealth(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                string s = "SELECT Health.ID_numb, Workers.Tabel_numb AS [Табельный номер], Workers.Surname AS [Фамилия], Workers.Name AS [Имя], Workers.Patronymic AS [Отчество], Post.Name AS [Должность], Marital_status.Name AS [Семейное\n положение]," +
                    "Health.[Document number] AS [Номер документа], Health.Ill AS [Болезнь], Health.[Organization] AS [Организация]," +
                    "Health.Doctor AS [Врач], Health.[Sick leave date] AS [Дата выхода\n на больничный], Health.[Date of release from sick leave] AS [Дата выхода\n на работу] FROM Workers " +
                    "JOIN Health ON Health.ID_wrk = Workers.ID_wrk " +
                    "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post" +
                    " LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms ";

                if (sort.Text != "")
                {
                    switch (sort.Text)
                    {
                        case "Фамилия":
                            {
                                s += $"ORDER BY Workers.Surname";
                                break;
                            }
                        case "Имя":
                            { s += $"ORDER BY Workers.Name"; break; }
                        case "Отчество":
                            {
                                s += $"ORDER BY Workers.Patronymic"; break;
                            }
                        case "Семейное положение":
                            {
                                s += $"ORDER BY Marital_status.Name";
                                break;
                            }
                        case "Номер документа":
                            {
                                s += $"ORDER BY Health.[Document number]";
                                break;
                            }
                        case "Дата выхода на больничный":
                            {
                                s += $"ORDER BY Health.[Sick leave date]";
                                break;
                            }
                        case "Дата выхода на работу":
                            {
                                s += $"ORDER BY Health.[Date of release from sick leave]";
                                break;
                            }
                        case "Табельный номер сотрудника":
                            {
                                s += $"ORDER BY Workers.Tabel_numb";
                                break;
                            }

                        case "Должность":
                            {
                                s += $"ORDER BY Post.Name";
                                break;
                            }
                        case "Врач":
                            {
                                s += $"ORDER BY Health.[Doctor]";
                                break;
                            }
                        case "Болезнь":
                            {
                                s += $"ORDER BY Health.[Ill]";
                                break;
                            }
                        case "Организация":
                            {
                                s += $"ORDER BY Health.[Organization]";
                                break;
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
        public void TruncateTable(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                cmd = new SqlCommand("TRUNCATE TABLE Health", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowHealth(sort,dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void DeleteHealth(string Idh, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (Idh != "")
                {
                    cmd = new SqlCommand("DELETE FROM Health WHERE ID_numb = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(Idh));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись о здоровье удалена");
                    Console.WriteLine("Successful");
                    ShowHealth(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбрана запись о здоровье");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }


        public void ChangeHealth(string number, string org, string doctor, string tabel, string sickleavedate, string dateofrealease,
            string ill,
           string id, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (id != "" && number != "" && org != "" && doctor != "" && tabel != "" &&
                    sickleavedate != "" && dateofrealease != "" && ill != "")
                {
                    cmd = new SqlCommand("UPDATE Health SET [Document number] = @number, Organization = @org," +
                        "Doctor = @doctor, [Sick leave date] = @sickleavedate, [Date of release from sick leave] = @dateofrealease," +
                        "ID_wrk = (SELECT ID_wrk FROM Workers WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Post = (SELECT Post.ID_Post FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Ms = (SELECT Marital_status.ID_Ms FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = @tabel), " +
                        "Ill = @ill WHERE @id = ID_numb",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@number", number);
                    cmd.Parameters.AddWithValue("@org", org);
                    cmd.Parameters.AddWithValue("@doctor", doctor);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@sickleavedate", sickleavedate);
                    cmd.Parameters.AddWithValue("@dateofrealease", dateofrealease);
                    cmd.Parameters.AddWithValue("@ill", ill);

                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись обновлена");
                    Console.WriteLine("Successful");
                    ShowHealth(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddHealth(string number, string org, string doctor, string tabel, string sickleavedate, string dateofrealease,
            string ill,string wrk, string post11, string ms, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (number != "" && org != "" && doctor != "" && tabel != "" &&
                    sickleavedate != "" &&dateofrealease !="" &&ill!="" && wrk != "" && post11 != "" && ms != "")
                {
                    connection.Open();
                    cmd = new SqlCommand("INSERT INTO Health ( [Document number], Organization, Doctor, [Sick leave date], " +
                        "[Date of release from sick leave],ID_wrk, ID_Post, ID_Ms, Ill)" +
                        " VALUES (@number, @org, @doctor, @sickleavedate, @dateofrealease, @wrk, @post11, @ms,@ill)", connection);
                     
                    cmd.Parameters.AddWithValue("@number", number);
                    cmd.Parameters.AddWithValue("@org", org);
                    cmd.Parameters.AddWithValue("@doctor", doctor);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@sickleavedate", sickleavedate);
                    cmd.Parameters.AddWithValue("@dateofrealease", dateofrealease);
                    cmd.Parameters.AddWithValue("@ill", ill);

                    cmd.Parameters.AddWithValue("@wrk", wrk);
                    cmd.Parameters.AddWithValue("@post11", post11);
                    cmd.Parameters.AddWithValue("@ms", ms);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись добавлена");
                    Console.WriteLine("Successful");
                    ShowHealth(sort, dataGrid);

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
