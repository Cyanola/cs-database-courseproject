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
using System.Data.SqlTypes;

namespace cs_database_courseproject.service
{
    internal class IncomeService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public IncomeService() { }
        public void ShowIncome(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                string s = "SELECT Income.ID_inc, Workers.Tabel_numb AS [Табельный номер], Workers.Surname AS [Фамилия], Workers.Name AS [Имя]," +
                    " Workers.Patronymic AS [Отчество], Post.Name AS [Должность],  Marital_status.Name AS [Семейное\n положение], " +
                    "Income.[Date of enrollment] AS [Дата зачисления], Income.[Total, rub] AS [Всего, руб], Income.[Personal income tax] AS НДФЛ, " +
                    "Income.[Total to be paid] AS [К выплате], Income.Month AS Период FROM Income " +
                  "" +
"INNER JOIN Workers ON Income.ID_wrk = Workers.ID_wrk " +
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
                        case "Дата зачисления":
                            {
                                s += $"ORDER BY Income.[Date of enrollment]";
                                break;
                            }
                        case "Всего, руб":
                            {
                                s += $"ORDER BY Income.[Total, rub]";
                                break;
                            }
                        case "К выплате":
                            {
                                s += $"ORDER BY Income.[Total to be paid]";
                                break;
                            }
                        case "Период":
                            {
                                s += $"ORDER BY Income.Month";
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
                        case "Семейное положение":
                            {
                                s += $"ORDER BY Marital_status.[Name]";
                                break;
                            }
            
                    }
                }
                adapter = new SqlDataAdapter(s,connection);
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
                cmd = new SqlCommand("TRUNCATE TABLE Income", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowIncome(sort, dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void DeleteIncome(string Idinc, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (Idinc != "")
                {
                    cmd = new SqlCommand("DELETE FROM Income WHERE ID_inc = @id", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", int.Parse(Idinc));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись о доходе удалена");
                    Console.WriteLine("Successful");
                    ShowIncome(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбрана запись о доходе");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void ChangeIncome(string date, string total, string ndfl, string totaltobepaid, string month,
   string id, string tabel, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                SqlMoney moneyValue3 = new SqlMoney(Math.Round(Double.Parse(total), 2));
                SqlMoney moneyValue = new SqlMoney(Math.Round(Double.Parse(ndfl), 2));
                SqlMoney moneyValue2 = new SqlMoney(Math.Round(Double.Parse(totaltobepaid), 2));
                if (id != "" && date != "" && total != "" && ndfl != "" && totaltobepaid != "" &&
                    month != "" && tabel != "")
                {
                    cmd = new SqlCommand("UPDATE Income SET [Date of enrollment] = @date, [Total, rub] = @moneyValue3," +
                        "[Personal income tax] = @moneyValue, [Total to be paid] = @moneyValue2, [Month] = @month, " +
                        "ID_wrk = (SELECT Workers.ID_wrk FROM Workers WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Post = (SELECT Post.ID_Post FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Ms = (SELECT Marital_status.ID_Ms FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = @tabel) " +
                        " WHERE @id = ID_inc",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@moneyValue3", moneyValue3);
                    cmd.Parameters.AddWithValue("@moneyValue", moneyValue);
                    cmd.Parameters.AddWithValue("@moneyValue2", moneyValue2);
                    cmd.Parameters.AddWithValue("@month", month);

                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись обновлена");
                    Console.WriteLine("Successful");
                    ShowIncome(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddIncome(string date, string total, string ndfl, string totaltobepaid, string month,
    string tabel,string wrk, string ms, string post11, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                SqlMoney moneyValue3 = new SqlMoney(Math.Round(Double.Parse(total), 2));
                SqlMoney moneyValue = new SqlMoney(Math.Round(Double.Parse(ndfl), 2));
                SqlMoney moneyValue2 = new SqlMoney(Math.Round(Double.Parse(totaltobepaid), 2));
                if (date != "" && total != "" && ndfl != "" && totaltobepaid != "" &&
                    month != "" && tabel != "" &&wrk!=""&&post11!="" &&ms!="")
                {
                    cmd = new SqlCommand("INSERT INTO Income ([Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month,ID_wrk, ID_Post, ID_Ms)" +
                        $" VALUES (@date, @moneyValue3,@moneyValue, @moneyValue2, @month, @wrk,@post11, @ms)", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@moneyValue3", moneyValue3);
                    cmd.Parameters.AddWithValue("@moneyValue", moneyValue);
                    cmd.Parameters.AddWithValue("@moneyValue2", moneyValue2);
                    cmd.Parameters.AddWithValue("@month", month);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@wrk", wrk);
                    cmd.Parameters.AddWithValue("@post11", post11);
                    cmd.Parameters.AddWithValue("@ms",ms);

                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись добавлена");
                    Console.WriteLine("Successful");
                    ShowIncome(sort, dataGrid);

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
