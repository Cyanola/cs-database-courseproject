using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Data.SqlTypes;

namespace cs_database_courseproject.service
{
    internal class AccuralsService
    {
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public AccuralsService() { }
        public void ShowAccurals(System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
  
                string s = "SELECT Accurals.ID_accur, Workers.Tabel_numb AS [Табельный номер], Workers.Surname AS [Фамилия], Workers.Name AS [Имя], Workers.Patronymic AS [Отчество], Post.Name AS [Должность], " +
                    "Type_of_accural.Accurals AS [Вид начисления], Accurals.Date_ AS [Дата начислений], Accurals.Commentary AS [Комментарий], " +
                "Accurals.Amount AS [Сумма начислений], Income.[Date of enrollment] AS [Дата зачисления], " +
                "Income.[Total, rub] AS [Всего, руб], Income.[Personal income tax] AS НДФЛ, " +
                "Income.[Total to be paid] AS [К выплате], Income.Month AS Период FROM Workers " +
            "INNER JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk "+
"INNER JOIN Income ON Income.ID_inc = Accurals.ID_inc " +
"LEFT JOIN Type_of_accural ON Type_of_accural.ID_tpaccr = Accurals.ID_tpaccr " +
"LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post "+
"LEFT JOIN Marital_status ON Accurals.ID_Ms = Marital_status.ID_Ms ";
                if(sort.Text !="")
                    {
                    switch(sort.Text)
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
                        case "Дата":
                            {
                                s += $"ORDER BY Accurals.Date_";
                              break; }
                        case "Комментарий":
                            {
                                s += $"ORDER BY Accurals.Commentary";
                                break; }
                        case "Сумма начислений": {
                                s += $"ORDER BY Accurals.Amount";
                               break; }
                        case"Вид начисления":
                            {
                                s += $"ORDER BY Type_of_accural.Accurals";
                             break; }
                        case "Табельный номер сотрудника":
                            {
                                s += $"ORDER BY Workers.Tabel_numb";
                               break; }
          
                        case "Должность":
                            {
                                s += $"ORDER BY Post.Name";
                                break;
                            }
                        case "Дата зачисления":
                            {
                                s += $"ORDER BY Income.[Date of enrollment]";
                                break;  }
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
                cmd = new SqlCommand("TRUNCATE TABLE Accurals", connection);
                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Таблица очищена");
                Console.WriteLine("Successful");
                ShowAccurals(sort,dataGrid);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void DeleteAccurals(string Idacc, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                if (Idacc != "")
                {
                    cmd = new SqlCommand("DELETE FROM Accurals WHERE ID_accur = @Idacc", connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@Idacc", int.Parse(Idacc));
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Запись о начислениях удалена");
                    Console.WriteLine("Successful");
                    ShowAccurals(sort,dataGrid);
                }
                else
                {
                    MessageBox.Show("Не выбрана запись о начислениях");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }


        }
        public void ChangeAccurals(string date, string comment, string amount, string tabel, string tpaccr,
            string id, System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        {
            try
            {
                SqlMoney moneyValue = new SqlMoney(Math.Round(Double.Parse(amount), 2));
                if (id!="" && date != "" && comment != "" && amount != "" && tabel != "" &&
                     tpaccr != "")
                {
                    cmd = new SqlCommand("UPDATE Accurals SET Date_ = @date, Commentary = @comment," +
                        "Amount = @moneyValue, "+
                        "ID_tpaccr = (SELECT ID_tpaccr FROM Type_of_accural WHERE Type_of_accural.Accurals = @tpaccr ), "+
                        "ID_inc =  (SELECT Income.ID_inc FROM Income " +
                        "JOIN Accurals ON Accurals.ID_inc = Income.ID_inc " +
"JOIN Workers ON Workers.ID_wrk = Accurals.ID_wrk  WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_wrk = (SELECT ID_wrk FROM Workers WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Post = (SELECT Post.ID_Post FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = @tabel), " +
                        "ID_Ms = (SELECT Marital_status.ID_Ms FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = @tabel) WHERE @id = ID_accur",
                   connection);
                    connection.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@comment", comment);
                    cmd.Parameters.AddWithValue("@moneyValue", moneyValue);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
                    cmd.Parameters.AddWithValue("@tpaccr", tpaccr);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Начисления обновлены");
                    Console.WriteLine("Successful");
                    ShowAccurals(sort, dataGrid);
                }
                else
                {
                    MessageBox.Show("Введите данные");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "State"); }
        }
        public void AddAccurals(string date, string comment, string amount, string tabel, string tpaccr,
            System.Windows.Forms.ComboBox sort, DataGridView dataGrid)
        { 
            try
            {
                SqlMoney moneyValue = new SqlMoney(Math.Round(Double.Parse(amount), 2));
                if (date != "" && comment != "" && amount!="" && tabel != "" &&
                    tpaccr != "")
                {
                    connection.Open();
                    cmd = new SqlCommand($"INSERT INTO Accurals (Date_, Commentary, Amount,ID_tpaccr, ID_inc, ID_wrk, ID_Post, ID_Ms) " +
                        $"SELECT @date, @comment, @moneyValue, Type_of_accural.ID_tpaccr, Income.ID_inc, Workers.ID_wrk, Post.ID_Post, Marital_status.ID_Ms   "+
                      "FROM Workers JOIN Income ON Workers.ID_wrk = Income.ID_wrk JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                      "JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                      "JOIN Type_of_accural ON Type_of_accural.Accurals = @tpaccr " +
                      "WHERE Workers.Tabel_numb = @tabel ", connection); 

                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@comment", comment);
                    cmd.Parameters.AddWithValue("@moneyValue", moneyValue);
                    cmd.Parameters.AddWithValue("@tpaccr", tpaccr);
                    cmd.Parameters.AddWithValue("@tabel", tabel);
               
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Начисления добавлены");
                    Console.WriteLine("Successful");
                    ShowAccurals(sort,dataGrid);

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
