using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.Remoting.Lifetime;
using System.Reflection.Emit;
using System.Drawing;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Runtime.Remoting.Activation;
using System.Runtime.Remoting.Messaging;

namespace cs_database_courseproject.service
{
    internal class ReportService
    {
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public ReportService()
        {
        }

        public void Report(string surname, string name, string patr, DataGridView datagrid)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            string d = $"SELECT Workers.ID_wrk FROM Workers ";
            List<string> list1 = new List<string>();
            connection.Open();
            cmd = new SqlCommand(d, connection);
           SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list1.Add(reader[0].ToString());

            }
            connection.Close();
            reader.Close();
            string dequery = $"SELECT Workers.ID_wrk FROM Workers WHERE " +
  $"Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' ";
            string s = "";
            connection.Open();
            cmd = new SqlCommand(dequery, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = reader[0].ToString();

            }
            connection.Close();
            reader.Close();
            foreach (var item1 in list1)
            {
                string deQueryOne = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
                    "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
                    " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
                    "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
                    "\r\nWHERE Health.ID_wrk IS NULL AND Accurals.ID_wrk IS NULL AND Income.ID_wrk IS NULL;";
                connection.Open();
                List < string>    list = new List<string>();
                cmd = new SqlCommand(deQueryOne, connection);
              reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                string query = "";
                foreach (var item in list)
                {
                    if (item == s &&item ==item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                                "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, " +
                                "Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
                                "COUNT(Health.[Sick leave date]) AS [Кол - во пропусков по болезни],AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Workers " +
                                "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                                "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                                "LEFT JOIN(     SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                                " [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] FROM Health " +
                                " GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] )" +
                                " Health ON Workers.ID_wrk = Health.ID_wrk " +
                                "LEFT JOIN( SELECT  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax],[Total to be paid], Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income " +
                                "GROUP BY  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month ) " +
                                "Income ON Workers.ID_wrk = Income.ID_wrk  " +
                                "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex," +
                                " Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, Marital_status.Name " +
                                $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                                "\nORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";
                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }

                string deQueryTwo = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
                  "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
                  " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
                  "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
                  "\r\nWHERE Health.ID_wrk IS NULL AND Accurals.ID_wrk IS NULL AND Income.ID_wrk IS NOT NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQueryTwo, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                            "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, " +
                            "Marital_status.Name AS[Семейное положение], COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни], " +
                            "Income.[Date of enrollment] AS[Дата зачисления], Income.[Total, rub] AS[Всего, руб], Income.[Personal income tax] AS НДФЛ, Income.[Total to be paid] AS[К выплате]," +
                            "Income.Month AS Период, AVG(Income.[Total to be paid]) AS [Средний доход] FROM Workers " +
                            "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                            "LEFT JOIN( " +
                            " SELECT Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], " +
                            "[Total to be paid], Month, AVG(Income.[Total to be paid]) AS [Средний доход]  " +
                            "    FROM Income GROUP BY  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "LEFT JOIN(     SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                            " [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] " +
                            "  FROM Health " +
                            "   GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] ) Health ON Workers.ID_wrk = Health.ID_wrk " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name, Income.[Date of enrollment], Income.[Total, rub], Income.[Personal income tax]," +
                            " Income.[Total to be paid], Income.Month " +
                             $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";
                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
                string deQueryThree = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
             "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
             " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
             "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
             "\r\nWHERE Health.ID_wrk IS NULL AND Accurals.ID_wrk IS NOT NULL AND Income.ID_wrk IS NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQueryThree, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей]," +
                            "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS[Семейное положение]," +
                            "COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                            "Accurals.Date_ AS[Дата начислений], Accurals.Commentary AS[Комментарий], " +
                            "Accurals.Amount AS[Сумма начислений], Type_of_accural.Accurals AS[Вид начисления], AVG(Income.[Total to be paid]) AS [Средний доход] " +
                            "FROM Workers " +
                            "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                            "LEFT JOIN(     SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                            " [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] " +
                            "  FROM Health  GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] )" +
                            " Health ON Workers.ID_wrk = Health.ID_wrk " +
                            "LEFT JOIN(     SELECT Accurals.ID_wrk, Date_, Commentary, Amount, Accurals.ID_tpaccr " +
                            "    FROM Accurals ) Accurals ON Workers.ID_wrk = Accurals.ID_wrk " +
                            "LEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr " +
                            "LEFT JOIN( SELECT  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income GROUP BY  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name, Accurals.Date_, Accurals.Commentary, Accurals.Amount, " +
                            "Type_of_accural.Accurals " +
                            $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";


                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
                string deQueryFour = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
          "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
          " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
          "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
          "\r\nWHERE Health.ID_wrk IS NOT NULL AND Accurals.ID_wrk IS NULL AND Income.ID_wrk IS NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQueryFour, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                            "Workers.Tabel_numb AS [Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение], " +
                            "COUNT(Health.[Sick leave date]) AS [Кол - во пропусков по болезни], Health.[Document number] AS [Номер документа], Health.Ill AS[Болезнь], " +
                            "Health.[Organization] AS [Организация], Health.Doctor AS[Врач], Health.[Sick leave date] AS [Дата выхода\n на больничный]," +
                            " Health.[Date of release from sick leave] AS [Дата выхода\n на работу], AVG(Income.[Total to be paid]) AS [Средний доход]  " +
                            "FROM Workers " +
                            "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                            "LEFT JOIN( " +
                            "    SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни], [Document number], Ill," +
                            " [Organization], Doctor, [Sick leave date], [Date of release from sick leave]                     FROM Health " +
                            "    GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] " +
                            ") Health ON Workers.ID_wrk = Health.ID_wrk " +
                            "LEFT JOIN( SELECT  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax],[Total to be paid], Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income  GROUP BY Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name, Health.[Document number], Health.Ill, Health.[Organization], Health.Doctor, Health.[Sick leave date], Health.[Date of release from sick leave] " +
                                                 $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";


                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
                string deQueryFive = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
          "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
          " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
          "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
          "\r\nWHERE Health.ID_wrk IS NOT NULL AND Accurals.ID_wrk IS NOT NULL AND Income.ID_wrk IS NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQueryFive, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (        item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                            "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS[Семейное положение], " +
                            "COUNT(Health.[Sick leave date]) AS[Кол - во пропусков по болезни], Health.[Document number] AS[Номер документа], Health.Ill AS[Болезнь], Health.[Organization] AS[Организация], " +
                            "Health.Doctor AS[Врач], Health.[Sick leave date] AS[Дата выхода\n на больничный], Health.[Date of release from sick leave] AS[Дата выхода\n на работу], " +
                            " Accurals.Date_ AS[Дата начислений], Accurals.Commentary AS[Комментарий], Accurals.Amount AS[Сумма начислений], " +
                            "Type_of_accural.Accurals AS[Вид начисления], AVG(Income.[Total to be paid]) AS [Средний доход]  " +
                            "FROM Workers LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                            "LEFT JOIN(     SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                            " [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] " +
                            " FROM Health " +
                            "    GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] " +
                            ") Health ON Workers.ID_wrk = Health.ID_wrk LEFT JOIN( " +
                            "    SELECT Accurals.ID_wrk, Date_, Commentary, Amount, Accurals.ID_tpaccr " +
                            "    FROM Accurals " +
                            ") Accurals ON Workers.ID_wrk = Accurals.ID_wrk " +
                            "LEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr " +
                            "LEFT JOIN( SELECT  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax],[Total to be paid], Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income  GROUP BY Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month ) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name, Health.[Document number], Health.Ill, Health.[Organization], Health.Doctor, Health.[Sick leave date], Health.[Date of release from sick leave], " +
                            " Accurals.Date_, Accurals.Commentary, Accurals.Amount, Type_of_accural.Accurals " +
                                                 $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";


                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
                string deQuerySix = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
        "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
        " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
        "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
        "\r\nWHERE Health.ID_wrk IS NOT NULL AND Accurals.ID_wrk IS NOT NULL AND Income.ID_wrk IS NOT NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQuerySix, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество," +
                      " Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей],\r\nWorkers.Tabel_numb AS [Табельный номер]," +
                      " Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
                      "\r\nCOUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни], Health.[Document number] AS [Номер документа], Health.Ill AS [Болезнь]," +
                      " Health.[Organization] AS [Организация],\r\nHealth.Doctor AS [Врач], Health.[Sick leave date] AS [Дата выхода\n на больничный]," +
                      " Health.[Date of release from sick leave] AS [Дата выхода\\n на работу], \r\nIncome.[Date of enrollment] AS [Дата зачисления]," +
                      " Income.[Total, rub] AS [Всего, руб], Income.[Personal income tax] AS НДФЛ, Income.[Total to be paid] AS [К выплате]," +
                      "\r\nIncome.Month AS Период, Accurals.Date_ AS [Дата начислений], Accurals.Commentary AS [Комментарий], Accurals.Amount AS [Сумма начислений]," +
                      " Type_of_accural.Accurals AS [Вид начисления], AVG(Income.[Total to be paid]) AS [Средний доход]  \r\nFROM Workers\r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \nLEFT JOIN Marital_status" +
                      " ON Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN (\r\n    SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни]," +
                      " [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave]\r\n    FROM Health\r\n   " +
                      " GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave]\r\n) " +
                      "Health ON Workers.ID_wrk = Health.ID_wrk " +
                      "LEFT JOIN( SELECT Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid]," +
                                " Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income GROUP BY Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                      "LEFT JOIN (   SELECT Accurals.ID_wrk, Date_," +
                      " Commentary, Amount, Accurals.ID_tpaccr    FROM Accurals )" +
                      " Accurals ON Workers.ID_wrk = Accurals.ID_wrk " +
                      "\nLEFT JOIN Type_of_accural ON " +
                      "Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr " +
                      "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count]," +
                      " Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, Marital_status.Name, Health.[Document number], Health.Ill, Health.[Organization]," +
                      " Health.Doctor, Health.[Sick leave date], Health.[Date of release from sick leave], Income.[Date of enrollment], Income.[Total, rub], " +
                      "Income.[Personal income tax], Income.[Total to be paid], Income.Month, Accurals.Date_, Accurals.Commentary, Accurals.Amount, " +
                      "Type_of_accural.Accurals " +
                      $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                      "\nORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";
                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
                string deQuerySeven = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
        "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
        " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
        "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
        "\r\nWHERE Health.ID_wrk IS NOT NULL AND Accurals.ID_wrk IS NULL AND Income.ID_wrk IS NOT NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQuerySeven, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();


                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                            "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS[Семейное положение], " +
                            "COUNT(Health.[Sick leave date]) AS[Кол - во пропусков по болезни], Health.[Document number] AS[Номер документа], Health.Ill AS[Болезнь], Health.[Organization] AS[Организация], " +
                            "Health.Doctor AS[Врач], Health.[Sick leave date] AS[Дата выхода\n на больничный], Health.[Date of release from sick leave] AS[Дата выхода\n на работу], " +
                            "Income.[Date of enrollment] AS[Дата зачисления], Income.[Total, rub] AS[Всего, руб], Income.[Personal income tax] AS НДФЛ, Income.[Total to be paid] AS[К выплате]," +
                            "Income.Month AS Период, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Workers " +
                            "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                            "LEFT JOIN( " +
                            "    SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни], [Document number], Ill, [Organization], Doctor, [Sick leave date]," +
                            " [Date of release from sick leave]  FROM Health " +
                            "    GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] ) Health ON Workers.ID_wrk = Health.ID_wrk " +
                            "LEFT JOIN( SELECT Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid]," +
                                " Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income GROUP BY  Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name, Health.[Document number], Health.Ill, Health.[Organization], Health.Doctor, Health.[Sick leave date], Health.[Date of release from sick leave], " +
                            "Income.[Date of enrollment], Income.[Total, rub], Income.[Personal income tax], Income.[Total to be paid], Income.Month " +
                                                 $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";


                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;

                    }
                }
                string deQueryEight = " SELECT Workers.ID_wrk \r\nFROM Workers \r\nLEFT JOIN Health ON Workers.ID_wrk = Health.ID_wrk " +
       "\r\nLEFT JOIN Accurals ON Workers.ID_wrk = Accurals.ID_wrk \r\nLEFT JOIN Income ON Workers.ID_wrk = Income.ID_wrk" +
       " \r\nLEFT JOIN Post ON Workers.ID_Post = Post.ID_Post \r\nLEFT JOIN Marital_status ON " +
       "Workers.ID_Ms = Marital_status.ID_Ms\r\nLEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr" +
       "\r\nWHERE Health.ID_wrk IS NULL AND Accurals.ID_wrk IS NOT NULL AND Income.ID_wrk IS NOT NULL;";
                connection.Open();
                list = new List<string>();
                cmd = new SqlCommand(deQueryEight, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader[0].ToString());

                }
                connection.Close();
                reader.Close();

                foreach (var item in list)
                {
                    if (item == s && item == item1)
                    {
                        connection.Open();
                        query = "SELECT Workers.Surname AS Фамилия, Workers.Name AS Имя, Workers.Patronymic AS Отчество, Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей], " +
                            "Workers.Tabel_numb AS[Табельный номер], Post.Name AS Должность , Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS[Семейное положение]," +
                            "COUNT([Sick leave date]) AS [Кол-во пропусков по болезни], Income.[Date of enrollment] AS[Дата зачисления], Income.[Total, rub] AS[Всего, руб], Income.[Personal income tax] AS НДФЛ, Income.[Total to be paid] AS[К выплате], " +
                            "Income.Month AS Период, Accurals.Date_ AS[Дата начислений], Accurals.Commentary AS[Комментарий], Accurals.Amount AS[Сумма начислений], Type_of_accural.Accurals AS[Вид начисления], AVG(Income.[Total to be paid]) AS [Средний доход]  " +
                            "FROM Workers " +
                            "LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                            "LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms LEFT JOIN( SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни], [Document number]," +
                            " Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave]  FROM Health " +
                            " GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] ) Health ON Workers.ID_wrk = Health.ID_wrk "
                            +
                            "LEFT JOIN( SELECT Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid]," +
                                " Month, AVG(Income.[Total to be paid]) AS [Средний доход]  FROM Income GROUP BY Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month) Income ON Workers.ID_wrk = Income.ID_wrk  " +
                            "LEFT JOIN( " +         
                            "    SELECT Accurals.ID_wrk, Date_, Commentary, Amount, Accurals.ID_tpaccr     FROM Accurals ) Accurals ON Workers.ID_wrk = Accurals.ID_wrk " +
                            "LEFT JOIN Type_of_accural ON Accurals.ID_tpaccr = Type_of_accural.ID_tpaccr " +
                            "GROUP BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                            "Marital_status.Name,  Income.[Date of enrollment], Income.[Total, rub], Income.[Personal income tax], Income.[Total to be paid]," +
                            " Income.Month, Accurals.Date_, Accurals.Commentary, Accurals.Amount, Type_of_accural.Accurals " +
                                                 $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' " +
                            "ORDER BY Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Tabel_numb";

                        adapter = new SqlDataAdapter(query, connection);
                        adapter.Fill(dt);
                        datagrid.DataSource = dt;
                        connection.Close();
                        break;
                    }
                }
            }
        }

        
        public void ReportCity(string city,string inc, DataGridView datagrid)
        {
            try
            {
                if (city != ""&&inc !="")
                {
                    string s = "SELECT Workers.Surname AS Фамилия,  Workers.Name AS Имя,  Workers.Patronymic AS Отчество, " +
                        " Workers.Sex AS Пол,  Workers.[Children count] AS[Количество детей], Workers.Adress,  Workers.Tabel_numb AS[Табельный номер], " +
                        "   Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор," +
                        "   Marital_status.Name AS[Семейное положение], Income.[Total to be paid] AS[К выплате], " +
                        "   Income.Month AS Период, SUM(Income.[Total to be paid]) AS[Общий доход]  FROM Workers " +
                        "   LEFT JOIN Post ON Workers.ID_Post = Post.ID_Post " +
                        "   LEFT JOIN Marital_status ON Workers.ID_Ms = Marital_status.ID_Ms " +
                        "  LEFT JOIN(SELECT Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month " +
                        "  FROM Income GROUP BY Income.ID_wrk, [Date of enrollment], [Total, rub], [Personal income tax], [Total to be paid], Month ) Income ON Workers.ID_wrk = Income.ID_wrk " +
                        "    GROUP BY  Workers.Surname, Workers.Name, Workers.Patronymic, Workers.Sex, Workers.[Children count], Workers.Tabel_numb, Post.Name, Post.Salary, Post.Director, " +
                        " Marital_status.Name,  Income.[Total to be paid], Workers.Adress, " +
                        "Income.Month  " +
                        $" HAVING SUM(Income.[Total to be paid]) > {inc} AND Workers.Adress LIKE '{city}%' ";
                    connection.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapter = new SqlDataAdapter(s, connection);

                    adapter.Fill(dt);
                    datagrid.DataSource = dt;
                    connection.Close();
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void ReportHealth(string count, DataGridView datagrid)
        {
            try
            {
                if (count != "")
                {
                    string s = "SELECT Workers.Surname AS Фамилия,  Workers.Name AS Имя,  Workers.Patronymic AS Отчество, " +
                       "Workers.Sex AS Пол,  Workers.[Children count] AS [Количество детей],  Workers.Tabel_numb AS [Табельный номер], " +
                       "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
                       "COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни],Health.[Sick leave date] AS[Дата выхода\n на больничный]," +
                       " Health.[Date of release from sick leave] AS[Дата выхода\n на работу] FROM Workers " +
                       "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                       "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
                       "LEFT JOIN( SELECT Health.ID_wrk, COUNT([Sick leave date]) AS [Кол-во пропусков по болезни], " +
                       "[Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave]" +
                       "  FROM Health  GROUP BY Health.ID_wrk, [Document number], Ill, [Organization], Doctor, [Sick leave date], [Date of release from sick leave] ) Health ON Workers.ID_wrk = Health.ID_wrk " +
                       "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Post.Name, Workers.[Children count], " +
                       "Workers.Tabel_numb,  Post.Salary, Post.Director, Marital_status.Name, Health.[Sick leave date],Health.[Date of release from sick leave] " +
                       $"HAVING COUNT(Health.[Sick leave date]) >={count}";
                    connection.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapter = new SqlDataAdapter(s, connection);

                    adapter.Fill(dt);
                    datagrid.DataSource = dt;
                    connection.Close();
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void ReportYear(string year, DataGridView datagrid)
        {
            try
            {
                if (year != "")
                {
                    string s = "SELECT Workers.Surname AS Фамилия,  Workers.Name AS Имя,  Workers.Patronymic AS Отчество, " +
                   "Workers.Sex AS Пол, Workers.Birthdate, Workers.[Children count] AS [Количество детей],  Workers.Tabel_numb AS [Табельный номер], " +
                   "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение] FROM Workers " +
                   "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                   "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
                   "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Workers.Birthdate, Post.Name, Workers.[Children count], " +
                   "Workers.Tabel_numb,  Post.Salary, Post.Director, Marital_status.Name " +
                   $"HAVING Workers.Birthdate >'{year}-01-01'";
                    connection.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapter = new SqlDataAdapter(s, connection);

                    adapter.Fill(dt);
                    datagrid.DataSource = dt;
                    connection.Close();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void ReportPost(string post, DataGridView datagrid)
        {
            try
            {
                if (post != "")
                {
                    string s = "SELECT Workers.Surname AS Фамилия,  Workers.Name AS Имя,  Workers.Patronymic AS Отчество, " +
                           "Workers.Sex AS Пол, Workers.[Children count] AS [Количество детей],  Workers.Tabel_numb AS [Табельный номер], " +
                           "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение] FROM Workers " +
                           "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                           "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
                           "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Post.Name, Workers.[Children count], " +
                           "Workers.Tabel_numb,  Post.Salary, Post.Director, Marital_status.Name " +
                           $"HAVING Post.Name = '{post}'";
                    connection.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapter = new SqlDataAdapter(s, connection);

                    adapter.Fill(dt);
                    datagrid.DataSource = dt;
                    connection.Close();
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, ""); }
        }


        public void ExportExcel(DataGridView dataGrid)
        {
            try
            {
                Microsoft.Office.Interop.Excel._Application app = new
               Microsoft.Office.Interop.Excel.Application();
                // создаем новый WorkBook
                Microsoft.Office.Interop.Excel._Workbook workbook =
               app.Workbooks.Add(Type.Missing);
                // новый Excelsheet в workbook 
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;
                worksheet = workbook.Sheets["Лист1"];
                worksheet = workbook.ActiveSheet;
                // задаем имя для worksheet
                worksheet.Name = "Exported from gridview";
                for (int i = 1; i < dataGrid.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGrid.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGrid.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGrid.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] =
                       dataGrid.Rows[i].Cells[j].Value.ToString();
                    }
                }

                workbook.SaveAs("output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange);
                // закрываем подключение к excel 
                app.Quit();
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, ""); }
        }
        public void ExportWord(DataGridView DGV)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //добавим поля и колонки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    }
                }
                Microsoft.Office.Interop.Word.Document oDoc = new
               Microsoft.Office.Interop.Word.Document();
                oDoc.Application.Visible = true;
                //страницы
                oDoc.PageSetup.Orientation =
               Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;
                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблиц
                oRange.Text = oTemp;
                object Separator =
               Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior =
               Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent;
                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                Type.Missing, Type.Missing, ref ApplyBorders,
                Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, ref AutoFit, ref AutoFitBehavior,
               Type.Missing);
                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //Стили заголовков
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {

                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text =
    DGV.Columns[c].HeaderText;
                }
                //Текст заголовка
                foreach (Microsoft.Office.Interop.Word.Section section in
               oDoc.Application.ActiveDocument.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange =
                   section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange,
                   Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Отчет";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

      
            }
        
        }
    }
}
