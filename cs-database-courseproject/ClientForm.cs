using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace cs_database_courseproject
{
    public partial class ClientForm : Form
    {
        private readonly service.ReportService report = new service.ReportService();
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public ClientForm()
        {
            InitializeComponent();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            var auth = new Authorization();
            this.Hide();
            auth.Closed += (s, args) => this.Close();
            auth.ShowDialog();
        }

        private void vihod_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (StreamWriter writer = new StreamWriter("Report.txt"))
            {
                writer.WriteLine(comboBox3.Text + " Ошибка в начислениях.");
                writer.Close();
            }
            MessageBox.Show("Сообщение об ошибке передано администратору на проверку", "");
        }
        public void select(string surname, string name, string patr, DataGridView datagrid)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            connection.Open();
            System.Data.DataTable dt = new System.Data.DataTable();
            string s = "SELECT Surname AS Фамилия, Workers.Name AS Имя, Patronymic AS Отчество, " +
                "Sex AS Пол, [Children count] AS [Количество детей], Tabel_numb AS [Табельный номер], " +
                "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
                "COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни] FROM Workers " +
                "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
                "JOIN Health ON Health.ID_wrk = Workers.ID_wrk " +
                "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Workers.ID_Post, Post.Name, [Children count]," +
                "Tabel_numb,  Post.Salary, Post.Director,Marital_status.Name " +
                $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}' ";
            SqlDataAdapter adapter = new SqlDataAdapter(s, connection);
            adapter.Fill(dt);
            datagrid.DataSource = dt;
            connection.Close();
        }
   
        private void button2_Click(object sender, EventArgs e)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
         
            if (!checkBox1.Checked)
            {
                string x = "";
                System.Data.DataTable dt = new System.Data.DataTable();
                connection.Open();
                string t = $"SELECT Post.Salary FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post  WHERE Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
                cmd = new SqlCommand(t, connection);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    x = (reader[0]).ToString();

                }
                reader.Close();
                connection.Close();
                string s = "SELECT Surname AS Фамилия, Workers.Name AS Имя, Patronymic AS Отчество, " +
               "Sex AS Пол, [Children count] AS [Количество детей], Tabel_numb AS [Табельный номер], " +
               "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
               "COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни] FROM Workers " +
               "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
               "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
               "JOIN Health ON Health.ID_wrk = Workers.ID_wrk " +
               "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Workers.ID_Post, Post.Name, [Children count]," +
               "Tabel_numb,  Post.Salary, Post.Director,Marital_status.Name " +
               $"HAVING Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";

                connection.Open();
                adapter = new SqlDataAdapter(s, connection);

                adapter.Fill(dt);
                dataGridView2.DataSource = dt;
                connection.Close();
            }
            else if(checkBox1.Checked)
            {
                string x = "";
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                string t = $"SELECT Post.Salary*0.5 FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post  WHERE Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
                cmd = new SqlCommand(t, connection);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    x = (reader[0]).ToString();

                }
                reader.Close();
                connection.Close();
                string s = "SELECT Surname AS Фамилия, Workers.Name AS Имя, Patronymic AS Отчество, " +
              "Sex AS Пол, [Children count] AS [Количество детей], Tabel_numb AS [Табельный номер], " +
              "Post.Name AS Должность, Post.Salary*0.5 AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
              "COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни] FROM Workers " +
              "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
              "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
              "JOIN Health ON Health.ID_wrk = Workers.ID_wrk " +
              "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Workers.ID_Post, Post.Name, [Children count]," +
              "Tabel_numb,  Post.Salary, Post.Director,Marital_status.Name " +
              $"HAVING Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
                connection.Open();
                adapter = new SqlDataAdapter(s, connection);

                adapter.Fill(dt);
                dataGridView2.DataSource = dt;
                connection.Close();
                
            
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            System.Windows.Forms.CheckBox checkBox1 = (System.Windows.Forms.CheckBox)sender;
        
                select(a[0], a[2], a[3], dataGridView2);
            

        }

        private void ClientForm_Load(object sender, EventArgs e)
        {
            connection.Open();
            string s = "SELECT COUNT(*) FROM Workers";
            cmd = new SqlCommand(s, connection);
          SqlDataReader  reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();
                Console.WriteLine(reader[0]);

            }

            reader.Close();
            connection.Close();

           int t = Convert.ToInt32(s);
          List<string>  count = new List<string>(); ;
            for (int i = 0; i < t; i++)
            {
                connection.Open();
                s = $"SELECT Workers.Surname, Workers.Name, Workers.Patronymic FROM Workers WHERE ID_wrk = {i + 1}";
                cmd = new SqlCommand(s, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    s = ((reader[0]).ToString() + reader[1].ToString() + reader[2].ToString());
                    Console.WriteLine(reader[0]);

                }
                reader.Close();
                connection.Close();
                String[] words = s.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                count.Add(words[0] + " " + words[1] + " " + words[2]);
            }
            foreach (var item in count)
            {
                if (!comboBox3.Items.Contains(item))
                {
                    this.comboBox3.Items.Add(item);
                }
               
                else { }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView2);
        }

    
    }
}