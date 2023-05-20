using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace cs_database_courseproject
{
    public partial class AccountantForm : Form
    {
       
        private readonly service.HealthService health = new service.HealthService();
        private readonly service.IncomeService inc = new service.IncomeService();
        private readonly service.AccuralsService accur = new service.AccuralsService();

        private readonly service.ReportService report = new service.ReportService();
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public AccountantForm()
        {
            InitializeComponent();

            inc.ShowIncome(SortBox3, dataInc);
            health.ShowHealth(SortBox4, dataHealth);
            accur.ShowAccurals(SortBox5, dataAccurals);
            this.dataInc.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataInc_RowHeaderMouseClick);
            this.dataHealth.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataHealth_RowHeaderMouseClick);
            this.dataAccurals.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataAccurals_RowHeaderMouseClick);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (totalField.Text != "")
                {
                    double ndfl = Convert.ToDouble(totalField.Text) * 0.13;
                    ndflLabel.Text = ndfl.ToString();
                    double total = Convert.ToDouble(totalField.Text) - ndfl;
                    totaltobepaidLabel.Text = total.ToString();
                }
                else MessageBox.Show("Не введены значения", "");

                connection.Open();

                string s = $"SELECT Marital_status.Name FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = {tabelField2.Text}";
                cmd = new SqlCommand(s, connection);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    label37.Text = (reader[0]).ToString();
                    Console.WriteLine(reader[0]);
                }

                connection.Close();
                reader.Close();
                connection.Open();


                string t = $"SELECT Post.Name FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = {tabelField2.Text}";
                cmd = new SqlCommand(t, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    label53.Text = (reader[0]).ToString();

                }
                reader.Close();
                connection.Close();


                connection.Open();

                t = $"SELECT Workers.Surname, Workers.Name, Workers.Patronymic FROM Workers" +
                        $" WHERE Workers.Tabel_numb = {tabelField2.Text}";
                cmd = new SqlCommand(t, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {

                    string str = ((reader[0]).ToString() + reader[1].ToString() + reader[2].ToString());
                    String[] words = str.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    label65.Text = (words[0] + " " + words[1] + " " + words[2]).ToString();
                }
                reader.Close();
                connection.Close();

                label52.Show(); label37.Show(); label53.Show();
                label54.Show();
                label62.Show(); label65.Show();

                string a = dateFiled1.Text;

                for (int i = 0; i < a.Length; i++)
                {
                    if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                       Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                    {
                        throw new Exception("Введенная последовательность символов не является датой");
                    }
                    if (a[i] == ' ') { }
                }


                a = totalField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if ((a[i] >= 'a' && a[i] <= 'z') || (a[i] >= 'а' && a[i] <= 'я') || (a[i] >= 'А' && a[i] <= 'Я') || (a[i] >= 'A' && a[i] <= 'Z'))
                    {
                        throw new Exception("Не допускаются буквенные символы");
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void Cleandatainc()
        {
            delFiled3.Text = "";
            tabelField2.Text = "";
            dateFiled1.Text = "";
            totalField.Text = "";
            ndflLabel.Text = "";
            totaltobepaidLabel.Text = "";
            MonthField.Text = "";
            label37.Text = "";
            label53.Text = "";

        }
        public void CleandataHealth()
        {
            IdField3.Text = "";
            numberField.Text = "";
            IllField.Text = "";
            OrgField.Text = "";
            DocField.Text = "";
            sickleaveField.Text = "";
            realeaseField.Text = "";
            TabelField3.Text = "";
            label55.Text = "";
            label57.Text = "";
        }
        public void CleandataAccurals()
        {
            IDField5.Text = "";
            dateField2.Text = "";
            CommentField.Text = "";
            amountField.Text = "";
            typeField.Text = "";
            tabelField4.Text = "";
            label64.Text = "";
            label63.Text = "";
        }
        private void Add2_Click(object sender, EventArgs e)
        {
            connection.Open();
            string post11 = "", wrk11 = "", ms11 = "";
            string s = $"SELECT Workers.ID_wrk FROM Workers WHERE Workers.Tabel_numb = {tabelField2.Text}";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                wrk11 = (reader[0]).ToString();

            }
            connection.Close();
            reader.Close();

            connection.Open();

            string t = $"SELECT Post.ID_Post FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = {tabelField2.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                post11 = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();

            connection.Open();


            string n = $"SELECT Marital_status.ID_Ms FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = {tabelField2.Text}";
            cmd = new SqlCommand(n, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ms11 = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();

            inc.AddIncome(dateFiled1.Text, totalField.Text, ndflLabel.Text, totaltobepaidLabel.Text,
                MonthField.Text, tabelField2.Text, wrk11, ms11, post11, SortBox3, dataInc);
            Cleandatainc();
        }
        private void dataInc_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            delFiled3.Text = dataInc.Rows[e.RowIndex].Cells[0].Value.ToString();
            tabelField2.Text = dataInc.Rows[e.RowIndex].Cells[1].Value.ToString();
            dateFiled1.Text = ((DateTime)dataInc.Rows[e.RowIndex].Cells[7].Value).ToShortDateString();
            totalField.Text = dataInc.Rows[e.RowIndex].Cells[8].Value.ToString();
            ndflLabel.Text = dataInc.Rows[e.RowIndex].Cells[9].Value.ToString();
            totaltobepaidLabel.Text = dataInc.Rows[e.RowIndex].Cells[10].Value.ToString();
            MonthField.Text = dataInc.Rows[e.RowIndex].Cells[11].Value.ToString();

            label52.Hide(); label37.Hide(); label53.Hide();
            label54.Hide();
        }
        private void dataHealth_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            IdField3.Text = dataHealth.Rows[e.RowIndex].Cells[0].Value.ToString();
            numberField.Text = dataHealth.Rows[e.RowIndex].Cells[7].Value.ToString();
            IllField.Text = dataHealth.Rows[e.RowIndex].Cells[8].Value.ToString();
            OrgField.Text = dataHealth.Rows[e.RowIndex].Cells[9].Value.ToString();
            DocField.Text = dataHealth.Rows[e.RowIndex].Cells[10].Value.ToString();
            sickleaveField.Text = ((DateTime)dataHealth.Rows[e.RowIndex].Cells[11].Value).ToShortDateString();
            realeaseField.Text = ((DateTime)dataHealth.Rows[e.RowIndex].Cells[12].Value).ToShortDateString();

            TabelField3.Text = dataHealth.Rows[e.RowIndex].Cells[1].Value.ToString();
            label56.Hide(); label58.Hide(); label57.Hide(); label55.Hide();

        }
        private void dataAccurals_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            IDField5.Text = dataAccurals.Rows[e.RowIndex].Cells[0].Value.ToString();
            dateField2.Text = ((DateTime)dataAccurals.Rows[e.RowIndex].Cells[7].Value).ToShortDateString();
            CommentField.Text = dataAccurals.Rows[e.RowIndex].Cells[8].Value.ToString();
            amountField.Text = dataAccurals.Rows[e.RowIndex].Cells[9].Value.ToString();
            typeField.Text = dataAccurals.Rows[e.RowIndex].Cells[6].Value.ToString();

            tabelField4.Text = dataAccurals.Rows[e.RowIndex].Cells[1].Value.ToString();
            label60.Hide();
            label59.Hide();
            label64.Hide();
            label63.Hide();

        }

        private void Add3_Click(object sender, EventArgs e)
        {
            connection.Open();
            string post11 = "", wrk11 = "", ms11 = "";
            string s = $"SELECT Workers.ID_wrk FROM Workers WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                wrk11 = (reader[0]).ToString();

            }
            connection.Close();
            reader.Close();

            connection.Open();

            string t = $"SELECT Post.ID_Post FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                post11 = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();

            connection.Open();


            string n = $"SELECT Marital_status.ID_Ms FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(n, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ms11 = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();
            health.AddHealth(numberField.Text, OrgField.Text, DocField.Text, TabelField3.Text, sickleaveField.Text,
                realeaseField.Text, IllField.Text, wrk11, post11, ms11, SortBox4, dataHealth); CleandataHealth();
        }

        private void Add5_Click(object sender, EventArgs e)
        {
            accur.AddAccurals(dateField2.Text, CommentField.Text, amountField.Text, tabelField4.Text, typeField.Text, SortBox5, dataAccurals);
            CleandataAccurals();
        }

        private void Change2_Click(object sender, EventArgs e)
        {
            inc.ChangeIncome(dateFiled1.Text, totalField.Text, ndflLabel.Text, totaltobepaidLabel.Text,
               MonthField.Text, delFiled3.Text, tabelField2.Text, SortBox3, dataInc);

            tabelField2.Items.Clear();
            connection.Open();
            string s = "SELECT COUNT(*) FROM Workers";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();
                Console.WriteLine(reader[0]);

            }

            reader.Close();
            connection.Close();

            int t = Convert.ToInt32(s);
            List<string> count = new List<string>(); ;
            for (int i = 0; i < t; i++)
            {
                connection.Open();
                s = $"SELECT Workers.Tabel_numb FROM Workers WHERE ID_wrk = {i + 1} ";
                cmd = new SqlCommand(s, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    s = (reader[0]).ToString();
                    Console.WriteLine(reader[0]);

                }
                reader.Close();
                connection.Close();
                count.Add(s);
            }
            foreach (var item in count)
            {
                if (!tabelField2.Items.Contains(item))
                {
                    this.tabelField2.Items.Add(item);
                }
                else { }
            }
            Cleandatainc();
        }

        private void Change3_Click(object sender, EventArgs e)
        {
            health.ChangeHealth(numberField.Text, OrgField.Text, DocField.Text, TabelField3.Text, sickleaveField.Text,
              realeaseField.Text, IllField.Text, IdField3.Text, SortBox4, dataHealth);
            CleandataHealth();
        }

        private void Del3_Click(object sender, EventArgs e)
        {

            health.DeleteHealth(IdField3.Text, SortBox4, dataHealth); CleandataHealth();
        }

        private void Change5_Click(object sender, EventArgs e)
        {
            accur.ChangeAccurals(dateField2.Text, CommentField.Text, amountField.Text, tabelField4.Text, typeField.Text, IDField5.Text, SortBox5, dataAccurals);

            CleandataAccurals();
        }

        private void Del5_Click(object sender, EventArgs e)
        {
            accur.DeleteAccurals(IDField5.Text, SortBox5, dataAccurals);
            CleandataAccurals();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string a = numberField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (!(a[i] >= '0' && a[i] <= '9'))
                {
                    throw new Exception("Введенная последовательность символов не является номером");
                }
                if (a[i] == ' ') { }
            }
            a = sickleaveField.Text;

            for (int i = 0; i < a.Length; i++)
            {
                if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                   Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                {
                    throw new Exception("Введенная последовательность символов не является датой");
                }
                if (a[i] == ' ') { }
            }

            a = realeaseField.Text;

            for (int i = 0; i < a.Length; i++)
            {
                if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                   Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                {
                    throw new Exception("Введенная последовательность символов не является датой");
                }
                if (a[i] == ' ') { }

            }
            AdminForm form = new AdminForm();
            form.SpellChecking(DocField, "Фамилия доктора должна", "фамилии доктора", "фамилией доктора");
            connection.Open();
            string s = $"SELECT Marital_status.Name FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                label57.Text = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();
            connection.Open();


            string t = $"SELECT Post.Name FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                label55.Text = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }
            reader.Close();
            connection.Close();

            connection.Open();

            t = $"SELECT Workers.Surname, Workers.Name, Workers.Patronymic FROM Workers" +
                    $" WHERE Workers.Tabel_numb = {TabelField3.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string str = ((reader[0]).ToString() + reader[1].ToString() + reader[2].ToString());
                String[] words = str.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                label66.Text = (words[0] + " " + words[1] + " " + words[2]).ToString();
            }
            reader.Close();
            connection.Close();
            label56.Show(); label58.Show(); label57.Show(); label55.Show(); label67.Show(); label66.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            connection.Open();
            string a = dateField2.Text;

            for (int i = 0; i < a.Length; i++)
            {
                if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                   Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                {
                    throw new Exception("Введенная последовательность символов не является датой");
                }
                if (a[i] == ' ') { }

            }
            AdminForm form = new AdminForm();
            form.SpellChecking(typeField, "Наименование вида начисления должно", "наименовании вида начисления", "наименованием вида начисления");
            string s = $"SELECT Marital_status.Name FROM Marital_status JOIN Workers ON Workers.ID_Ms = Marital_status.ID_Ms WHERE Workers.Tabel_numb = {tabelField4.Text}";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                label64.Text = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }

            connection.Close();
            reader.Close();
            connection.Open();


            string t = $"SELECT Post.Name FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Tabel_numb = {tabelField4.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                label63.Text = (reader[0]).ToString();
                Console.WriteLine(reader[0]);
            }
            reader.Close();
            connection.Close();
            connection.Open();

            t = $"SELECT Workers.Surname, Workers.Name, Workers.Patronymic FROM Workers" +
                    $" WHERE Workers.Tabel_numb = {tabelField4.Text}";
            cmd = new SqlCommand(t, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string str = ((reader[0]).ToString() + reader[1].ToString() + reader[2].ToString());
                String[] words = str.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                label68.Text = (words[0] + " " + words[1] + " " + words[2]).ToString();
            }
            reader.Close();
            connection.Close();
            label68.Show();
            label69.Show();
            label60.Show();
            label59.Show();
            label64.Show();
            label63.Show();
            label59.Show();
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

        private void button36_Click(object sender, EventArgs e)
        {
            String[] a = comboBox1.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView1);
        }

        private void button37_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataGridView1);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataGridView1);
        }

        private void excel6_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataAccurals);
        }

        private void word6_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataAccurals);
        }

        private void excel4_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataHealth); ;
        }

        private void word4_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataHealth);
        }

        private void excel3_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataInc);
        }

        private void word3_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataInc);
        }
        public void select(string surname, string name, string patr, DataGridView datagrid)
        {
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
                $"HAVING Workers.Surname = '{surname}' AND Workers.Name = '{name}' AND Workers.Patronymic = '{patr}'" ;
            SqlDataAdapter adapter = new SqlDataAdapter(s, connection);
            adapter.Fill(dt);
            datagrid.DataSource = dt;
            connection.Close();
        }
      
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            System.Windows.Forms.CheckBox checkBox1 = (System.Windows.Forms.CheckBox)sender; 
      
                select(a[0], a[2], a[3],  dataGridView2);
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (!checkBox1.Checked)
            {
              
                string x = "";
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                string t = $"SELECT Post.Salary FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
                cmd = new SqlCommand(t, connection);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    x = (reader[0]).ToString();

                }
                reader.Close();
                connection.Close();
                connection.Open();
                string s = "SELECT Surname AS Фамилия, Workers.Name AS Имя, Patronymic AS Отчество, " +
                "Sex AS Пол, [Children count] AS [Количество детей], Tabel_numb AS [Табельный номер], " +
                "Post.Name AS Должность, Post.Salary AS Оклад, Post.Director AS Директор, Marital_status.Name AS [Семейное положение]," +
                "COUNT(Health.[Sick leave date]) AS [Кол-во пропусков по болезни] FROM Workers " +
                "LEFT JOIN Post ON Workers.ID_Post=Post.ID_Post " +
                "LEFT JOIN Marital_status ON Workers.ID_Ms=Marital_status.ID_Ms " +
                "JOIN Health ON Health.ID_wrk = Workers.ID_wrk " +
                "GROUP BY Workers.Surname, Workers.Name,Workers.Patronymic, Workers.Sex, Workers.ID_Post, Post.Name, [Children count]," +
                "Tabel_numb,  Post.Salary, Post.Director,Marital_status.Name " +
                $"HAVING Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}' ";

                MessageBox.Show($"Сотруднику  {a[0]} {a[1]} {a[2]}" +
                    $"начислена выплата в размере {x} ", "");
                adapter = new SqlDataAdapter(s, connection);

                adapter.Fill(dt);
                dataGridView2.DataSource = dt;
                connection.Close();
            }
            else
            {
                string x = "";
                connection.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                string t = $"SELECT Post.Salary*0.5 FROM Post JOIN Workers ON Workers.ID_Post = Post.ID_Post WHERE Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
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
               $"HHAVING Workers.Surname = '{a[0]}' AND Workers.Name = '{a[1]}' AND Workers.Patronymic = '{a[2]}'";
                MessageBox.Show($"Сотруднику  {a[0]} {a[1]} {a[2]}" +
                                   $"начислена выплата в размере {x} ", "");
                connection.Open();
                adapter = new SqlDataAdapter(s, connection);

                adapter.Fill(dt);
                dataGridView2.DataSource = dt;
                connection.Close();
            }
        }

        private void AccountantForm_Load(object sender, EventArgs e)
        {

            label60.Hide();
            label64.Hide();
            label68.Hide();
            label63.Hide();
            label59.Hide();
            label69.Hide();
            label54.Hide();
            label52.Hide();
            label62.Hide();

            label58.Hide();
            label56.Hide();
            label67.Hide();
            label66.Hide();
            connection.Open();
            string s = "SELECT COUNT(*) FROM Post";
            cmd = new SqlCommand(s, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();


            }

            reader.Close();
            connection.Close();

            int t = Convert.ToInt32(s);
            List<string> count = new List<string>(); ;
            for (int i = 0; i < t; i++)
            {
                connection.Open();
                s = $"SELECT Post.Name FROM Post WHERE ID_Post = {i + 1} ";
                cmd = new SqlCommand(s, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    s = (reader[0]).ToString();


                }
                reader.Close();
                connection.Close();
                count.Add(s);
            }
            foreach (var item in count)
            {

                if (!comboBox2.Items.Contains(item))
                {
                    this.comboBox2.Items.Add(item);
                }
                else { }
            }


            connection.Open();
            s = "SELECT COUNT(*) FROM Workers";
            cmd = new SqlCommand(s, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();
                Console.WriteLine(reader[0]);

            }

            reader.Close();
            connection.Close();

            t = Convert.ToInt32(s);
            count = new List<string>(); ;
            for (int i = 0; i < t; i++)
            {
                connection.Open();
                s = $"SELECT Workers.Tabel_numb FROM Workers WHERE ID_wrk = {i + 1} ";
                cmd = new SqlCommand(s, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    s = (reader[0]).ToString();
                    Console.WriteLine(reader[0]);

                }
                reader.Close();
                connection.Close();
                count.Add(s);
            }
            foreach (var item in count)
            {
                if (!tabelField2.Items.Contains(item))
                {
                    this.tabelField2.Items.Add(item);
                }
                else { }
                if (!TabelField3.Items.Contains(item))
                {
                    this.TabelField3.Items.Add(item);
                }
                else { }
                if (!tabelField4.Items.Contains(item))
                {
                    this.tabelField4.Items.Add(item);
                }
                else { }
            }

            connection.Open();
            s = "SELECT COUNT(*) FROM Workers";
            cmd = new SqlCommand(s, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();
                Console.WriteLine(reader[0]);

            }

            reader.Close();
            connection.Close();

            t = Convert.ToInt32(s);
            count = new List<string>(); ;
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
                if (!comboBox1.Items.Contains(item))
                {
                    this.comboBox1.Items.Add(item);
                }
                else { }
                if (!comboBox3.Items.Contains(item)) 
                {
                    this.comboBox3.Items.Add(item);
                }
                else { }
            }
        }

        private void newData2_Click(object sender, EventArgs e)
        {
            inc.ShowIncome(SortBox3, dataInc);
        }

        private void newData3_Click(object sender, EventArgs e)
        {
            health.ShowHealth(SortBox4, dataHealth);
        }

        private void newData5_Click(object sender, EventArgs e)
        {
            accur.ShowAccurals(SortBox5, dataAccurals);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataGridView3);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataGridView3);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String[] a = comboBox1.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView1);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            String[] a = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView2);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            report.ReportCity(textBox2.Text,textBox6.Text, dataGridView3);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            report.ReportHealth(textBox3.Text, dataGridView3);
        }

      
        private void button12_Click(object sender, EventArgs e)
        {
            report.ReportPost(comboBox2.Text, dataGridView3);
        }


    }
}
