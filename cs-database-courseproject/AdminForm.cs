using cs_database_courseproject.users;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace cs_database_courseproject
{
    public partial class AdminForm : Form
    {
        private readonly service.WorkerService worker = new service.WorkerService();
        private readonly service.HealthService health = new service.HealthService();
        private readonly service.IncomeService inc = new service.IncomeService();
        private readonly service.AccuralsService accur = new service.AccuralsService();
        private readonly service.PostService post = new service.PostService();
        private readonly service.Type_of_accuralsService type = new service.Type_of_accuralsService();
        private readonly service.Marital_statusService status = new service.Marital_statusService();
        private readonly service.ReportService report = new service.ReportService();
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public AdminForm()
        {
            InitializeComponent();
   
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
            worker.ShowWorkers(SortBox1, dataWrk);
            post.ShowPost(SortBox2, dataPost);
            inc.ShowIncome(SortBox3, dataInc);
            health.ShowHealth(SortBox4, dataHealth);
            accur.ShowAccurals(SortBox5, dataAccurals);
            status.ShowStatus(sorrt1, dataMarital);
            type.ShowTypes(sort, dataType);

     
            this.dataWrk.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataWrk_RowHeaderMouseClick);
            this.dataPost.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataPost_RowHeaderMouseClick);
            this.dataInc.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataInc_RowHeaderMouseClick);
            this.dataHealth.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataHealth_RowHeaderMouseClick);
            this.dataMarital.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataMarital_RowHeaderMouseClick); ;
            this.dataAccurals.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataAccurals_RowHeaderMouseClick);
            this.dataType.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataType_RowHeaderMouseClick);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, 25, 25);
            Region rgn = new Region(path);
          button8.Region = rgn;
            button9.Region = rgn;
            button10.Region = rgn;
           
            button12.Region = rgn;
            for (int i = 0; i < dataWrk.Rows.Count; i++)
            {
                for (int j = 0; j > dataWrk.ColumnCount; j++)
                {
                    dataWrk.Rows[i].Cells[j].Style.BackColor = SystemColors.Control;
                }
            }
 

            dataWrk.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataWrk.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataPost.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataPost.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataInc.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataInc.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataHealth.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataHealth.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataMarital.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataMarital.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataType.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataType.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataAccurals.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataAccurals.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataGridView2.RowHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dataPost.EnableHeadersVisualStyles = false;
            dataInc.EnableHeadersVisualStyles = false;
            dataWrk.EnableHeadersVisualStyles = false;
            dataHealth.EnableHeadersVisualStyles = false;
            dataAccurals.EnableHeadersVisualStyles = false;
            dataMarital.EnableHeadersVisualStyles = false;
            dataType.EnableHeadersVisualStyles = false;
            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView2.EnableHeadersVisualStyles = false;

            button13.Region = rgn;
            button11.Region = rgn;
            button14.Region = rgn;
            button15.Region = rgn;
            dataWrk.BackgroundColor = this.button13.BackColor;
            dataPost.BackgroundColor = this.button13.BackColor;
            dataInc.BackgroundColor = this.button13.BackColor;
            dataHealth.BackgroundColor = this.button13.BackColor;
            dataMarital.BackgroundColor = this.button13.BackColor;
            dataAccurals.BackgroundColor = this.button13.BackColor;
            dataType.BackgroundColor = this.button13.BackColor;
            dataGridView1.BackgroundColor = this.button13.BackColor;
            dataGridView2.BackgroundColor = this.button13.BackColor;


            dataWrk.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor;
            dataPost.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
               dataInc.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
            dataHealth.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
             dataMarital.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
                dataAccurals.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
               dataType.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
             dataGridView1.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; 
                 dataGridView2.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor;

            dataWrk.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; 
               dataPost.DefaultCellStyle.SelectionForeColor = this.button14.BackColor;
                 dataInc.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; 
                dataHealth.DefaultCellStyle.SelectionForeColor = this.button14.BackColor;
                 dataMarital.DefaultCellStyle.SelectionForeColor = this.button14.BackColor;
                 dataAccurals.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; 
                dataType.DefaultCellStyle.SelectionForeColor = this.button14.BackColor;
                 dataGridView1.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; 
                 dataGridView2.DefaultCellStyle.SelectionForeColor = this.button14.BackColor;

            dataWrk.DefaultCellStyle.SelectionBackColor = this.button11.BackColor;
       dataPost.DefaultCellStyle.SelectionBackColor = this.button11.BackColor;
               dataInc.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 
              dataHealth.DefaultCellStyle.SelectionBackColor = this.button11.BackColor;
              dataMarital.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 
               dataAccurals.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 
             dataType.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 
               dataGridView1.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 
               dataGridView2.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; 


    }
        bool sort = false;
        bool sorrt1 = false;
        private void Form1_Load(object sender, EventArgs e)
        {
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
                if (!PstField.Items.Contains(item))
                {
                    this.PstField.Items.Add(item);
                }
                if (!comboBox2.Items.Contains(item))
                {
                    this.comboBox2.Items.Add(item);
                }
                else { }
            }

            connection.Open();
            s = "SELECT COUNT(*) FROM Marital_status";
            cmd = new SqlCommand(s, connection);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                s = (reader[0]).ToString();

            }
            reader.Close();
            connection.Close();

            t = Convert.ToInt32(s);

            count = new List<string>(); ;
            for (int i = 0; i < t; i++)
            {
                connection.Open();
                s = $"SELECT Marital_status.Name FROM Marital_status WHERE ID_Ms = {i + 1}";
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
                if (!msField.Items.Contains(item))
                {
                    this.msField.Items.Add(item);
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
                s = $"SELECT Workers.Surname, Workers.Name, Workers.Patronymic FROM Workers WHERE ID_wrk = {i+1}";
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
                count.Add(words[0] +" "+ words[1] +" "+ words[2]);
            }
            foreach (var item in count)
            {
                if (!comboBox1.Items.Contains(item))
                {
                    this.comboBox1.Items.Add(item);
                }
                else { }
               
            }
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

        private void Add6_Click(object sender, EventArgs e)
        {
            type.AddTypes(AccurField.Text, sort, dataType);
            CleandataType();
        }

        private void Change6_Click(object sender, EventArgs e)
        {
            type.ChangeTypes(AccurField.Text, textBox4.Text, sort, dataType); CleandataType();
        }

        private void Del6_Click(object sender, EventArgs e)
        {
            type.DeleteTypes(textBox4.Text, sort, dataType); CleandataType();
        }

        private void truncatetable6_Click(object sender, EventArgs e)
        {
            type.TruncateTable(sort, dataType); CleandataType();
        }

        private void truncateBut1_Click(object sender, EventArgs e)
        {
            post.TruncateTable(SortBox2, dataPost); CleandataPost();
        }

        private void Add1_Click(object sender, EventArgs e)
        {
            try
            {
                string a = postField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[0] == a.ToLower()[0]) { throw new Exception($"Наименование должности должно начинаться с прописной буквы!"); }

                    if (a[i] == ' ') { }
                }
                a = DirectorField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[0] == a.ToLower()[0]) { throw new Exception($"Обозначение директора должно начинаться с прописной буквы!"); }
                    if (a[0] >= 'A' && a[0] <= 'Z')
                    {

                        throw new Exception("Допускаются только символы русского алфавита");

                    }
                    if (a[0] >= 'А' && a[0] <= 'Я')
                    {
                        if (a[i] >= 'a' && a[i] <= 'z')
                        {
                            throw new Exception("Допускаются только символы русского алфавита");
                        }
                      
                    }
                    if (a[i] >= '0' && a[i] <= '9')
                    {

                        throw new Exception($"Введенная последовательность символов не является обозначением директора");
                    }
                    if (a[i] == '@' || a[i] == '$' || a[i] == '!' || a[i] == '?' || a[i] == '#'
                        || a[i] == ';' || a[i] == ':' || a[i] == ',' || a[i] == '.' || a[i] == '/'
                        || a[i] == '|' || a[i] == '*' || a[i] == '&' || a[i] == '"' || a[i] == '`' || a[i] == '~' || a[i] == '^'
                        || a[i] == '>' || a[i] == '<')
                    {
                        throw new Exception($"Введенная последовательность символов не является обозначением директора");
                    }
                    if (a[i] == ' ') { }
                }
                post.AddPost(postField.Text, SalaryField.Text, DirectorField.Text, SortBox2, dataPost);
                CleandataPost();
                connection.Open();
                string s = "SELECT COUNT(*) FROM Post";
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
                    s = $"SELECT Post.Name FROM Post WHERE ID_Post = {i + 1} ";
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
                    if (!PstField.Items.Contains(item))
                    {
                        this.PstField.Items.Add(item);
                    }
                    if (!comboBox2.Items.Contains(item))
                    {
                        this.comboBox2.Items.Add(item);
                    }
                    else { }
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, " "); }
        }

        private void Change1_Click(object sender, EventArgs e)
        {
            try
            {
                string a = postField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[0] == a.ToLower()[0]) { throw new Exception($"Наименование должности должно начинаться с прописной буквы!"); }

                    if (a[i] == ' ') { }
                }
                a = DirectorField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[0] == a.ToLower()[0]) { throw new Exception($"Обозначение директора должно начинаться с прописной буквы!"); }
                    if (a[0] >= 'A' && a[0] <= 'Z')
                    {

                        throw new Exception("Допускаются только символы русского алфавита");

                    }
                    if (a[0] >= 'А' && a[0] <= 'Я')
                    {
                        if (a[i] >= 'a' && a[i] <= 'z')
                        {
                            throw new Exception("Допускаются только символы русского алфавита");
                        }

                    }
                    if (a[i] >= '0' && a[i] <= '9')
                    {

                        throw new Exception($"Введенная последовательность символов не является обозначением директора");
                    }
                    if (a[i] == '@' || a[i] == '$' || a[i] == '!' || a[i] == '?' || a[i] == '#'
                        || a[i] == ';' || a[i] == ':' || a[i] == ',' || a[i] == '.' || a[i] == '/'
                        || a[i] == '|' || a[i] == '*' || a[i] == '&' || a[i] == '"' || a[i] == '`' || a[i] == '~' || a[i] == '^'
                        || a[i] == '>' || a[i] == '<')
                    {
                        throw new Exception($"Введенная последовательность символов не является обозначением директора");
                    }
                    if (a[i] == ' ') { }
                }
                post.ChangePost(postField.Text, SalaryField.Text, DirectorField.Text, IdFiled2.Text, SortBox2, dataPost);
                CleandataPost();
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, ""); }
        }

        private void Del1_Click(object sender, EventArgs e)
        {
            post.DeletePost(SortBox2, IdFiled2.Text, dataPost);
            PstField.Items.Clear();
            connection.Open();
            string s = "SELECT COUNT(*) FROM Post";
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
                s = $"SELECT Post.Name FROM Post WHERE ID_Post = {i + 1} ";
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
                if (!PstField.Items.Contains(item))
                {
                    this.PstField.Items.Add(item);
                }
                if (!comboBox2.Items.Contains(item))
                {
                    this.comboBox2.Items.Add(item);
                }
                else { }
            }

            CleandataPost();
        }

        private void Add4_Click(object sender, EventArgs e)
        {
            status.AddMs(MaritalField.Text, sorrt1, dataMarital);
            connection.Open();
            string s = "SELECT COUNT(*) FROM Marital_status";
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
                s = $"SELECT Marital_status.Name FROM Marital_status WHERE ID_Ms = {i + 1}";
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
                if (!msField.Items.Contains(item))
                {
                    this.msField.Items.Add(item);
                }
                else { }
            }

            CleandataMarital();
        }

        private void Change4_Click(object sender, EventArgs e)
        {
            status.ChangeMs(MaritalField.Text, IdField4.Text, sorrt1, dataMarital);
            CleandataMarital();
        }

        private void Del4_Click(object sender, EventArgs e)
        {
            status.DeleteStatus(IdField4.Text, sorrt1, dataMarital);
            msField.Items.Clear();
            connection.Open();
            string s = "SELECT COUNT(*) FROM Marital_status";
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
                s = $"SELECT Marital_status.Name FROM Marital_status WHERE ID_Ms = {i + 1}";
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
                if (!msField.Items.Contains(item))
                {
                    this.msField.Items.Add(item);
                }
                else { }
            }
            CleandataMarital();
        }

        private void truncate4_Click(object sender, EventArgs e)
        {
            status.truncateTable(sorrt1, dataMarital);
            CleandataMarital();
        }

        private void sort1_Click(object sender, EventArgs e)
        {
            sorrt1 = true;
            status.ShowStatus(sorrt1, dataMarital);
            CleandataMarital();
            sorrt1 = false;
        }

        private void Sort2_Click(object sender, EventArgs e)
        {
            sort = true;
            type.ShowTypes(sort, dataType);
            CleandataType();
            sort = false;
        }

        private void Add5_Click(object sender, EventArgs e)
        {
            accur.AddAccurals(dateField2.Text, CommentField.Text, amountField.Text, tabelField4.Text, typeField.Text, SortBox5, dataAccurals);
            CleandataAccurals();
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

        private void truncatetable5_Click(object sender, EventArgs e)
        {
            accur.TruncateTable(SortBox5, dataAccurals);
            CleandataAccurals();
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            try
            {
                SpellChecking(SurnameField, "Фамилия должна", "фамилии", "фамилией");
                SpellChecking(NameField, "Имя должно", "имени", "именем");
                SpellChecking(PatField, "Отчество должно", "отчестве", "отчеством");

              string a = ChildField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (!(a[i] >= '0' && a[i] <= '9'))
                    {
                        throw new Exception("Введенная последовательность символов не является числом");
                    }
                    if (a[i] == ' ') { }
                }
              
     a = BirthField.Text;
           
                for (int i = 0; i < a.Length; i++)
                {
                 if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                    Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                        {
                        throw new Exception("Введенная последовательность символов не является датой");
                    }
                    if (a[i] == ' ') { }
                   
                   
                }
                a = PhoneField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (!(a[i] >= '0' && a[i] <= '9'))
                    {
                        throw new Exception("Введенная последовательность символов не является номером");
                    }
                    if (a[i] == ' ') { }
                }
                a = AdressField.Text;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[0] >= 'A' && a[0] <= 'Z')
                    {

                        throw new Exception("Допускаются только символы русского алфавита");

                    }
                    if (a[0] >= 'А' && a[0] <= 'Я')
                    {
                        if (a[i] >= 'a' && a[i] <= 'z')
                        {
                            throw new Exception("Допускаются только символы русского алфавита");
                        }

                    }
                    if (a[i] == ' ') { }
                }
                    a = TabelField.Text;
                    for (int i = 0; i < a.Length; i++)
                    {
                        if (!(a[i] >= '0' && a[i] <= '9'))
                        {
                            throw new Exception("Введенная последовательность символов не является номером");
                        }
                    if (a[i] == ' ') { }
                }

                connection.Open();
                string post11 = " ";
                string k = $"SELECT Post.ID_Post FROM Post WHERE Post.Name = '{PstField.Text}'";
                cmd = new SqlCommand(k, connection);
              SqlDataReader  reader = cmd.ExecuteReader();
                while (reader.Read())
                {
               post11 = (reader[0]).ToString();
                    Console.WriteLine(reader[0]);
                }

                connection.Close();
                reader.Close();

                connection.Open();
                string ms11 = "";

                string n = $"SELECT Marital_status.ID_Ms FROM Marital_status WHERE Marital_status.Name = '{msField.Text}'";
                cmd = new SqlCommand(n, connection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                   ms11 = (reader[0]).ToString();
                    Console.WriteLine(reader[0]);
                }

                connection.Close();
                reader.Close();

                worker.AddWorker(NameField.Text, SurnameField.Text, PatField.Text, SexField.Text, ChildField.Text, BirthField.Text,
                    PhoneField.Text, AdressField.Text, TabelField.Text, emailField.Text, post11, ms11, SortBox1, dataWrk);


                connection.Open();
                string s = "SELECT COUNT(*) FROM Workers";
                cmd = new SqlCommand(s, connection);
            reader = cmd.ExecuteReader();
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
                CleandataWrk();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, ""); }
        }

        private void ChangeButton_Click(object sender, EventArgs e)
        {
            SpellChecking(SurnameField, "Фамилия должна", "фамилии", "фамилией");
            SpellChecking(NameField, "Имя должно", "имени", "именем");
            SpellChecking(PatField, "Отчество должно", "отчестве", "отчеством");

          string a = ChildField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (!(a[i] >= '0' && a[i] <= '9'))
                {
                    throw new Exception("Введенная последовательность символов не является числом");
                }
                if (a[i] == ' ') { }

            }
            a = BirthField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$")|| 
                    Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                {
                    throw new Exception("Введенная последовательность символов не является датой");
                }
                if (a[i] == ' ') { }
            }
            a = PhoneField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (!(a[i] >= '0' && a[i] <= '9'))
                {
                    throw new Exception("Введенная последовательность символов не является номером");
                }
                if (a[i] == ' ') { }
            }
            a = AdressField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[0] >= 'A' && a[0] <= 'Z')
                {

                    throw new Exception("Допускаются только символы русского алфавита");

                }
                if (a[0] >= 'А' && a[0] <= 'Я')
                {
                    if (a[i] >= 'a' && a[i] <= 'z')
                    {
                        throw new Exception("Допускаются только символы русского алфавита");
                    }

                }
                if (a[i] == ' ') { }
            }
            a = TabelField.Text;
            for (int i = 0; i < a.Length; i++)
            {
                if (!(a[i] >= '0' && a[i] <= '9'))
                {
                    throw new Exception("Введенная последовательность символов не является номером");
                }
                if (a[i] == ' ') { }
            }


            worker.ChangeWorker(NameField.Text, SurnameField.Text, PatField.Text, SexField.Text, ChildField.Text, BirthField.Text,
                PhoneField.Text, AdressField.Text, TabelField.Text, emailField.Text, PstField.Text, msField.Text, IdField.Text, SortBox1, dataWrk);
            CleandataWrk();

        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            worker.DeleteWorker(SortBox1, IdField.Text, dataWrk);
            CleandataWrk();
            tabelField2.Items.Clear();
            TabelField3.Items.Clear();
            tabelField4.Items.Clear();
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
        }

        private void truncateButton_Click(object sender, EventArgs e)
        {
            worker.truncateTable(SortBox1, dataWrk);
            CleandataWrk();
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

        private void Del2_Click(object sender, EventArgs e)
        {
            inc.DeleteIncome(delFiled3.Text, SortBox3, dataInc);
            Cleandatainc();
        }

        private void Truncatebut3_Click(object sender, EventArgs e)
        {
            inc.TruncateTable(SortBox3, dataInc); Cleandatainc();
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

        private void TruncateTable3_Click(object sender, EventArgs e)
        {
            health.TruncateTable(SortBox4, dataHealth); CleandataHealth();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            String[] a = comboBox1.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView1);
        }

        private void SortBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            worker.ShowWorkers(SortBox1, dataWrk); CleandataWrk();
        }

        private void SortBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            post.ShowPost(SortBox2, dataPost); CleandataPost();
        }

        private void SortBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            inc.ShowIncome(SortBox3, dataInc); Cleandatainc();
        }

        private void SortBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            health.ShowHealth(SortBox4, dataHealth); CleandataHealth();
        }

        private void SortBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            accur.ShowAccurals(SortBox5, dataAccurals); CleandataAccurals();
        }

        private void excel1_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataWrk);
        }

        private void excel2_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataPost);
        }

        private void excel3_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataInc);
        }

        private void excel4_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataHealth);
        }

        private void excel5_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataMarital);
        }

        private void excel6_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataAccurals);
        }

        private void excel7_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataType);
        }

        private void button37_Click(object sender, EventArgs e)
        {
            report.ExportExcel(dataGridView1);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataGridView1);
        }

        private void word7_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataType);
        }

        private void word6_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataAccurals);
        }

        private void word5_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataMarital);
        }

        private void word4_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataHealth);
        }

        private void word3_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataInc);
        }

        private void word2_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataPost);
        }

        private void word1_Click(object sender, EventArgs e)
        {
            report.ExportWord(dataWrk);
         
        }
        private void dataWrk_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            IdField.Text = dataWrk.Rows[e.RowIndex].Cells[0].Value.ToString();
            NameField.Text = dataWrk.Rows[e.RowIndex].Cells[2].Value.ToString();
            SurnameField.Text = dataWrk.Rows[e.RowIndex].Cells[1].Value.ToString();
            PatField.Text = dataWrk.Rows[e.RowIndex].Cells[3].Value.ToString();
            SexField.Text = dataWrk.Rows[e.RowIndex].Cells[4].Value.ToString();
            ChildField.Text = dataWrk.Rows[e.RowIndex].Cells[5].Value.ToString();
          BirthField.Text =  ((DateTime)dataWrk.Rows[e.RowIndex].Cells[6].Value).ToShortDateString();
            PhoneField.Text = dataWrk.Rows[e.RowIndex].Cells[7].Value.ToString();
            AdressField.Text = dataWrk.Rows[e.RowIndex].Cells[8].Value.ToString();
            TabelField.Text = dataWrk.Rows[e.RowIndex].Cells[9].Value.ToString();
            emailField.Text = dataWrk.Rows[e.RowIndex].Cells[10].Value.ToString();
            PstField.Text = dataWrk.Rows[e.RowIndex].Cells[11].Value.ToString();
            msField.Text = dataWrk.Rows[e.RowIndex].Cells[14].Value.ToString();
        }
        public void CleandataWrk()
        {
            IdField.Text = "";
            NameField.Text = "";
            SurnameField.Text = "";
            PstField.Text = "";
            SexField.Text = "";
            ChildField.Text = "";
            BirthField.Text = "";
            PhoneField.Text = "";
            AdressField.Text = "";
            TabelField.Text = "";
            emailField.Text = "";
            PatField.Text = "";
            msField.Text = "";
        }
        private void dataPost_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            IdFiled2.Text = dataPost.Rows[e.RowIndex].Cells[0].Value.ToString();
            postField.Text = dataPost.Rows[e.RowIndex].Cells[1].Value.ToString();
            SalaryField.Text = dataPost.Rows[e.RowIndex].Cells[2].Value.ToString();
            DirectorField.Text = dataPost.Rows[e.RowIndex].Cells[3].Value.ToString();

        }
        public void CleandataPost()
        {
            IdFiled2.Text = "";
            postField.Text = "";
            SalaryField.Text = "";
            DirectorField.Text = "";
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
            label65.Hide(); label62.Hide();
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
            label67.Hide(); label66.Hide();

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
        private void dataMarital_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            IdField4.Text = dataMarital.Rows[e.RowIndex].Cells[0].Value.ToString();


            MaritalField.Text = dataMarital.Rows[e.RowIndex].Cells[1].Value.ToString();

        }
        public void CleandataMarital()
        {
            IdField4.Text = "";
            MaritalField.Text = "";
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
            label69.Hide();
            label68.Hide();

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
        private void dataType_RowHeaderMouseClick(object sender,
DataGridViewCellMouseEventArgs e)
        {
            textBox4.Text = dataType.Rows[e.RowIndex].Cells[0].Value.ToString();
            AccurField.Text = dataType.Rows[e.RowIndex].Cells[1].Value.ToString();
        }
        public void CleandataType()
        {
            textBox4.Text = "";
            AccurField.Text = "";
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
                    if ((a[i] >= 'a' && a[i] <= 'z')|| (a[i] >= 'а' && a[i] <= 'я')|| (a[i] >= 'А' && a[i] <= 'Я')|| (a[i] >= 'A' && a[i] <= 'Z'))
                    {
                        throw new Exception("Не допускаются буквенные символы");
                    }
                 
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
         string   a = numberField.Text;
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
            SpellChecking(DocField, "Фамилия доктора должна", "фамилии доктора", "фамилией доктора");
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
           string a =dateField2.Text;

            for (int i = 0; i < a.Length; i++)
            {
                if (!(Regex.IsMatch(a, @"^\d{4}-\d{2}-\d{2}$") || Regex.IsMatch(a, @"^\d{2}-\d{2}-\d{4}$") ||
                   Regex.IsMatch(a, @"^\d{4}.\d{2}.\d{2}$") || Regex.IsMatch(a, @"^\d{2}.\d{2}.\d{4}$")))
                {
                    throw new Exception("Введенная последовательность символов не является датой");
                }
                if (a[i] == ' ') { }

            }
            SpellChecking(typeField, "Наименование вида начисления должно", "наименовании вида начисления", "наименованием вида начисления");
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

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                this.Hide();
                var reports = new Reports();
              
                string[] s = File.ReadAllLines($"Report.txt");
                
                foreach (var item in s)
                {
                    reports.listBox1.Items.Add(item);
                }
                reports.ShowDialog();
            }
            catch(Exception ex ) { MessageBox.Show(ex.Message, ""); }
        }

        private void newData_Click(object sender, EventArgs e)
        {
            worker.ShowWorkers(SortBox1, dataWrk);
        }

        private void newData1_Click(object sender, EventArgs e)
        {
            post.ShowPost(SortBox2, dataPost);
        }

        private void newData2_Click(object sender, EventArgs e)
        {
            inc.ShowIncome(SortBox3, dataInc);

        }

        private void newData3_Click(object sender, EventArgs e)
        {
            health.ShowHealth(SortBox4, dataHealth);

        }

        private void newData4_Click(object sender, EventArgs e)
        {
            status.ShowStatus(sorrt1, dataMarital);

        }

        private void newData5_Click(object sender, EventArgs e)
        {
            accur.ShowAccurals(SortBox5, dataAccurals);


        }

        private void newData6_Click(object sender, EventArgs e)
        {
            type.ShowTypes(sort, dataType);
        }

        public void SpellChecking(System.Windows.Forms.TextBox b, string form1, string form2, string form3)
        {
            string a = b.Text;
            
            for (int i = 0; i < a.Length; i++)
            {
                if (a[0] == a.ToLower()[0]) { throw new Exception($"{form1} начинаться с прописной буквы!"); }
                if (a[0] >= 'A' && a[0] <= 'Z')
                {

                    throw new Exception("Допускаются только символы русского алфавита");

                }
                if (a[0] >= 'А' && a[0] <= 'Я')
                {
                    if (a[i] >= 'a' && a[i] <= 'z')
                    {
                        throw new Exception("Допускаются только символы русского алфавита");
                    }
                    if (a[i] >= 'А' && a[i] <= 'Я' && i != 0)
                    {
                        throw new Exception($"Использование разных регистров в {form2}");
                    }
                }
                if (a[i] >= '0' && a[i] <= '9')
                {

                    throw new Exception($"Введенная последовательность символов не является {form3}");
                }
                if (a[i] == '@' || a[i] == '$' || a[i] == '!' || a[i] == '?' || a[i] == '#'
                    || a[i] == ';' || a[i] == ':' || a[i] == ',' || a[i] == '.' || a[i] == '/'
                    || a[i] == '|' || a[i] == '*' || a[i] == '&' || a[i] == '"' || a[i] == '`' || a[i] == '~' || a[i] == '^'
                    || a[i] == '>' || a[i] == '<')
                {
                    throw new Exception($"Введенная последовательность символов не является {form3}");
                }
                if (a[i] == ' ') { }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Passwords form = new Passwords();
            form.ShowDialog();
       
        }

        private void button7_Click(object sender, EventArgs e)
        {

            report.ExportWord(dataGridView2);
        }

        private void button6_Click(object sender, EventArgs e)
        {

            report.ExportExcel(dataGridView2);
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            report.ReportHealth(textBox3.Text, dataGridView2);
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String[] a = comboBox1.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            report.Report(a[0], a[1], a[2], dataGridView1);
        }

        private void button8_Click(object sender, EventArgs e)
        {

            try
            {
                if (textBox1.Text != "")
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    string s = this.textBox1.Text;
                    connection.Open();
                    adapter = new SqlDataAdapter(s, connection);
                    adapter.Fill(dt);
                    dataGridView2.DataSource = dt;
                    connection.Close();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, ""); }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            report.ReportCity(textBox2.Text,textBox6.Text, dataGridView2);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            report.ReportHealth(textBox3.Text, dataGridView2);
        }

   
        private void button12_Click(object sender, EventArgs e)
        {
            report.ReportPost(comboBox2.Text, dataGridView2);
        }


        private void button13_Click(object sender, EventArgs e)
        {
            using (var color_dialog = new ColorDialog())
            {
                if (color_dialog.ShowDialog() != DialogResult.OK) return;
                (sender as System.Windows.Forms.Button).BackColor = color_dialog.Color;
            }
            switch(tabControl1.SelectedIndex)
            {
                case 0: { dataWrk.BackgroundColor = this.button13.BackColor; break; }
                case 1: { dataPost.BackgroundColor = this.button13.BackColor; break; }
                case 2: { dataInc.BackgroundColor = this.button13.BackColor; break; }
                case 3: { dataHealth.BackgroundColor = this.button13.BackColor; break; }
                case 4: { dataMarital.BackgroundColor = this.button13.BackColor; break; }
                case 5: { dataAccurals.BackgroundColor = this.button13.BackColor; break; }
                case 6: { dataType.BackgroundColor = this.button13.BackColor; break; }
                case 7: { dataGridView1.BackgroundColor = this.button13.BackColor; break; }
                case 8: { dataGridView2.BackgroundColor = this.button13.BackColor; break; }
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            using (var color_dialog = new ColorDialog())
            {
                if (color_dialog.ShowDialog() != DialogResult.OK) return;
                (sender as System.Windows.Forms.Button).BackColor = color_dialog.Color;
            }
            switch (tabControl1.SelectedIndex)
            {
                case 0: { dataWrk.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 1: { dataPost.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 2: { dataInc.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 3: { dataHealth.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 4: { dataMarital.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 5: { dataAccurals.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 6: { dataType.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 7: { dataGridView1.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
                case 8: { dataGridView2.DefaultCellStyle.SelectionBackColor = this.button11.BackColor; break; }
            }

  
           

        }

        private void button14_Click(object sender, EventArgs e)
        {
            using (var color_dialog = new ColorDialog())
            {
                if (color_dialog.ShowDialog() != DialogResult.OK) return;
                (sender as System.Windows.Forms.Button).BackColor = color_dialog.Color;
            }
            switch (tabControl1.SelectedIndex)
            {
                case 0: { dataWrk.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 1: { dataPost.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 2: { dataInc.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 3: { dataHealth.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 4: { dataMarital.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 5: { dataAccurals.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 6: { dataType.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 7: { dataGridView1.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
                case 8: { dataGridView2.DefaultCellStyle.SelectionForeColor = this.button14.BackColor; break; }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            using (var color_dialog = new ColorDialog())
            {
                if (color_dialog.ShowDialog() != DialogResult.OK) return;
                (sender as System.Windows.Forms.Button).BackColor = color_dialog.Color;
            }
            switch (tabControl1.SelectedIndex)
            {
                case 0: { dataWrk.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 1: { dataPost.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 2: { dataInc.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 3: { dataHealth.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 4: { dataMarital.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 5: { dataAccurals.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 6: { dataType.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 7: { dataGridView1.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
                case 8: { dataGridView2.RowHeadersDefaultCellStyle.SelectionBackColor = this.button15.BackColor; break; }
            }
        }

      
     
     
    }
}