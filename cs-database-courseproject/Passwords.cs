using Microsoft.Office.Interop.Excel;
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
using System.Xml.Linq;

namespace cs_database_courseproject
{
    public partial class Passwords : Form
    {
        Client client = new Client();
        SystemAdministrator administrator = new SystemAdministrator();
        Accountant accountant = new Accountant();
        public string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public SqlDataAdapter adapter;
        public SqlCommand cmd;
        public SqlConnection connection = new
SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        public Passwords()
        {
            InitializeComponent();
            label1.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
           new AdminForm().Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string one = textBox1.Text;
            string two = textBox2.Text;
            string three = textBox3.Text;
            if (radioButton1.Checked)
            {
                if (one != "" && two != "" && three != "")
                {
                    if (one == administrator.getPassword() && textBox3.Text == textBox2.Text)
                    {
                        cmd = new SqlCommand($"UPDATE Users SET Users.Password = '{three}', ID_Role = 1 WHERE ID_User = 1",
                  connection);
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Пароль изменен");
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                    }
                }
            }
            else if (radioButton2.Checked)
            {
                if (one != "" && two != "" && three != "")
                {
                    if (one == accountant.getPassword() && textBox3.Text == textBox2.Text)
                    {
                        cmd = new SqlCommand($"UPDATE Users SET Users.Password = '{three}', ID_Role = 2 WHERE ID_User = 2",
                 connection);
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Пароль изменен");
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                    }
                }
            }
            else if (radioButton3.Checked)
            {
                if (one != "" && two != "" && three != "")
                {
                    if (one == client.getPassword() && textBox3.Text == textBox2.Text)
                    {
                        cmd = new SqlCommand($"UPDATE Users SET Users.Password = '{three}', ID_Role = 3 WHERE ID_User = 3",
                 connection);
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Пароль изменен");
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                    }
                }
            }
            else { MessageBox.Show("Не выбран пользователь", ""); }
      
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text != textBox2.Text)
            {
                label1.Show();
                label1.Text = "Пароли не совпадают";
            }
            else { label1.Hide(); label1.Text = " "; }
        }
    }
}
