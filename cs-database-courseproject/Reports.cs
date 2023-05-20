using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace cs_database_courseproject.users
{
    public partial class Reports : Form
    {
        SystemAdministrator administrator= new SystemAdministrator();
        public Reports()
        {
            InitializeComponent();
            button2.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
            var wel = new AdminForm();
            this.Hide();
            wel.Closed += (s, args) => this.Close();
            wel.ShowDialog();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(administrator.ReportAnswer(),"Состояние");
      
            System.Collections.Generic.List<string> linesList = File.ReadAllLines("Report.txt").ToList();
            linesList.Remove(listBox1.Items[listBox1.SelectedIndex].ToString());
            File.WriteAllLines("Report.txt", linesList.ToArray());
            listBox1.Items.Remove(listBox1.Items[listBox1.SelectedIndex]);
            button2.Hide();
        }
    }
}
