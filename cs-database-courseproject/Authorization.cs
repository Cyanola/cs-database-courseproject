using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace cs_database_courseproject
{
    public partial class Authorization : Form
    {
        Client client = new Client();
        SystemAdministrator administrator= new SystemAdministrator();
        Accountant accountant = new Accountant();
        bool Check = false;
        public Authorization()
        {
            InitializeComponent();
            passwordField.Font = new Font("Comic Sans MS", 12);
        
        }
        private void Log_in_Click(object sender, EventArgs e)
        {
            if(passwordField.Text == administrator.getPassword())
            {
                this.Hide();
                AdminForm form = new AdminForm();
                    form.Closed += (s, args) => this.Close();
                form.ShowDialog();
            }
          else  if(passwordField.Text == accountant.getPassword())
                {
                this.Hide();
                AccountantForm form = new AccountantForm();
                form.Closed += (s, args) => this.Close();
                form.ShowDialog();
            }
         else   if (passwordField.Text == client.getPassword())
            {
                this.Hide();
               ClientForm form = new ClientForm();
                form.Closed += (s, args) => this.Close();
                form.ShowDialog();
            }
            else MessageBox.Show("Неверный пароль");
          
        }

        private void Back_Click(object sender, EventArgs e)
        {
            this.Close();
            var wel = new Welcome();
            this.Hide();
            wel.Closed += (s, args) => this.Close(); 
            wel.ShowDialog();
       

        }
        private void Show_Click(object sender, EventArgs e)
        {
            if (Check)
            {
                passwordField.PasswordChar = '\0';
                Show.Image = Image.FromFile(@"view_see_hide_eye_close_search_look_icon_232697 (4).png");
                Check = false;
            }
           else if (!Check)
            {
                passwordField.PasswordChar = '*';
                Show.Image = Image.FromFile(@"kisspng-computer-icons-symbol-eye-eye-5ac71bffab9197.2674698215229982717028 (1) (1).png");
                Check = true;
            }
        }
    }
}
