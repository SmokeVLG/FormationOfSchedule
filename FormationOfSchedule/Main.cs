using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace FormationOfSchedule
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }


        private void Main_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tb_login.Text != "" )
            {
                Form1 frm = new Form1(tb_login.Text, tb_password.Text, this);
                frm.Show();
                this.Hide();
            }

            else
            {
                MessageBox.Show("Введите все учетные данные!");
            }
            
        }




    }
}
