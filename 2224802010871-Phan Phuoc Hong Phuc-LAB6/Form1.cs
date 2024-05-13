using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _2224802010871_Phan_Phuoc_Hong_Phuc_LAB6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private Form currentFormChild;
        private void OpenChildForm(Form chilForm)
        {
            if(currentFormChild != null)
            {
                currentFormChild.Close();
            }
            currentFormChild = chilForm;
            chilForm.TopLevel = false;
            chilForm.FormBorderStyle = FormBorderStyle.None;
            chilForm.Dock = DockStyle.Fill;
            panel_body.Controls.Add(chilForm);
            panel_body.Tag = chilForm;
            chilForm.BringToFront();
            chilForm.Show();
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (currentFormChild != null)
            {
                currentFormChild.Close();
            }
            label1.Text = "Home";
        }

        private void btnA_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormA());
            label1.Text = btnA.Text;
        }

        private void btnB_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormB());
            label1.Text = btnB.Text;
        }

        private void btnC_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormC());
            label1.Text = btnC.Text;
        }
    }
}
