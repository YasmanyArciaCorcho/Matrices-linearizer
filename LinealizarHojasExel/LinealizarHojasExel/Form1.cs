using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LinealizarHojasExel
{
    public partial class Form1 : Form
    {
        OpenFileDialog openFileDialog1;
        public Form1()
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Text Files (.xls*)|*.xls*|All Files (*.*)|*.*"; ;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                new SecondForm(this,openFileDialog1.FileName).Show();
                Enabled = false;
                Visible = false;
            }
        }
    }
}
