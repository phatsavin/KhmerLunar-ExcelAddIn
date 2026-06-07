using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KhmerLunar_ExcelAddIn
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Done_Setting = textBox1.Text;
            Properties.Settings.Default.Save(); // This saves the change
            MessageBox.Show("Your customization has been saved to settings!");
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.Done_Setting;
        }
    }
}
