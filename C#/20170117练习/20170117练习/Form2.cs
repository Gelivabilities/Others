using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _20170117练习
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 f1;
            f1 = (Form1)this.Owner;
            f1.lianglv=(double)numericUpDown1.Value /100;
            f1.tongjirefresh();
            f1.Enabled = true;
            this.Dispose();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1 = (Form1)this.Owner;
            numericUpDown1.Value=(int)(100*f1.lianglv);
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Owner.Enabled = true;
        }
    }
}
