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

namespace cq
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string[] s=new string[10000];

        int current = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Height = 177;
            if (button1.Text == "抽")
            {
                button1.Text = "停";
                timer1.Enabled = true;
            }
            else
            {
                button1.Text = "抽";
                timer1.Enabled = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var file = File.Open(Application.StartupPath +"\\chouqian.txt", FileMode.Open);

            int i = 0;
            using (var stream = new StreamReader(file))
            {
                while (!stream.EndOfStream)
                {
                    s[i]=stream.ReadLine();
                    i++;
                }
            }

            current = i;

            file.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Random r = new Random();
            label1.Text=s[r.Next(0,current)];
        }
    }
}
