using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = textBox2.Text;
            string pattern=@"\d\d:\d\d.\d\d";
            MatchCollection a= Regex.Matches(s,pattern);
            foreach (Match match in a)
            {

                string temp=match.Value;
                int millisecond = int.Parse(temp.Substring(6, 2));
                int second=int.Parse(temp.Substring(3,2));
                int minute = int.Parse(temp.Substring(0, 2));
                if (second + 10 >= 60) minute++;
                second = (second + 10) % 60;
                s=s.Replace(temp,
                    (minute<10?"0"+minute:minute.ToString())+":"+
                    (second<10?"0"+second:second.ToString())+"."+
                    (millisecond<10?"0"+millisecond:millisecond.ToString()));
            }
            textBox2.Text = s;
        }
    }
}
