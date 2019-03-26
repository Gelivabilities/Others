using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void executebutton1event() 
        {
            try
            {
                if (button1.Text == "开始" || button1.Text == "继续")
                {
                    totalstoptimespan += stoptimespan;
                    button1.Text = "暂停";
                    timeflag = true;
                    button2.Enabled = true;
                }
                else
                    if (button1.Text == "暂停")
                    {
                        stoptime = DateTime.Now;
                        button1.Text = "继续";
                        timeflag = false;
                    }
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            executebutton1event();
        }
        static string timestring = "2:30:00.000";
        DateTime initialrundt;//程序运行时间
        DateTime stoptime;//当前暂停计时的时间
        TimeSpan stoptimespan = TimeSpan.Parse("0");//本次暂停了多长时间
        TimeSpan totalstoptimespan = TimeSpan.Parse("0");//一共暂停了多长时间
        TimeSpan timelimit = TimeSpan.Parse(timestring);
        bool timeflag = false;//切换暂停和计时状态
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (timeflag)
                {
                    TimeSpan t = timelimit - (DateTime.Now - initialrundt - totalstoptimespan);
                    string ms = (t.Milliseconds!=0)?((t.Milliseconds >= 100) ? (t.Milliseconds.ToString()) : ((t.Milliseconds >= 10) ? ("0" + t.Milliseconds) : ("00" + t.Milliseconds))):"000";
                    string s = t.Seconds >= 10 ? t.Seconds.ToString() : "0" + t.Seconds.ToString();
                    string min = t.Minutes >= 10 ? t.Minutes.ToString() : "0" + t.Minutes.ToString();
                    label1.Text = t.Hours+":"+min+":"+s+"."+ms ;
                    if (t < TimeSpan.Parse("0"))
                    {
                        label1.Text = "时间到";
                        t = timelimit;
                        totalstoptimespan =TimeSpan.Parse("0");
                        stoptimespan = TimeSpan.Parse("0");
                        initialrundt=stoptime = DateTime.Now;
                        button1.Text = "开始";
                        timeflag = false;
                    }
                }
                else
                {
                    stoptimespan = DateTime.Now - stoptime;
                }
            }
            catch { timer1.Enabled = false; }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                initialrundt = DateTime.Now;
                stoptime = DateTime.Now;
                label1.Text = timestring;
            }
            catch { }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Text="不让关";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timeflag = false;
            label1.Text = timestring;
            totalstoptimespan = TimeSpan.Parse("0");
            stoptimespan = TimeSpan.Parse("0");
            initialrundt = stoptime = DateTime.Now;
            button1.Text = "开始";
            button2.Enabled = false;
        }
    }
}
