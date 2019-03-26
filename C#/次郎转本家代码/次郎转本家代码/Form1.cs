using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 次郎转本家代码
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            analyze();
            toCode();
        }

        private void toCode() 
        {
                  
        }

        private void analyze() 
        {
            if (getLength(textBox1.Text) == 0) return;
            label3.Text = "此小节由" + getLength(textBox1.Text) + "分音符构成,音符个数为" + getNote(textBox1.Text);
            label3.Text += "。\n第1-第" + getNote(textBox1.Text) + "个音符相对小节播放时间（16进制）分别为：\n";
            string s;
            for (int i = 0; i < textBox1.Text.Length; i++)
            {
                s = textBox1.Text;
                if (s.Substring(i, 1) != "0")
                {
                    label3.Text += format((BitConverter.ToInt32(BitConverter.GetBytes(i  * 15 / float.Parse(textBox3.Text)), 0)).ToString("X")) + "  ";
                    //label3.Text += format((Convert.ToInt32((i - 1) * 15 / float.Parse(textBox3.Text)*1000)).ToString("X")) + "  ";
                }
            }
            //label3.Text += (BitConverter.ToInt32(BitConverter.GetBytes((11 - 1) * 15 / float.Parse(textBox3.Text)), 0)).ToString("X"); 
        }
        private string format(string s)
        {
            while (s.Length!=8) 
            {
                s = "0"+s;         
            }
            string t;
            t = s.Substring(0,2);


            s = s.Insert(6, " ");
            s = s.Insert(4, " ");
            s = s.Insert(2, " ");
            return s;
        }

        private int getLength(string s) 
        {
            return s.Length;   
        }
        private int getNote(string s) 
        {
            int l=0;
            for (int i = 0; i < s.Length; i++)
                if (s.Substring(i, 1) != "0") 
                    l++;
            return l;
        }
    }
}
