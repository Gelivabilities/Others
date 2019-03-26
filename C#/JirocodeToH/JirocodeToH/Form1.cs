using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JirocodeToH
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = textBox1.Text;
            double currentNoteTime = -Convert.ToDouble(textBox4.Text);
            for (double i = 1; i <= s.IndexOf(','); i++)
            {
                textBox2.Text += "第" + i + "个音符时间为" + currentNoteTime + "\r\n";
                currentNoteTime+=Convert.ToDouble(textBox3.Text) / 240 / s.IndexOf(',');
            }

                textBox2.Text += headToCode(false, 20, 8, -40, 0, 0, 300, 150, 300, 600, 300, 600, 0, 0, 100, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
        }
        string headToCode
            (
                bool fenqi,
                int liangHun,int keHun,int bukeHun,
                int xuanren,int daren,
                int liang,int ke,int lianda,int teliang,int teke,int dalianda,int ballon,int digua,
                int sectionNum,float bpm,float offset
            )
        {
            string code="";
            //分歧
            if (fenqi) code += "01 ";
            else code += "00 ";
            //无关代码
            code+="01 39 10 27 00 40 1F 00 00 ";
            //1良所给魂槽量
            code+=liangHun.ToString("x8")+" 00 00 00 ";
            //1可所给魂槽量
            code += keHun.ToString("x8") + " 00 00 00 ";
            //1不可所给魂槽量
            code += bukeHun.ToString("x8") + " FF FF FF 00 00 01 00 ";
            //玄人谱面所给魂槽量相对加成
            code += "00 00 "+xuanren.ToString("x8")+" 00 ";
            //达人谱面所给魂槽量相对加成
            code += "00 00 " + daren.ToString("x8") + " 00 ";
            //分歧点数判定
            code += liang.ToString("x8") + " 00 00 00 " + ke.ToString("x8") + " 00 00 00 00 00 00 00 "
                  + teliang.ToString("x8") + " 00 00 00 " + teke.ToString("x8") + " 00 00 00 "
                  + lianda.ToString("x8") + " 00 00 00 " + ballon.ToString("x8") + " 00 00 00 "
                  + digua.ToString("x8") + " 00 00 00 ";
            //无关代码
            code+="00 00 00 00 E8 86 12 00 ";
            //小节数量
            code+=sectionNum.ToString("x8")+" 00 00 00 56 37 40 00 ";
            //bpm
            if (bpm >= 256) code += "00 01 ";//bpm超过255要进一位
            else code += "00 00 ";


            return code;
        }
        string sectionToCode(int spectrum,int nodeNum){
            string code="";

            return code;
        }
        string nodeToCode(int type,float time) { 
            string code="";

            return code;
        }
    }
}
