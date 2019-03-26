using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace liangyuanchao
{
    
    public partial class Form1 : Form
    {
        //用于移动鼠标位置
        [DllImport("user32.dll")]
        private static extern int SetCursorPos(int x, int y);

        //用于点击鼠标
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        const int MOUSEEVENTF_LEFTUP = 0x0004;

        //是否在回放
        bool replaying = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem mnu = (ToolStripMenuItem)sender;
            textBox1.Text = textBox1.Text + "["+mnu.Name+"]" + mnu.Text+ "\r\n";
            textBox1.SelectionStart = textBox1.Text.Length;
            textBox1.ScrollToCaret();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("空的不用保存");
                return;
            }
            try
            {
                System.IO.File.WriteAllText(System.Environment.CurrentDirectory + "\\script.txt",textBox1.Text);
                MessageBox.Show("已保存至"+ System.Environment.CurrentDirectory + "\\script.txt");
            }
            catch { MessageBox.Show("保存失败"); }
            
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 8) { e.Handled = false; return; };
            if (e.KeyChar=='。') e.KeyChar='.';
            //输入合法性（无法转换说明输入不合法）
            try { float.Parse(textBox2.Text+e.KeyChar+ (e.KeyChar == '.'? "0" : "")); }
            catch { e.Handled = true; }
        }

        private bool check_button3_enabled()
        {
            return textBox2.Text == "" || textBox2.Text == "." || float.Parse(textBox2.Text) == 0 || textBox1.Text == "" || replaying ? false : true;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            button3.Enabled = check_button3_enabled();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try {ReadTxtContent(System.Environment.CurrentDirectory + "\\script.txt"); }
            catch { MessageBox.Show("没有保存的脚本"); }
        }

        public void ReadTxtContent(string Path)
        {
            StreamReader sr = new StreamReader(Path, Encoding.GetEncoding("UTF-8"));
            string content;
            textBox1.Text = "";
            while ((content = sr.ReadLine()) != null)textBox1.Text += content.ToString() + "\r\n";
            sr.Close();
        }

        //菜单左上角起点和间隔
        int start_0_x = 36;
        int start_0_y = 42;
        int delta_0_x =44;
        int delta_0_y = 15;
        //下拉间隔
        int drop_down_delta_y =18;
        //与下一级菜单间隔
        int delta_1_x =140;

        private string string_cut(string str,string str1,string str2,bool last,int tail_startindex)
        {
            try
            {
                string str_head = str1;
                string str_tail = str2;
                int head = str.IndexOf(str_head);
                int tail = (last) ? str.LastIndexOf(str_tail) : str.IndexOf(str_tail, tail_startindex);
                return str.Substring(head + 1, tail - head - 1);
            }
            catch { return ""; }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            button3.Text = "正在回放";
            button3.Enabled = false;
            replaying = true;
            textBox2.Enabled = false;
            try
            {
                int rows = Count(textBox1.Text, "[");//脚本行数
                string temp = textBox1.Text;
                textBox1.Text = "";
                //记录脚本对应的控件，如果是文本框，还要记录textbox内容是什么
                string[] scripts = new string[rows];
                string[] content = new string[rows];
                bool[] step = new bool[rows];//True对应菜单，false对应文本框
                for (int i = 0; i < rows; i++)
                {
                    scripts[i] = string_cut(temp, "[", "]", false, 0);
                    content[i] = string_cut(temp, "]", i != (rows - 1) ? "[" : "\r\n", false, 1);
                    step[i] = scripts[i][0] == 'M';
                    temp = string_cut(temp, "]", "\r\n", true, 0) + "\r\n";
                    temp = "[" + string_cut(temp, "[", "\r\n", true, 0) + "\r\n";
                }
                //回放脚本
                for (int i = 0; i < rows; i++)
                {
                    if (step[i])//菜单
                    {
                        //窗口绝对位置
                        int start_x = this.Location.X;
                        int start_y = this.Location.Y;

                        //分解菜单
                        int level = Count(scripts[i], "_") + 1;//菜单级数
                        int[] menu_no = new int[level];
                        temp = scripts[i].Substring(7, scripts[i].Length - 7);
                        for (int j = 0; j < level; j++)
                        {
                            menu_no[j] = int.Parse(temp.Substring(1,1));
                            temp = temp.Substring(2,temp.Length-2);
                        }
                        //点击菜单
                        for (int j = 0; j < level; j++)
                        {
                            int x_0 = start_x + start_0_x + (menu_no[0] - 1) * delta_0_x;
                            int y_0 = start_y + start_0_y;
                            //计算菜单在屏幕中的位置，因为脚本中只有控件信息，所以点击位置要算出来
                            switch (j)
                            {
                                case 0: SetCursorPos(x_0, y_0); break;
                                case 1: SetCursorPos(x_0, y_0+delta_0_y+ (menu_no[1]) * drop_down_delta_y); break;
                                case 2: SetCursorPos(x_0+delta_1_x, y_0 + delta_0_y + (menu_no[1]) * drop_down_delta_y+(menu_no[2]-1)*drop_down_delta_y); break;
                                case 3: SetCursorPos(x_0 + 2*delta_1_x, y_0 + delta_0_y + (menu_no[1]) * drop_down_delta_y + (menu_no[2] - 1) * drop_down_delta_y+ (menu_no[3] - 1) * drop_down_delta_y); break;
                            }
                            if(j==level-1)Delay(Convert.ToInt32(1000 / float.Parse(textBox2.Text)));
                            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                            Delay(Convert.ToInt32(1000 / float.Parse(textBox2.Text)));
                        }
                    }
                    else//文本框
                    {
                        textBox.Text = content[i].Replace("\n", "").Replace("\r", "");
                        //延迟
                        Delay(Convert.ToInt32(1000 / float.Parse(textBox2.Text)));
                    }
                }
            }
            catch { MessageBox.Show("代码有问题"); }
            finally
            {
                button3.Text = "回放上面的脚本";
                button3.Enabled = true;
                replaying = false;
                textBox2.Enabled = true;
            }
        }

        //防假死延迟
        public static void Delay(int milliSecond)
        {
            try
            {
                int start = Environment.TickCount;
                while (Math.Abs(Environment.TickCount - start) < milliSecond)
                {
                    Application.DoEvents();
                }
            }
            catch { }
        }

        private int Count(string str, string constr)
        {
            return str.Length-str.Replace(constr, "").Length;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text +="[textBox]"+textBox.Text+"\r\n";
            textBox1.SelectionStart = textBox1.Text.Length;
            textBox1.ScrollToCaret(); 
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button3.Enabled = check_button3_enabled();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }
    }
}
