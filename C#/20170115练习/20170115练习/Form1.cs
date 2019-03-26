using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _20170115练习
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            switchbutton(true,false,false);
            appendjudgement();
        }

        private void switchbutton(bool r1,bool r2, bool r3)
        {
            //第一个组的按钮
            textBox1.Enabled = r1;
            textBox2.Enabled = r1;
            textBox3.Enabled = r1;
            textBox4.Enabled = r1;
            textBox5.Enabled = r1;
            textBox6.Enabled = r1;
            textBox7.Enabled = r1;
            button1.Enabled = r1; 
            //第二个组的按钮
            comboBox1.Enabled = r2;
            comboBox2.Enabled = r2;
            textBox8.Enabled = r2;
            button2.Enabled = r2;
            //第三个组的按钮
            comboBox3.Enabled = r3;
            button3.Enabled = r3;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            switchbutton(false,true,false);
            amendjudgement();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            switchbutton(false,false,true);
            if (comboBox3.Text == "") button3.Enabled = false;
            else button3.Enabled = true;
        }

        private void renewdata() 
        {
            try
            {
                string conStr = @"server=.;database=taikosongsinfo;integrated security=true";
                SqlConnection connection = new SqlConnection(conStr);//创建一个数据库连接
                connection.Open();//打开数据库连接
                SqlCommand cmd = connection.CreateCommand();
                cmd.CommandText = "select songid as 曲id,songname as 曲名,liang as 良,ke as 可,buke as 不可,score as 得点,lianda as 连打 from songsinfo";
                SqlDataAdapter ad = new SqlDataAdapter(cmd.CommandText, connection);
                DataSet ds = new DataSet();
                ad.Fill(ds, "songsinfo");
                dataGridView1.DataSource = ds.Tables["songsinfo"];
                connection.Close();

                SqlConnection con = new SqlConnection(@"server=.;database=taikosongsinfo;integrated security=true");
                con.Open();
                SqlCommand sc = con.CreateCommand();
                sc.CommandText = "select songname from songsinfo";
                SqlDataReader reader = sc.ExecuteReader();
                comboBox1.Items.Clear();
                comboBox3.Items.Clear();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader.GetString(0).Trim());
                    comboBox3.Items.Add(reader.GetString(0).Trim());
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库更新失败,原因:\n" + ex);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Show();
            //更新数据库读取部分
            renewdata();  
            //combobox2添加选项内容
            comboBox2.Items.Add("良");
            comboBox2.Items.Add("可");
            comboBox2.Items.Add("不可");
            comboBox2.Items.Add("得点");
            comboBox2.Items.Add("连打");
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con2 = new SqlConnection(@"server=.;database=taikosongsinfo;integrated security=true");
                con2.Open();
                SqlCommand cmd = con2.CreateCommand();
                cmd.CommandText = "insert into songsinfo values(" +"\'"+textBox1.Text+"\'" 
                    +','+"\'"+ textBox5.Text +"\'"+ "," + textBox6.Text + "," + textBox4.Text 
                    + "," + textBox3.Text + "," + textBox7.Text + "," + textBox2.Text+")";
                if (0 != cmd.ExecuteNonQuery()) 
                {
                    renewdata();
                    textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox7.Text = "";
                }
                else MessageBox.Show("添加失败");
                con2.Close();
            }
            catch
            {
                MessageBox.Show("添加失败，可能的原因如下：\n1、曲id重复\n2、得点不是10的整数倍\n3、曲名重复\n4、数据库连接失败\n");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s="";
            switch(comboBox2.Text)
            {
                case "良": s="liang"; break;
                case "可": s = "ke"; break;
                case "不可": s = "buke"; break;
                case "连打": s = "lianda"; break;
                case "得点": s = "score"; break;
                default: MessageBox.Show("请选择选项"); break;
            }
            try
            {
                SqlConnection con = new SqlConnection(@"server=.;database=taikosongsinfo;integrated security=true");
                SqlCommand cmd = con.CreateCommand();
                con.Open();
                cmd.CommandText="update songsinfo set "+s+"="+textBox8.Text+" where songname=\'"+comboBox1.Text+"\'";
                cmd.ExecuteNonQuery();
                con.Close();
                renewdata();
                button2.Enabled = false;
            }
            catch
            {
                MessageBox.Show("修改失败。可能的原因：\n1、数据库连接失败\n2、数值不合法");
            }
        }
        void amendjudgement() 
        {
            if (textBox8.Text == "" || comboBox2.Text == "" || comboBox1.Text == "") button2.Enabled = false;
            else button2.Enabled = true;
        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            amendjudgement();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            amendjudgement();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            amendjudgement();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text != "") button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from songsinfo where songname=\'" + comboBox3.Text + "\'";
                con.Open();
                cmd.ExecuteNonQuery();
                renewdata();
                button3.Enabled=false;
            }
            catch 
            {
                MessageBox.Show("删除失败");
            }
        }
        void appendjudgement() //添加时检查数据有没有全部填上
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == ""
                || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "") button1.Enabled = false;
            else button1.Enabled = true;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            appendjudgement();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void button4_Click(object sender, EventArgs e)
        {
            renewdata();
        }
    }
}