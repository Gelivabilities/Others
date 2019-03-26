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

namespace _20170117练习
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void updatedatagridview() 
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");//usr=xx;psw=xx;
                con.Open();
                string s = "select songid as 曲id,songname as 曲名,liang as 良,ke as 可,buke as 不可,score as 得点,lianda as 连打 from songsinfo";
                SqlDataAdapter ad = new SqlDataAdapter(s, con);
                DataSet ds = new DataSet();
                ad.Fill(ds, "songsinfo");
                dataGridView1.DataSource = ds.Tables["songsinfo"];
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public double lianglv = 0.9;

        public void tongjirefresh() 
        {
            dataGridView2.DataSource = null;
            bool flag = false;
            string s="select ";
            if (checkBox1.Checked) 
            { 
                s += "sum(case when buke=0 then 1 else 0 end) as 全连";
                flag = true;
            }
            if (checkBox2.Checked) 
            {
                if(flag)s+=",";
                s+=" sum(case when ke<10 and buke=0 then 1 else 0 end) as 全连单可";
                flag = true;
            }
            if (checkBox3.Checked)
            {
                if (flag) s += ",";
                s += " sum(case when ke=0 and buke=0 then 1 else 0 end) as 全良";
                flag = true;
            }
            if (checkBox4.Checked)
            {
                if (flag) s += ",";
                s += " sum(case when buke<10 then 1 else 0 end) as 单不可";
                flag = true;
            }
            if (checkBox5.Checked)
            {
                if (flag) s += ",";
                s += " sum(case when score>=1200000 then 1 else 0 end) as '120W以上'";
                flag = true;
            }
            if (checkBox6.Checked)
            {
                if (flag) s += ",";
                s += " sum(case when liang/(liang+ke+buke)>"+lianglv+" then 1 else 0 end) as '良率高("+((int)(100*lianglv)).ToString()+"%)'";
                flag=true;
            }
            s += " from songsinfo";
            if (flag)
            {
                try
                {
                    SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");
                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(s, con);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "songsinfo");
                    dataGridView2.DataSource = ds.Tables["songsinfo"];
                }
                catch (Exception ex) 
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else dataGridView2.DataSource = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            updatedatagridview();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox7.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.SelectedCells[0].Value.ToString().Trim();
            textBox2.Text = dataGridView1.SelectedCells[1].Value.ToString().Trim();
            textBox3.Text = dataGridView1.SelectedCells[2].Value.ToString();
            textBox4.Text = dataGridView1.SelectedCells[3].Value.ToString();
            textBox5.Text = dataGridView1.SelectedCells[4].Value.ToString();
            textBox6.Text = dataGridView1.SelectedCells[5].Value.ToString();
            textBox7.Text = dataGridView1.SelectedCells[6].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into songsinfo values('" +( textBox1.Text.Trim() == "" ? "0" : textBox1.Text.Trim()) + "','" + (textBox2.Text.Trim() == "" ? "0" : textBox2.Text.Trim()) + "'," +( textBox3.Text.Trim() == "" ? "0" : textBox3.Text.Trim() )+ "," + (textBox4.Text.Trim() == "" ? "0" : textBox4.Text.Trim() )+ "," + (textBox5.Text.Trim() == "" ? "0" : textBox5.Text.Trim()) + "," + (textBox6.Text.Trim() == "" ? "0" : textBox6.Text.Trim()) + "," + (textBox7.Text.Trim() == "" ? "0" : textBox7.Text.Trim()) + ")";
                cmd.ExecuteNonQuery();
                updatedatagridview();
                tongjirefresh();
            }
            catch 
            {
                MessageBox.Show("添加失败。原因可能是：\n1、曲名或曲id可能有重复\n2、数据库被移除或破坏");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");
                con.Open();
                bool flag = false;
                SqlCommand cmd = con.CreateCommand();
                string sql = "update songsinfo set ";
                if (textBox2.Text != "") 
                { 
                    sql += "songname='" + textBox2.Text.Trim() + "'";
                    flag = true;
                }
                if (textBox3.Text != "") 
                {
                    if (flag) sql += ",";
                    sql += "liang="+textBox3.Text.Trim();
                    flag = true;
                }
                if (textBox4.Text != "")
                {
                    if (flag) sql += ",";
                    sql += "ke=" + textBox4.Text;
                    flag = true;
                }
                if (textBox5.Text != "")
                {
                    if (flag) sql += ",";
                    sql += "buke=" + textBox5.Text;
                    flag = true;
                }
                if (textBox6.Text != "")
                {
                    if (flag) sql += ",";
                    sql += "score=" + textBox6.Text;
                    flag = true;
                }
                if (textBox7.Text != "")
                {
                    if (flag) sql += ",";
                    sql += "lianda=" + textBox7.Text;
                }
                sql += " where songid='" + textBox1.Text.Trim()+"'";
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                updatedatagridview();
                tongjirefresh();
            }
            catch 
            {
                MessageBox.Show("修改失败，可能的原因有：\n1、没有此id的歌曲\n2、成绩输入数值不是10的倍数\n3、数据库被删除");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            amendbuttonrefresh();
            button2.Enabled = button3.Enabled=(textBox1.Text != "") ;

        }
        void amendbuttonrefresh() 
        {
            button1.Enabled = (textBox1.Text != "" && textBox2.Text != "");
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            amendbuttonrefresh();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=taikosongsinfo;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from songsinfo where songid='"+textBox1.Text+"'";
                if (cmd.ExecuteNonQuery()==0)MessageBox.Show("删除失败，没有此id的歌曲") ;
                updatedatagridview();
                textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox7.Text = "";
                tongjirefresh();
            }
            catch
            {
                MessageBox.Show("删除失败，数据库可能已被删除或破坏");
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
            {
                e.Handled = true;
            } 
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
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

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                this.Height = 364;
                tabControl1.Height = 92;
            }
            if (tabControl1.SelectedIndex == 1) 
            {
                tabControl1.Height = 137;
                this.Height=405;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            tongjirefresh();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Owner = this;
            f2.Show();
            this.Enabled = false;
        }
    }
}
