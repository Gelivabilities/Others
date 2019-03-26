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

namespace _2011机试题目2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void executecommand(string s,int load)
        {
            try
            {
                string tablename = "";
                switch (load)
                {
                    case 0: tablename = "buses"; break;
                    case 1: tablename = "driver"; break;
                    case 2: tablename = "buslines"; break;
                    case 3: tablename = "timetable"; break;
                    default: break;
                }
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(s, con);
                DataSet ds = new DataSet();
                da.Fill(ds, tablename);
                switch (load)
                {
                    case 0: dataGridView1.DataSource = ds.Tables[tablename]; break;
                    case 1: dataGridView2.DataSource = ds.Tables[tablename]; break;
                    case 2: dataGridView3.DataSource = ds.Tables[tablename]; break;
                    case 3: dataGridView4.DataSource = ds.Tables[tablename]; break;
                    default: break;
                }
                con.Close();
            }
            catch(Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            executecommand("select carnum as '车牌号',producer as '生产厂商' from buses", 0);
            executecommand("select idnum as '员工号',name as '姓名',gender as '性别',age as '年龄',tel as '电话号码',psw as '密码' from driver", 1);
            executecommand("select linenum as '线路编号',start as '始发站',terminal as '终点站',distance as '距离' from buslines", 2);
            executecommand("select linenum as '线路编号',idnum as '员工号',carnum as '车牌号',starttime as '发车时间' from timetable", 3);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from driver where idnum='" + textBox3.Text +"'";
                cmd.ExecuteNonQuery();
                executecommand("select idnum as '员工号',name as '姓名',gender as '性别',age as '年龄',tel as '电话号码',psw as '密码' from driver", 1);
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox7.Text = textBox8.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.SelectedCells[0].Value.ToString().Trim();
            textBox2.Text = dataGridView1.SelectedCells[1].Value.ToString().Trim();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView2.SelectedCells[0].Value.ToString().Trim();
            textBox4.Text = dataGridView2.SelectedCells[1].Value.ToString().Trim();
            textBox5.Text = dataGridView2.SelectedCells[2].Value.ToString().Trim();
            textBox6.Text = dataGridView2.SelectedCells[3].Value.ToString().Trim();
            textBox7.Text = dataGridView2.SelectedCells[4].Value.ToString().Trim();
            textBox8.Text = dataGridView2.SelectedCells[5].Value.ToString().Trim();
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox9.Text = dataGridView3.SelectedCells[0].Value.ToString().Trim();
            textBox10.Text = dataGridView3.SelectedCells[1].Value.ToString().Trim();
            textBox11.Text = dataGridView3.SelectedCells[2].Value.ToString().Trim();
            textBox12.Text = dataGridView3.SelectedCells[3].Value.ToString().Trim();

        }

        private void button16_Click(object sender, EventArgs e)
        {
            textBox13.Text = textBox14.Text = textBox15.Text = textBox16.Text = "";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            textBox9.Text = textBox10.Text = textBox11.Text = textBox12.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox2.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into driver values('" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "'," + textBox6.Text + "," + textBox7.Text + ",'"  + textBox8.Text + "')";
                cmd.ExecuteNonQuery();
                executecommand("select idnum as '员工号',name as '姓名',gender as '性别',age as '年龄',tel as '电话号码',psw as '密码' from driver", 1);
                con.Close();
            }
            catch(Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into buses values('" + textBox1.Text + "','" + textBox2.Text + "')" ;
                cmd.ExecuteNonQuery();
                executecommand("select carnum as '车牌号',producer as '生产厂商' from buses", 0);
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into buslines values(" + textBox9.Text + ",'" + textBox10.Text+"','"+textBox11.Text+"',"+textBox12.Text + ")";
                cmd.ExecuteNonQuery();
                executecommand("select linenum as '线路编号',start as '始发站',terminal as '终点站',distance as '距离' from buslines", 2);
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into timetable values(" + textBox13.Text + ",'" + textBox14.Text + "','" + textBox15.Text + "','" + textBox16.Text + "')";
                cmd.ExecuteNonQuery();
                executecommand("select linenum as '线路编号',idnum as '员工号',carnum as '车牌号',starttime as '发车时间' from timetable", 3);
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                string sql = "select * from buslines";
                string where = " where ";
                bool flag = false;//where处是否有两个条件，有就要加and
                if (textBox17.Text != "")
                {
                    where+="linenum='"+textBox17.Text+"'";
                    flag = true;
                }
                if (textBox18.Text!="") 
                {
                    if (flag) where += " and ";
                    where += "start='"+textBox18.Text+"'";
                }
                sql += where;
                SqlDataAdapter da = new SqlDataAdapter(sql, con);
                DataSet ds = new DataSet();
                da.Fill(ds, "buslines");
                dataGridView5.DataSource = ds.Tables["buslines"];
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=gjcgl;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from buses where carnum='" + textBox1.Text +"'";
                cmd.ExecuteNonQuery();
                executecommand("select carnum as '车牌号',producer as '生产厂商' from buses", 0);
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    }
}
