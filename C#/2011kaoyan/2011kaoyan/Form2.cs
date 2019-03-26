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

namespace _2011kaoyan
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into buses values(\'" + textBox1.Text + "\','" + textBox2.Text + "\')";
                cmd.ExecuteNonQuery();
                MessageBox.Show("添加成功");
            }
            catch { MessageBox.Show("添加失败"); }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into buslines values(" + textBox9.Text + ",'" + textBox10.Text + "\','" + textBox11.Text + "\'," + textBox12.Text + ")";
                cmd.ExecuteNonQuery();
                MessageBox.Show("添加成功");
            }
            catch 
            {
                MessageBox.Show("添加失败");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into driver values('" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "'," + textBox6.Text +","+textBox7.Text+",'"+textBox8.Text+ "')";
                cmd.ExecuteNonQuery();
                MessageBox.Show("添加成功");
            }
            catch
            {
                MessageBox.Show("添加失败");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "insert into timetable values(" + textBox13.Text + ",'" + textBox14.Text + "','" + textBox15.Text + "','" + textBox16.Text + "')";
                cmd.ExecuteNonQuery();
                MessageBox.Show("添加成功");
            }
            catch(Exception ex)
            {
                MessageBox.Show("添加失败"+ex);
            }
        }
    }
}
