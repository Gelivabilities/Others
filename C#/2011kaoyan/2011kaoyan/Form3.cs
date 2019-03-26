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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        void combobox1add() 
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select carnum from buses";
                SqlDataReader read = cmd.ExecuteReader();
                while (read.Read())
                {
                    comboBox1.Items.Add(read.GetString(0).Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
            }
        }

        void combobox2add()
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select linenum from buslines";
                SqlDataReader read = cmd.ExecuteReader();
                while (read.Read())
                {
                    comboBox2.Items.Add(read.GetValue(0));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
            }
        }

        void combobox3add()
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select idnumber from driver";
                SqlDataReader read = cmd.ExecuteReader();
                while (read.Read())
                {
                    comboBox3.Items.Add(read.GetValue(0));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
            }
        }

        void combobox4add()
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select linenum from timetable";
                SqlDataReader read = cmd.ExecuteReader();
                while (read.Read())
                {
                    comboBox4.Items.Add(read.GetValue(0));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
            }
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            combobox1add();
            combobox2add();
            combobox3add();
            combobox4add();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from buses where carnum='"+comboBox1.Text+"'";
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("删除成功");
                comboBox1.Items.Clear();
                comboBox1.Text = "";
                combobox1add();
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除失败" + ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from buslines where linenum='" + comboBox2.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("删除成功");
                comboBox2.Items.Clear();
                comboBox2.Text = "";
                combobox2add();
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除失败" + ex);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from driver where idnumber='" + comboBox3.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("删除成功");
                comboBox3.Items.Clear();
                comboBox3.Text = "";
                combobox3add();
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除失败" + ex);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "delete from timetable where linenum='" + comboBox4.Text + "'";
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("删除成功");
                comboBox4.Items.Clear();
                comboBox4.Text = "";
                combobox4add();
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除失败" + ex);
            }
        }
    }
}
