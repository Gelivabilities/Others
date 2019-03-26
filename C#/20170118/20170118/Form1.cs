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

namespace _20170118
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void loaddatagridview() 
        {
            SqlConnection con = new SqlConnection("server=.;database=test20170118;integrated security=true");
            con.Open();
            string s = "select xy as 学院,bj as 班级,nj as 年级,xh as 学号,kcdm as 课程代码,kcmc as 课程名称,xf as 学分,xq as 学期 from cxdr order by xh";
            SqlDataAdapter ad = new SqlDataAdapter(s, con);
            DataSet ds = new DataSet();
            ad.Fill(ds, "cxdr");
            dataGridView1.DataSource = ds.Tables["cxdr"];
            label6.Text = "共"+dataGridView1.RowCount+"条结果";
            con.Close();
        }

        void loadcombobox() 
        {
            try
            {
                SqlConnection con = new SqlConnection("server=.;database=test20170118;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select xy as 学院, bj as 班级,xh as 学号 ,kcdm as 课程代码,kcmc as 课程名称 from cxdr order by xh";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read()) 
                {
                    string s0=dr.GetString(0).Trim();
                    if (comboBox1.Items.Cast<object>().All(x => x.ToString() != s0)) 
                        comboBox1.Items.Add(s0);
                    string s1 = dr.GetString(1).Trim();
                    if (comboBox2.Items.Cast<object>().All(x => x.ToString() != s1)) 
                        comboBox2.Items.Add(s1);
                    string s2 = dr.GetValue(2).ToString();
                    if (comboBox3.Items.Cast<object>().All(x => x.ToString() != s2)) 
                        comboBox3.Items.Add(s2);
                    string s3 = dr.GetValue(3).ToString();
                    if (comboBox4.Items.Cast<object>().All(x => x.ToString() != s3)) 
                        comboBox4.Items.Add(s3);
                    string s4 = dr.GetString(4).Trim();
                    if (comboBox5.Items.Cast<object>().All(x => x.ToString() != s4)) 
                        comboBox5.Items.Add(s4);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(""+ex);
            }
        }

        void refreshdata() 
        {
            try
            {
                string s = "select xy as 学院,bj as 班级,nj as 年级,xh as 学号,kcdm as 课程代码,kcmc as 课程名称,xf as 学分,xq as 学期 from cxdr";
                string t = " where ";
                string sql = "";
                bool tiaojianflag = false;
                bool andflag = false;
                if (comboBox1.Text != "")
                {
                    t += "xy like'%" + comboBox1.Text + "%'";
                    andflag = true;
                    tiaojianflag = true;
                }
                if (comboBox2.Text != "")
                {
                    if (andflag) t += " and ";
                    t += "bj like'%" + comboBox2.Text + "%'";
                    andflag = true;
                    tiaojianflag = true;
                }
                if (comboBox3.Text != "")
                {
                    if (andflag) t += " and ";
                    t += "convert(varchar(50),xh) like '%" + comboBox3.Text+"%'";
                    andflag = true;
                    tiaojianflag = true;
                }
                if (comboBox4.Text != "")
                {
                    if (andflag) t += " and ";
                    t += "kcdm like'%" + comboBox4.Text + "%'";
                    andflag = true;
                    tiaojianflag = true;
                }
                if (comboBox5.Text != "")
                {
                    if (andflag) t += " and ";
                    t += "kcmc like'%" + comboBox5.Text + "%'";
                    tiaojianflag = true;
                }
                if (tiaojianflag) sql = s + t;
                else sql = s;
                sql += " order by xh";
                SqlConnection con = new SqlConnection("server=.;database=test20170118;integrated security=true");
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(sql,con);
                DataSet ds = new DataSet();
                da.Fill(ds, "cxdr");
                dataGridView1.DataSource=ds.Tables["cxdr"];
                label6.Text = "共"+(dataGridView1.RowCount-1)+"条结果";
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Show();
            loaddatagridview();
            loadcombobox();
            this.Text = "重修信息查询统计系统";
            this.Enabled = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox5_TextChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            refreshdata();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
    }
}
