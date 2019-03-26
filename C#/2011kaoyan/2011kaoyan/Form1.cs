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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void fillbuses() 
        {
                SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = "select carnum as 车牌号,manufacture as 生产厂商 from buses";
                SqlDataAdapter ad = new SqlDataAdapter(cmd.CommandText, con);
                DataSet ds = new DataSet();
                ad.Fill(ds, "buses");
                dataGridView1.DataSource = ds.Tables["buses"];
                con.Close();
        }

        void filldrivers() 
        {
            SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = "select idnumber as 员工号,name as 姓名,gender as 性别,age as 年龄,tel as 电话,pswrd as 密码 from driver";
            SqlDataAdapter ad = new SqlDataAdapter(cmd.CommandText, con);
            DataSet ds = new DataSet();
            ad.Fill(ds, "drivers");
            dataGridView2.DataSource = ds.Tables["drivers"];
            con.Close();
        }

        void fillbuslines() 
        {
            SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = "select linenum as 线路编号,startpoint as 起点,terminal as 终点,distance as 距离 from buslines";
            SqlDataAdapter ad = new SqlDataAdapter(cmd.CommandText, con);
            DataSet ds = new DataSet();
            ad.Fill(ds, "buslines");
            dataGridView3.DataSource = ds.Tables["buslines"];
            con.Close();
        }

        void filltimetable() 
        {
            SqlConnection con = new SqlConnection("server=.;database=2011kaoyan;integrated security=true");
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = "select linenum as 线路编号,idnumber as 员工号,carnum as 车牌号,starttime as 发车时间 from timetable";
            SqlDataAdapter ad = new SqlDataAdapter(cmd.CommandText, con);
            DataSet ds = new DataSet();
            ad.Fill(ds, "timetable");
            dataGridView4.DataSource = ds.Tables["timetable"];
            con.Close();  
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                fillbuses();
                filldrivers();
                fillbuslines();
                filltimetable();
            }
            catch { MessageBox.Show("1"); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();
            f4.Show();
        }
    }
}
