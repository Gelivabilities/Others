using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace iOSTaikoSrADl
{
    //定义每个分类的内容的结构体
    public struct Song
    {
        public string num,difficulty,combo,name,id,keywords,variety,url;
    }
    public struct Pack 
    {
        public string num,name, keywords, date,url;
    }
    public struct Ipa 
    {
        public string num,version,url;
    }
    public struct Others 
    {
        public string num,name,keywords,url;
    }
    public struct All 
    {
        public string type,date,url;
    }
    //主窗口
    public partial class Form1 : Form
    {
        Form2 downloadForm = new Form2();

        //public delegate void AddDownloadTask(string id,string name,string url,string savaAddress);

        //public event AddDownloadTask addTask;

        public Form1()
        {
            InitializeComponent();
        }

        //定义好相关数组
        Pack[] pack = new Pack[100];
        Song[] song = new Song[1000];
        Ipa[] ipa = new Ipa[10];
        Others[] others = new Others[100];

        private void Form1_Load(object sender, EventArgs e)
        {
            //按包下载
            string s = getCode("http://a754571662.lingd.cc/article-6455317-1.html");
            staPack(s, ref pack);
            //按曲下载
            s = getCode("http://a754571662.lingd.cc/article-6455249-1.html");
            staSong(s, ref song);
            //不同版本ipa
            s = getCode("http://a754571662.lingd.cc/article-6455295-1.html");
            staIpa(s, ref ipa);
            //其他下载
            s = getCode("http://a754571662.lingd.cc/article-6460250-1.html");
            staOthers(s, ref others);

            //检查本程序版本
            string version = "0";
            s = getCode("http://a754571662.lingd.cc/article-6460168-1.html");
            s = delStrLeft(s,"【");
            s = delStrRight(s,"】");
            if (version != s)
            {
                DialogResult update = MessageBox.Show("发现新版本下载器，是否立即更新？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (DialogResult.Yes == update)
                {
                    s = getCode("http://a754571662.lingd.cc/article-6463855-1.html");
                    s = delStrLeft(s, "【");
                    s = delStrRight(s, "】");
                    System.Diagnostics.Process.Start(s);
                }
            }
            //默认添加按包下载的列表
            buildPack();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true) buildPack();
            if (radioButton1.Checked == true) buildSong();
            if (radioButton3.Checked == true) buildIpa();
            if (radioButton4.Checked == true) buildOthers();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton1.Checked==true)buildSong();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true) buildPack();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton3.Checked == true)buildIpa();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true) buildOthers();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("此功能未推出，敬请期待！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string s = getCode("http://a754571662.lingd.cc/article-6455293-1.html");
            All all = new All();
            staAll(ref s, ref all);
            if (all.type != "链接失效")
                System.Diagnostics.Process.Start("http://pan.baidu.com/s/1pJjyOhP#path=%252F");
            else
                MessageBox.Show("此链接已失效！请等待更新");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string svAdr ="";
            if(radioButton5.Checked==true)svAdr=System.IO.Directory.GetCurrentDirectory();
            if(radioButton6.Checked==true && textBox2.Text!="")svAdr=textBox2.Text;
            string url = "";
            int fileVariety;
            if (radioButton1.Checked == true) fileVariety = 1;
            else if (radioButton2.Checked == true) fileVariety = 2;
            else if (radioButton3.Checked == true) fileVariety = 3;
            else fileVariety = 4;
            int j;

            this.button5_Click(sender, e);

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].Selected == true)
                {
                    switch (fileVariety)
                    {
                        case 1:
                            for (j = 0; song[j].name != ""; j++)
                            {

                                if (listView1.Items[i].SubItems[0].Text == song[j].name)
                                {
                                    url = song[j].url;
                                }
                            }
                            break;
                        case 2:
                            for (j = 0; pack[j].name != null; j++)
                            {

                                if (listView1.Items[i].SubItems[0].Text == pack[j].name)
                                {
                                    url = pack[j].url;
                                }
                            }
                            break;
                        case 3:
                            for (j = 0; ipa[j].version != null; j++)
                            {

                                if (listView1.Items[i].SubItems[0].Text == ipa[j].version)
                                {
                                    url = pack[j].url;
                                }
                            }
                            break;
                        default:
                            for (j = 0; others[j].name != null; j++)
                            {

                                if (listView1.Items[i].SubItems[0].Text == others[j].name)
                                {
                                    url = pack[j].url;
                                }
                            }
                            break;
                    }
                    if (url != "")
                        downloadForm.AddBatchDownload
                        (
                            listView1.Items[i].SubItems[0].Text,
                            listView1.Items[i].SubItems[0].Text,
                            url,
                            svAdr
                        );
                }
                downloadForm.startDownload_Form1();
                button5.Enabled = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (Form2.form2Visible == false)
                downloadForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string svAdr = "";
            if (radioButton5.Checked == true) svAdr = System.IO.Directory.GetCurrentDirectory();
            if (radioButton6.Checked == true && textBox2.Text != "") svAdr = textBox2.Text;
            string url = "";
            int fileVariety;
            if (radioButton1.Checked == true) fileVariety = 1;
            else if (radioButton2.Checked == true) fileVariety = 2;
            else if (radioButton3.Checked == true) fileVariety = 3;
            else fileVariety = 4;
            int j;

            this.button5_Click(sender, e);

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                switch (fileVariety)
                {
                    case 1:
                        for (j = 0; song[j].name != ""; j++)
                        {

                            if (listView1.Items[i].SubItems[0].Text == song[j].name)
                            {
                                url = song[j].url;
                            }
                        }
                        break;
                    case 2:
                        for (j = 0; pack[j].name != null; j++)
                        {

                            if (listView1.Items[i].SubItems[0].Text == pack[j].name)
                            {
                                url = pack[j].url;
                            }
                        }
                        break;
                    case 3:
                        for (j = 0; ipa[j].version != null; j++)
                        {

                            if (listView1.Items[i].SubItems[0].Text == ipa[j].version)
                            {
                                url = pack[j].url;
                            }
                        }
                        break;
                    default:
                        for (j = 0; others[j].name != null; j++)
                        {

                            if (listView1.Items[i].SubItems[0].Text == others[j].name)
                            {
                                url = pack[j].url;
                            }
                        }
                        break;
                }
                if (url != "")
                    downloadForm.AddBatchDownload
                    (
                        listView1.Items[i].SubItems[0].Text,
                        listView1.Items[i].SubItems[0].Text,
                        url,
                        svAdr
                    );
            }
            downloadForm.startDownload_Form1();
            button5.Enabled = true;
        }
        //获源代码
        private string getCode(string url)
        {
            string strHTML = "";
            WebClient myWebClient = new WebClient();
            Stream myStream = myWebClient.OpenRead(url);
            StreamReader sr = new StreamReader(myStream, System.Text.Encoding.GetEncoding("utf-8"));
            strHTML = sr.ReadToEnd();
            myStream.Close();
            strHTML=strHTML.Replace("&nbsp;"," ");
            return strHTML;
        }
        //数组变列表
            private void buildPack()
            {
                int i;
                listView1.Items.Clear();
                for (i = listView1.Columns.Count - 1; i >= 0; i--)
                    listView1.Columns.RemoveAt(i);

                listView1.Columns.Add("名称");
                listView1.Columns.Add("关键词");
                listView1.Columns.Add("配信日期");
                listView1.Columns[0].Width = 200;
                listView1.Columns[1].Width = 200;
                listView1.Columns[2].Width = 150;
                i = 0;
                while (pack[i].name != null)
                {
                    if ((pack[i].name.ToLower().IndexOf(textBox1.Text) != -1 || pack[i].keywords.ToLower().IndexOf(textBox1.Text) != -1))
                    {
                        ListViewItem item = new ListViewItem();
                        item.SubItems[0].Text = pack[i].name;
                        item.SubItems.Add(pack[i].keywords);
                        item.SubItems.Add(pack[i].date);
                        listView1.Items.Add(item);
                    }
                    i++;
                }
            }
            private void buildSong()
            {
                int i;
                listView1.Items.Clear();
                for (i = listView1.Columns.Count - 1; i >= 0; i--)
                    listView1.Columns.RemoveAt(i);

                listView1.Columns.Add("名称");
                listView1.Columns.Add("关键词");
                listView1.Columns.Add("分类");
                listView1.Columns.Add("难度");
                listView1.Columns.Add("连击");
                listView1.Columns.Add("ID");
                listView1.Columns[0].Width = 150;
                listView1.Columns[1].Width = 150;
                listView1.Columns[2].Width = 75;
                listView1.Columns[3].Width = 50;
                listView1.Columns[4].Width = 50;
                listView1.Columns[5].Width = 75;
                i = 0;
                while (song[i].name.Length != 0)
                {
                    if ((song[i].name.ToLower().IndexOf(textBox1.Text) != -1 || song[i].keywords.ToLower().IndexOf(textBox1.Text) != -1)||song[i].variety.ToLower().IndexOf(textBox1.Text) != -1)
                    {
                        ListViewItem item = new ListViewItem();
                        item.SubItems[0].Text = song[i].name;
                        item.SubItems.Add(song[i].keywords);
                        item.SubItems.Add(song[i].variety);
                        item.SubItems.Add(song[i].difficulty);
                        item.SubItems.Add(song[i].combo);
                        item.SubItems.Add(song[i].id);
                        listView1.Items.Add(item);
                    }
                    i++;
                }
            }
            public void buildIpa()
            {
                int i;
                listView1.Items.Clear();
                for (i = listView1.Columns.Count - 1; i >= 0; i--)
                    listView1.Columns.RemoveAt(i);

                listView1.Columns.Add("版本");
                listView1.Columns[0].Width = 600;
                i = 0;
                while (ipa[i].version != null)
                {
                    if ((ipa[i].version.ToLower().IndexOf(textBox1.Text) != -1))
                    {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = ipa[i].version;
                    listView1.Items.Add(item);
                    }
                    i++;
                }
            }
            private void buildOthers()
            {
                int i;
                listView1.Items.Clear();
                for (i = listView1.Columns.Count - 1; i >= 0; i--)
                    listView1.Columns.RemoveAt(i);

                listView1.Columns.Add("名称");
                listView1.Columns.Add("关键词");
                listView1.Columns[0].Width = 200;
                listView1.Columns[1].Width = 400;
                i = 0;
                while (others[i].name != null)
                {
                    if ((others[i].name.ToLower().IndexOf(textBox1.Text) != -1 || others[i].keywords.ToLower().IndexOf(textBox1.Text) != -1))
                    {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = others[i].name;
                    item.SubItems.Add(others[i].keywords);
                    listView1.Items.Add(item);
                    }
                    i++;
                }
            }
        //源代码变成结构体
            private void staSong(string code, ref Song[] song)
            {
                string temp = delStrLeft(code, "【");
                int i = 0;
                while (temp.IndexOf("【") != -1)
                {
                    song[i].num = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].name = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].id = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].variety = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].difficulty = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].combo = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].keywords = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    song[i].url = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    i++;
                }
            }
            private void staPack(string code, ref Pack[] pack)
            {
                string temp = delStrLeft(code, "【");
                int i=0;
                 while(temp.IndexOf("【") != -1)
                 {
                     pack[i].num = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     pack[i].name = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     pack[i].keywords = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     pack[i].date = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     pack[i].url = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     i++;
                 }
            }
            private void staIpa(string code, ref Ipa[] ipa)
            {
                string temp = delStrLeft(code, "【");
                int i=0;
                 while(temp.IndexOf("【") != -1)
                 {
                     ipa[i].num = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     ipa[i].version = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     ipa[i].url = delStrRight(temp, "】");
                     temp = temp = delStrLeft(temp, "【");
                     i++;
                 }
            }
            private void staOthers(string code, ref Others[] others)
            {
                string temp = delStrLeft(code, "【");
                int i = 0;
                while (temp.IndexOf("【") != -1)
                {
                    others[i].num = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    others[i].name = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    others[i].keywords = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    others[i].url = delStrRight(temp, "】");
                    temp = temp = delStrLeft(temp, "【");
                    i++;
                }    
            }
            private void staAll(ref string code, ref All all)
            {
                string temp=delStrLeft(code,"【");
                all.type = delStrRight(temp,"】");
                temp=delStrLeft(temp,"【");
                all.date = delStrRight(temp,"】");
                temp = delStrLeft(temp,"【");
                all.url = delStrRight(temp,"】");
                
            }
        //字符串删左删右函数           
        private string delStrLeft(string process, string keyStr) //清掉字符串左边的字符
            {
                int a = process.IndexOf(keyStr), b = process.Length;
                process = process.Substring(a + 1, b - a - 1);
                return process;
            }

            private string delStrRight(string process, string keyStr)//清掉字符串右边的字符
            {
                int a = process.IndexOf(keyStr);
                try
                {
                    process = process.Substring(0, a);
                }
                catch
                {

                }
                return process;
            }

            private void listView1_SelectedIndexChanged(object sender, EventArgs e)
            {
                if (listView1.SelectedItems.Count != 0) button1.Enabled = true;
                else button1.Enabled = false;
            }

            private void radioButton6_CheckedChanged(object sender, EventArgs e)
            {
                if (radioButton6.Checked == true) button6.Enabled=true;
                if (radioButton6.Checked == true && textBox2.Text == "") this.button6_Click(sender, e);
            }

            private void button6_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog dilog = new FolderBrowserDialog();
                dilog.Description = "请选择存放文件夹";
                if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
                    textBox2.Text = dilog.SelectedPath;
                else
                    radioButton5.Checked = true;
            }

            private void radioButton5_CheckedChanged(object sender, EventArgs e)
            {
                if (radioButton5.Checked == true) button6.Enabled = false;
            }


    }
}
