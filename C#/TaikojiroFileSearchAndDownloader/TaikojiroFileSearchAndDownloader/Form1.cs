using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TaikojiroFileSearchAndDownloader
{
    struct file
    {
        public string name;
        public string keywords;
        public string url;
    }//定义了文件这么一个结构，包含名称、关键词、下载地址

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        file[] single = new file[2000];

        private void Form1_Load(object sender, EventArgs e)
        {
            
            string sc = sourceCodeProcess(getText("www.jq9837.com/taikojiro/data.htm"));
            sourceCodeToFile(ref single, sc);
            fileToList(single);
        }//运行时加载的东西

        private void fileToList(file[] file) 
        {
            int i = 0;
            do
            {
                ListViewItem item = new ListViewItem();
                item.SubItems[0].Text = file[i].name;
                item.SubItems.Add(file[i].keywords);
                listView1.Items.Add(item);
                i++;
            } while (file[i].name != null);
        }//将文件数组元素全部添加到列表

        private void sourceCodeToFile(ref file[] file,String sc)
        { 
            int i = 0;
            sc = sc.Replace("\r","");
            sc = sc.Replace("\n", "");
            sc = sc + "[";
            while (sc.IndexOf(']') != -1)
            {
                sc = delStrLeft(sc, "]");
                file[i].name = delStrRight(sc, "[");
                sc = delStrLeft(sc, "]");
                file[i].keywords = delStrRight(sc, "[");
                sc = delStrLeft(sc, "]");
                file[i].url = delStrRight(sc, "[");
                int a = sc.IndexOf("]");
                i++;
            }
        }//变好看的源码变成数组里的元素

        private string getText(string url)
        {
            webBrowser1.Navigate(url);
            return webBrowser1.DocumentText;
        }//获得网页源代码

        private string sourceCodeProcess(string sc) 
        {
            sc = sc.Replace("\n", "");
            sc = sc.Replace("&lt;name&gt;", System.Environment.NewLine+"[名称]");
            sc = sc.Replace("&lt;keywords&gt;", "[关键词]");
            sc = sc.Replace("&lt;url&gt;", "[下载地址]");
            sc = sc.Replace("&gt;", "");
            sc = sc.Replace("&lt;/url","");
            sc = sc.Replace("</p>", "");
            sc = sc.Replace("<p>", "");
            sc = sc.Replace("<html><head><meta http-equiv=\"Content-Language\" content=\"zh-cn\"><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><title>kobowadeさんの次郎配布センター</title></head><body>kobowadeさんの次郎配布センター　", "");
            sc = sc.Replace("&lt;/name", "");
            sc = sc.Replace("&lt;/keywords", "");
            sc = sc.Replace("</body></html>", "");
            return sc;
        }//将源代码恶心的地方去掉，变成可以看的东西

        private string delStrLeft(string process, string keyStr) //清掉字符串左边的字符
        {
            int a = process.IndexOf(keyStr),b=process.Length;
            process=process.Substring(a+1,b-a-1);
            return process;
        }

        private string delStrRight(string process, string keyStr)//清掉字符串右边的字符
        {
            int a = process.IndexOf(keyStr);
            process = process.Substring(0, a);
            return process;
        }

        private void filter(file[] file, string text) //（筛选搜索）搜索栏有什么字，文件数组中名称和关键词有它的都添加到列表里去
        {
            for (int i = 0; i < file.Length;i++)
            {
                if (file[i].name != null && (file[i].name.ToLower().IndexOf(text.ToLower()) != -1 || file[i].keywords.ToLower().IndexOf(text.ToLower()) != -1)) 
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = file[i].name;
                    item.SubItems.Add(file[i].keywords);
                    listView1.Items.Add(item);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)//搜索栏的字一改变，就将原来的列表清空，进行一次筛选搜索
        {
            listView1.Items.Clear();
            filter(single, textBox1.Text);
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)//同上，这是搜单曲的
        {
            if (radioButton2.Checked == true) 
            {
                listView1.Items.Clear();
                filter(single,textBox1.Text);
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)//同上，但这是搜相关下载的
        {
            listView1.Items.Clear();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)//同上，但这是搜整合的
        {
            listView1.Items.Clear();
        }
    }
}
