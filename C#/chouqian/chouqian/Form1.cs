using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;
 
namespace chouqian
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public int currentPlayer;

        public bool isyusaiqu=true;

        public struct playerScore 
        {
            public string id;
            public int group;
            public int song1,song2;
        }

        public playerScore[] player = new playerScore[43];

        bool moreThanThree;

        public string[] topThree=new string[3];

        int current=0;
        string[] nametxt = new string[10000];

        private void Form1_Load(object sender, EventArgs e)
        {
            var file = File.Open(Application.StartupPath + "\\chouqian.txt", FileMode.Open);

            int t = 0;
            using (var stream = new StreamReader(file))
            {
                while (!stream.EndOfStream)
                {
                    nametxt[t] = stream.ReadLine();
                    t++;
                }
            }

            current = t;

            for (int i=0;i<current;i++)
            {
                ListViewItem item = new ListViewItem();
                item.SubItems[0].Text = nametxt[i];
                listView1.Items.Add(item);
            }
            listView1.Columns[0].Width = 90;
            listView2.Columns[0].Width = 105;
            listView3.Columns[0].Width = 180;
            listView3.Columns[1].Width = 50;
            listView3.Columns[2].Width = 50;
        }
 
        private void button1_Click(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < listView1.Items.Count; i++)
                if (listView1.Items[i].Selected == true)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = listView1.Items[i].Text;
                    listView2.Items.Add(item);
                }
            for (i = listView1.Items.Count-1; i >=0 ; i--)
                if (listView1.Items[i].Selected == true)
                    listView1.Items.Remove(listView1.Items[i]);
            if (listView2.Items.Count>1) button5.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView2.Items.Count; i++)
                if (listView2.Items[i].Selected == true)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = listView2.Items[i].Text;
                    listView1.Items.Add(item);
                }
            for (int i = listView2.Items.Count-1; i >=0 ; i--)
                if (listView2.Items[i].Selected == true)
                    listView2.Items.Remove(listView2.Items[i]);
            if (listView2.Items.Count < 2) button5.Enabled=false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            bool flag = true;
            if (button5.Text == "预赛")
            {
                flag = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = true;
                listView1.Enabled = false;
                listView2.Enabled = false;
                
                for (int i = 0; i < listView2.Items.Count; i++)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = listView2.Items[i].Text;
                    item.SubItems.Add("无");
                    item.SubItems.Add("预赛曲");
                    item.SubItems.Add("0");
                    item.SubItems.Add("0");
                    listView3.Items.Add(item);

                    player[i].id = listView2.Items[i].Text;
                    player[i].song1 = 0;
                    player[i].song2 = 0;
                }
                currentPlayer = listView2.Items.Count;
                buildTotal();
                button5.Text = "下一轮";
                button5.Enabled = false;
            }

            if (button5.Text == "下一轮" && flag) 
            {
                if (currentPlayer % 2 == 0)
                {
                    isyusaiqu = false;
                    for (int i = 0; i < listView4.Items.Count; i++)
                    {
                        player[i].id = listView4.Items[i].Text;
                        player[i].group = -1;
                        player[i].song1 = 0;
                        player[i].song2 = 0;
                    }

                    int iSeed = 10,t;
                    Random ro = new Random(iSeed);
                    long tick = DateTime.Now.Ticks;
                    Random ran = new Random((int)(tick & 0xffffffffL) | (int)(tick >> 32));
                    for (int i = 0; i < currentPlayer / 2; i++)
                    {
                        t = ran.Next(0, currentPlayer);
                        while (player[t].group != -1) 
                        { 
                            Thread.Sleep(15); 
                            t = ran.Next(0, currentPlayer); 
                        }
                        player[t].group = i;
                        Thread.Sleep(15);
                        while (player[t].group != -1)
                        {
                            Thread.Sleep(15);
                            t = ran.Next(0, currentPlayer);
                        }
                        player[t].group = i;
                        Thread.Sleep(15);
                    }

                    listView3.Items.Clear();
                    for (int i = 0; i < currentPlayer/2;i++ )
                    {
                        for(int j=0;j<currentPlayer;j++)
                        {
                            if (player[j].group == i) 
                            {
                                ListViewItem item = new ListViewItem();
                                item.SubItems[0].Text = player[j].id;
                                item.SubItems.Add((i+1).ToString());
                                item.SubItems.Add("1");
                                item.SubItems.Add("0");
                                item.SubItems.Add("0");
                                listView3.Items.Add(item);
                                ListViewItem item2 = new ListViewItem();
                                item2.SubItems[0].Text = player[j].id;
                                item2.SubItems.Add((i+1).ToString());
                                item2.SubItems.Add("2");
                                item2.SubItems.Add("0");
                                item2.SubItems.Add("0");
                                listView3.Items.Add(item2);
                            }
                        }
                    }

                    buildTotal();

                    button5.Enabled = false;
                }
                else
                    MessageBox.Show("请确保剩余参赛人数为偶数！");
            }
            if (listView4.Items.Count > 3)
            {
                listView4.MultiSelect = true;
                moreThanThree = true;
            }
            else
            {
                listView4.MultiSelect = false;
                moreThanThree = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int iSeed = 10,varietyTab,songTab,direction1,direction2;
            String direction1CN,direction2CN;
            Random ro = new Random(iSeed);
            long tick = DateTime.Now.Ticks;
            Random ran = new Random((int)(tick & 0xffffffffL) | (int)(tick >> 32));
            varietyTab= ran.Next(1, 8);
            Thread.Sleep(15);
            songTab= ran.Next(1,54);
            Thread.Sleep(15);
            direction1 = ran.Next(1,3);
            Thread.Sleep(15);
            direction2 = ran.Next(1, 3);

            if (direction1 == 1)
                direction1CN = "左";
            else
                direction1CN = "右";

            if (direction2 == 1)
                direction2CN = "左";
            else
                direction2CN = "右";

            MessageBox.Show("请主持人在当前歌曲位置下，按"+varietyTab.ToString()+"下"+direction1CN+"箭头，再向"+direction2CN+"拨动"+songTab.ToString()+"首歌");
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            bool flag = true;
            int i,j;
            if (listView4.Items.Count - listView4.SelectedItems.Count >= 3 && moreThanThree)
            {
                for (i = listView4.Items.Count - 1; i >= 0; i--)
                {
                    if (listView4.Items[i].Selected == true)
                    {
                        listView4.Items[i].Selected = false;
                        for (j = listView3.Items.Count - 1; j >= 0; j--)
                        {
                            if (listView4.Items[i].Text == listView3.Items[j].Text)
                            {
                                listView3.Items.Remove(listView3.Items[j]);
                            }
                        }
                        listView4.Items.Remove(listView4.Items[i]);
                        if (listView4.Items.Count <= 3)
                        {
                            moreThanThree = false;
                        }
                        flag = false;
                    }
                }
            }
            else if (!moreThanThree) 
            {
                for (i = listView4.Items.Count - 1; i >= 0; i--)
                {
                    if (listView4.Items[i].Selected == true)
                    {
                        listView4.Items[i].Selected = false;
                        for (j = listView3.Items.Count - 1; j >= 0; j--)
                        {
                            if (listView4.Items[i].Text == listView3.Items[j].Text)
                            {
                                listView3.Items.Remove(listView3.Items[j]);
                            }
                        }
                        topThree[listView4.Items.Count-1]=listView4.Items[i].Text;
                        listView4.Items.Remove(listView4.Items[i]);
                        flag = false;
                    }
                }
            }
            else
            {
                MessageBox.Show("批量淘汰时，剩余玩家个数至少是3个以上！");
                flag = false;
            }

            currentPlayer = listView4.Items.Count;
            if (currentPlayer <= 3)
            {
                moreThanThree = false;
                listView4.MultiSelect = false;
            }
            if (flag) MessageBox.Show("请在总得分表中选中要被淘汰的人");

            if (listView4.Items.Count == 1)
            {
                MessageBox.Show("本次比赛：\n冠军：" 
                    + listView4.Items[0].Text + "\n"
                    + "亚军：" + topThree[1] + "\n" 
                    + "季军：" + topThree[2] + "\n");
                button5.Text = "预赛";
                listView2.Items.Clear();
                listView3.Items.Clear();
                listView4.Items.Clear();
                listView1.Enabled = true;
                listView2.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;
                button4.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;

                listView1.Items.Clear();
                for (int s = 0; s < current; s++)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = nametxt[s];
                    listView1.Items.Add(item);
                }
            }
            switch (listView4.Items.Count)
            {
                case 2: button7.Text = "评亚军"; break;
                case 3: button7.Text = "评季军"; break;
                default: button7.Text = "淘汰"; break;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form2 score = new Form2();
            int i;
            
            for (i = 0; i < currentPlayer; i++)
            {
                score.comboBox1.Items.Add(player[i].id);
            }

            if (isyusaiqu) score.comboBox2.Items.Add("预赛曲");
            else
            {
                score.comboBox2.Items.Add("1");
                score.comboBox2.Items.Add("2");
            }
            this.Enabled = false;

            score.Show(this);
        }
        public void buildTotal() 
        {
            listView4.Items.Clear();
            for (int i = 0; i < currentPlayer; i++) 
            {
                ListViewItem l4 = new ListViewItem();
                l4.SubItems[0].Text = player[i].id;
                l4.SubItems.Add((player[i].song1+player[i].song2).ToString());
                listView4.Items.Add(l4);
            }
        }
    }
}
