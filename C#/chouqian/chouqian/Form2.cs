using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace chouqian
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1 main = (Form1)this.Owner;
            main.Enabled = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "" && textBox1.Text != "" && textBox2.Text != "")
                button1.Enabled = true;
            else
                button1.Enabled = false;

            int curr = this.textBox1.SelectionStart;
            System.Text.RegularExpressions.Regex rg = new System.Text.RegularExpressions.Regex("[^0-9]+");
            this.textBox1.Text = rg.Replace(this.textBox1.Text, "");
            if (this.textBox1.SelectionStart == 0)
            {
                this.textBox1.SelectionStart = 0;
            }
            else
            {
                this.textBox1.SelectionStart = curr;
            }
        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "" && textBox1.Text != "" && textBox2.Text != "")
                button1.Enabled = true;
            else
                button1.Enabled = false;

            int curr = this.textBox2.SelectionStart;
            System.Text.RegularExpressions.Regex rg = new System.Text.RegularExpressions.Regex("[^0-9]+");
            this.textBox2.Text = rg.Replace(this.textBox2.Text, "");
            if (this.textBox2.SelectionStart == 0)
            {
                this.textBox2.SelectionStart = 0;
            }
            else
            {
                this.textBox2.SelectionStart = curr;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            if (comboBox1.Text != "" && comboBox2.Text!="" && textBox1.Text != "" && textBox2.Text != "")
                button1.Enabled = true;
            else
                button1.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 main = (Form1)this.Owner;
            int i, j;
            for (i = 0; i < main.listView3.Items.Count; i++) 
            {
                if(main.listView3.Items[i].Text==comboBox1.Text && main.listView3.Items[i].SubItems[2].Text==comboBox2.Text)
                {
                    main.listView3.Items[i].SubItems[3].Text = textBox1.Text;
                    main.listView3.Items[i].SubItems[4].Text = textBox2.Text;
                    for (j = 0; j < main.currentPlayer; j++)
                    {
                        if(main.player[j].id==comboBox1.Text)
                        {
                            if(comboBox2.Text=="2")
                                main.player[j].song2 = int.Parse(textBox1.Text) * 2 - int.Parse(textBox2.Text);
                            else
                                main.player[j].song1 = int.Parse(textBox1.Text) * 2 - int.Parse(textBox2.Text);
                        }
                    }
                    main.buildTotal();
                    break;
                }
            }
            main.button5.Enabled = true;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            if (comboBox1.Text != "" && comboBox2.Text != "" && textBox1.Text != "" && textBox2.Text != "")
                button1.Enabled = true;
            else
                button1.Enabled = false;
        }
    }
}
