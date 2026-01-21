using ODIS.ODIS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class CopyReason : Form
    {
        private string name; 
        private  IDataObject cb; 

        public CopyReason()
        {
            InitializeComponent();
        }

        public CopyReason(string s)
        {
            InitializeComponent();
            button1.Enabled = false;
            this.ControlBox = !this.ControlBox;
            name = s;
            cb = Clipboard.GetDataObject();
            Clipboard.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(cb);

            //履歴登録
            Com.InHistory(name, richTextBox1.Text, "コピー");

            this.Close();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox1.TextLength >= 5)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //履歴登録
            Com.InHistory("21_従業員検索", "キャンセル", "コピー");
            Clipboard.SetText("キャンセルしたさぁ");
            this.Close();
        }
    }
}
