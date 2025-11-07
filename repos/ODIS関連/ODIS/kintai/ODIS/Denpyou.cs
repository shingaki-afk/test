using C1.C1Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Denpyou : Form
    {
        private DataTable dt = new DataTable();

        private string yy = ""; //2020
        private string MM = ""; //06
        private string yyyyMM = ""; //202006
        private string mae = ""; //202005
        private string MMadd = ""; //07
        private string ymd = ""; //2020/06/30

        public Denpyou()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            textBox2.Text = "docomo ご利用料金管理サービス-データ集計・分析-確定情報-一括請求サービスグループ別-ご利用料金内訳-電話番号一覧、下4桁2113を選択。sqlserverテーブルdocomo利用金額でdel-ins";

            comboBox1.Items.Add("2023");
            comboBox1.Items.Add("2024");
            comboBox1.Items.Add("2025");
            string y = DateTime.Now.AddDays(-15).ToString("yyyy");
            comboBox1.SelectedIndex = comboBox1.FindString(y);

            comboBox2.Items.Add("01");
            comboBox2.Items.Add("02");
            comboBox2.Items.Add("03");
            comboBox2.Items.Add("04");
            comboBox2.Items.Add("05");
            comboBox2.Items.Add("06");
            comboBox2.Items.Add("07");
            comboBox2.Items.Add("08");
            comboBox2.Items.Add("09");
            comboBox2.Items.Add("10");
            comboBox2.Items.Add("11");
            comboBox2.Items.Add("12");

            string m = DateTime.Now.AddDays(-15).ToString("MM");
            comboBox2.SelectedIndex = comboBox2.FindString(m);

            //Bulk();

            DataTable listdt = Com.GetDB("select 伝票番号, max(摘要文), max(借方取引先名) from dbo.PCA会計仕訳データ where 伝票日付 like '" + yyyyMM + "%' and 摘要文 like '【%' group by 伝票番号");
            dataGridView2.DataSource = listdt;

        }

        private void GetYM()
        {
            DateTime dt = Convert.ToDateTime(this.comboBox1.SelectedItem.ToString() + "/" + this.comboBox2.SelectedItem.ToString() + "/01");
            yy = dt.ToString("yyyy");
            MM = dt.ToString("MM");
            yyyyMM = dt.ToString("yyyyMM");
            mae = dt.AddMonths(-1).ToString("yyyyMM");
            MMadd = dt.AddMonths(1).ToString("MM");
            ymd = dt.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd");
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;
            GetYM();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;
            GetYM();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from [dbo].[h日立リースPC相違チェック]('', '" + yyyyMM + "')");
            dataGridView1.DataSource = dt;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from [dbo].[dドコモ相違チェック]('', '" + yyyyMM + "')");
            dataGridView1.DataSource = dt;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from [dbo].[gGoogleアカウント相違チェック]('', '" + yyyyMM + "')");
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from [dbo].[gGoogleCOMアカウント相違チェック]('" + yyyyMM + "')");
            dataGridView1.DataSource = dt;
        }
    }
}
