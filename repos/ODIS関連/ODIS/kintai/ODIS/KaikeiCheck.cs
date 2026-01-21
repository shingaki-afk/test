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
    public partial class KaikeiCheck : Form
    {
        //private string yy = ""; //2020
        //private string MM = ""; //06
        private string yyyyMM = ""; //202006
        //private string mae = ""; //202005
        //private string MMadd = ""; //07
        //private string ymd = ""; //2020/06/30

        public KaikeiCheck()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            kizyunday.Value = DateTime.Now;

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

            kaimiback.Checked = true;
            zeikubuncb.Checked = true;

            DateTime date = DateTime.Today;
            dateTimePicker1.Value = new DateTime(date.Year, date.Month, 1).AddMonths(-1);

            GetLastDate();

            GetData();
            GetZanData();

            if (Program.loginname == "喜屋武　大祐" || Program.loginname == "太田　朋宏" || Program.loginname == "管理者" || Program.loginname == "RPA用AC")
            {
                dateTimePicker1.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
            }
        }



        private void GetData()
        {
            DateTime datet = Convert.ToDateTime(this.comboBox1.SelectedItem.ToString() + "/" + this.comboBox2.SelectedItem.ToString() + "/01");
            yyyyMM = datet.ToString("yyyyMM");

            DataTable dt = new DataTable();

            string sql = "";

            sql += "select * from dbo.PCA会計仕訳データ_エラーチェック('" + yyyyMM + "')";
            
            if (zeikubuncb.Checked)
            { 
                sql += " where エラー内容 not like '税区分%' ";
            }
            //    sqsqll2 = "select * from k会計チェック_科目組織工事(" + yyyyMM + ", " + mae + ") where 組合せ微妙内容 <> '【もやもや】科目共通' and 組合せ微妙内容 <> '【確認】先月売上ゼロ(事務所/全現場除)' order by 組合せ微妙内容, 科目コード, 部門コード, 工事コード";

            dt = Com.GetDB(sql);
            dataGridView2.DataSource = dt;

            //dataGridView2.Columns[0].Width = 200;
            //dataGridView2.Columns[7].DefaultCellStyle.Format = "#,0";
            //dataGridView2.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            Com.InHistory("07_会計チェック", "", "");
        }

        private void GetZanData()
        {
            DataTable dt3 = new DataTable();

            string kizyun = Convert.ToDateTime(kizyunday.Value).ToString("yyyyMMdd");

            string sql3 = "";
            sql3 = "select * from dbo.k科目別取引先別マイナスチェック(" + kizyun + ") ";
            if (kaimiback.Checked) sql3 += " where 金額 < 0 ";
            sql3 += " order by 科目コード, 取引先コード, 部門コード, 金額 ";

            dt3 = Com.GetDB(sql3);
            dataGridView3.DataSource = dt3;

            dataGridView3.Columns[0].Width = 80;
            dataGridView3.Columns[1].Width = 80;
            dataGridView3.Columns[2].Width = 80;
            dataGridView3.Columns[3].Width = 130;
            dataGridView3.Columns[4].Width = 80;
            dataGridView3.Columns[5].Width = 200;
            dataGridView3.Columns[6].Width = 80;

            dataGridView3.Columns[6].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;
            GetData();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;
            GetData();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt = Com.GetDB("exec PCA会計仕訳データBulkInsert " + dateTimePicker1.Value.ToString("yyyyMMdd"));
            MessageBox.Show(dt.Rows[0][0].ToString());
            GetLastDate();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            GetZanData();
        }

        private void kizyunday_ValueChanged(object sender, EventArgs e)
        {
            GetZanData();
        }

        private void zeikubuncb_CheckedChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void GetLastDate()
        {
            DataTable dt = Com.GetDB("select max(処理日時) from dbo.処理履歴 where 処理項目 = 'PCA会計仕訳データコピー'");
            label5.Text = "最終更新：" + dt.Rows[0][0].ToString();
        }
    }
}
