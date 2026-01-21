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
    public partial class TanmatsuK : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public TanmatsuK()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズ変更
            dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 8);

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("");
            comboBox1.Items.Add("00_ディスプレイ");
            comboBox1.Items.Add("01_PC");
            comboBox1.Items.Add("02_タブレット");
            comboBox1.Items.Add("03_携帯");
            comboBox1.Items.Add("04_メール");
            comboBox1.Items.Add("05_内線");
            comboBox1.Items.Add("06_ネット回線");
            comboBox1.Items.Add("07_電話回線");
            comboBox1.Items.Add("08_wi-fiルーター");
            comboBox1.Items.Add("09_kintone");


            comboBox2.Items.Add("");
            comboBox2.Items.Add("11_docomo");
            comboBox2.Items.Add("12_softbank");
            comboBox2.Items.Add("13_au");

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            


            //GetData();

            dataGridView1.Columns["No"].Width = 40;
            dataGridView1.Columns["組織CD"].Width = 40;
            dataGridView1.Columns["組織名"].Width = 100;
            dataGridView1.Columns["現場CD"].Width = 40;
            dataGridView1.Columns["現場名"].Width = 150;
            dataGridView1.Columns["社員番号"].Width = 60;
            dataGridView1.Columns["氏名"].Width = 120;
            dataGridView1.Columns["区分"].Width = 70;
            dataGridView1.Columns["使用区分"].Width = 75;
            dataGridView1.Columns["管理No(内線)"].Width = 70;
            dataGridView1.Columns["LSデバイス名"].Width = 70;
            dataGridView1.Columns["キャリア(メーカー)"].Width = 70;
            dataGridView1.Columns["種別"].Width = 70;
            dataGridView1.Columns["PC名電番Mail"].Width = 200;
            dataGridView1.Columns["タイプ"].Width = 70;
            dataGridView1.Columns["AD"].Width = 50;
            dataGridView1.Columns["OS"].Width = 100;
            dataGridView1.Columns["メモリ"].Width = 50;
            dataGridView1.Columns["備考"].Width = 150;
            dataGridView1.Columns["更新日"].Width = 100;

            dataGridView1.Columns["削除日"].Visible = false;

            Com.InHistory("92_端末管理テーブル設定", "", "");
        }

        private void GetData()
        {
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();
            string sql = "";

            string kubun = comboBox1.SelectedItem?.ToString();
            string career = comboBox2.SelectedItem?.ToString();
            sql = "SELECT [No],[組織CD],[組織名],[現場CD],[現場名],[社員番号],[氏名],[区分],[使用区分],[管理No(内線)],[LSデバイス名],[キャリア(メーカー)],[種別],[PC名電番Mail],[タイプ],[AD],[OS],[メモリ],[備考],[更新日],[削除日],[設定(基準ver)],[office] FROM [dbo].[t端末管理テーブル] where isnull(区分,'') like '%" + kubun + "%' and isnull([キャリア(メーカー)],'') like '%" + career + "%'";


            da = new SqlDataAdapter(sql, Cn);

            cb = new SqlCommandBuilder(da);

            da.Fill(dt);

            dataGridView1.DataSource = dt;

            count.Text = dt.Rows.Count.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }

            GetData();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
