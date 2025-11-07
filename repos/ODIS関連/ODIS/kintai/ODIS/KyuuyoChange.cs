using Npgsql;
using ODIS.ODIS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace ODIS.ODIS
{
    public partial class KyuuyoChange : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();
        private DataTable wkdt = new DataTable();

        //TODO 毎年変更
        private string yms = "202504";
        private string yme = "202603";

        //TODO 毎年変更
        private string exyms = "202404";
        private string exyme = "202503";

        //TODO 前年の合計行の詳細表示で使用
        private string exymin = "";
        private string exymax = "";

        private string zi = "";
        private string zimax = "";



        public KyuuyoChange()
        {
            //MessageBox.Show("すみません。人件費と賞与引当金を項目分けしています。3/22までに完了予定です。");
            //return;

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dgvyosan.Font = new Font(dgvyosan.Font.Name, 10);
            dgvyosansum.Font = new Font(dgvyosansum.Font.Name, 10);
            dgvex.Font = new Font(dgvzisseki.Font.Name, 10);
            dgvzisseki.Font = new Font(dgvzisseki.Font.Name, 10);
            dgvlist.Font = new Font(dgvlist.Font.Name, 9);

            dgvyosan.RowHeadersVisible = false;
            dgvyosansum.RowHeadersVisible = false;
            dgvex.RowHeadersVisible = false;
            dgvzisseki.RowHeadersVisible = false;
            dgvlist.RowHeadersVisible = false;

            //ヘッダーの高さ
            dgvyosan.ColumnHeadersHeight = 10;
            dgvex.ColumnHeadersHeight = 10;

            // 選択モードを行単位での選択のみにする
            dgvlist.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            cbki.Items.Add("53期(2024)");
            cbki.Items.Add("54期(2025)");

            cbki.SelectedIndex = cbki.Items.Count - 1;

            //label6.Text = yms + "～" + yme;

            IniSet();

            Com.InHistory("52_予算更新_来期", "", "");
        }

        private void IniSet()
        {
            yms = cbki.SelectedItem.ToString().Substring(4, 4) + "04";
            yme = (Convert.ToInt16(cbki.SelectedItem.ToString().Substring(4, 4)) + 1).ToString() + "03";

            exyms = (Convert.ToInt16(cbki.SelectedItem.ToString().Substring(4, 4)) - 1).ToString() + "04";
            exyme = cbki.SelectedItem.ToString().Substring(4, 4) + "03";

            DataTable mokutable = new DataTable();
            mokutable = Com.GetDB("select 次 from dbo.y予算マスタ where 始年 = '" + yms.Substring(0, 4) + "' order by 次 ");

            cbzi.Items.Clear();

            foreach (DataRow row in mokutable.Rows)
            {
                cbzi.Items.Add("第" + row[0] + "次");
                zi = row[0].ToString();
                zimax = row[0].ToString();
            }

            cbzi.SelectedIndex = cbzi.Items.Count - 1;

            SetBumon();
            SetSyokusyu();

            //リスト表示
            GetDispData();
        }

        private void SetBumon()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = " select distinct 担当区分 from dbo.yosankanri a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '" + yms + "' and '" + yme + "' and 次 = '" + cbzi.SelectedItem.ToString().Substring(1,1) +  "' order by 担当区分 ";
            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox1.Items.Add(row["担当区分"]);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void SetSyokusyu()
        {
            checkedListBox2.Items.Clear();

            DataTable dt = new DataTable();
            string sql = " select distinct 職種 from dbo.yosankanri a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '" + yms + "' and '" + yme + "' and 次 = '" + cbzi.SelectedItem.ToString().Substring(1, 1) + "' ";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            sql += " order by 職種 ";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["職種"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        //リスト表示
        private void GetDispData()
        {
            string sql = "";
            sql = "Select max(a.部門名) as 部門名, max(a.現場名) as 現場名, max(最終更新日) as 最終更新日, a.部門CD, a.現場CD, b.職種, sum(case when 最終更新日 is null then 1 else 0 end) as 未処理年月数 from dbo.yosankanri a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '" + yms + "' and '" + yme + "' and 次 = '" + cbzi.SelectedItem.ToString().Substring(1, 1) + "' ";


            if (textBox1.Text != "")
            { 
                sql += " and a.現場名 like '%" + textBox1.Text + "%'";
            }

            //部門
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    sql += " and isnull(担当区分,'') <> '" + checkedListBox1.Items[i].ToString() + "'";
                }
            }

            //職種
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    sql += " and isnull(職種,'') <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            sql += " group by a.部門CD, a.現場CD, b.職種 order by a.部門CD, a.現場CD ";

            DataTable dt = new DataTable();
            dt = Com.GetDB(sql);
            dgvlist.DataSource = dt;

            int ymct = 0;
            foreach (DataRow row in dt.Rows)
            {
                ymct += Convert.ToInt32(row[6].ToString());
            }

            label1.Text = "未処理年月数 : " + ymct.ToString();

            dgvlist.Columns[0].Width = 100;
            dgvlist.Columns[1].Width = 250;
            //dgvlist.Columns[2].Width = 250;

        }

        //予算表示
        private void GetYosan(string bumon, string genba)
        {
            //グリッド表示クリア
            //dataGridView1.DataSource = "";
            //テーブルクリア
            dt.Clear();
            dgvyosansum.DataSource = "";

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            //string sql = "select * from dbo.目論見更新データ取得 where 部門ＣＤ like '%" + bumon + "%' and 現場ＣＤ like '%" + genba + "%'";
            string sql = "select 次,年月,部門CD,現場CD,部門名,現場名,固定売上,臨時売上,isnull(固定売上,0) + isnull(臨時売上,0) as 売上,人件費,賞与,退職金,諸経費,isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0) as 経費, ";
            sql += "isnull(固定売上,0) + isnull(臨時売上,0) - (isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) as 利益, case when (isnull(固定売上,0) + isnull(臨時売上,0)) = 0 then 0 else (isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 end as 計数, ";

            sql += " case when isnull(固定売上,0) + isnull(臨時売上,0) = 0 then 0 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 100 then 1 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 90 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 100 then 2 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 85 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 90 then 3 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 80 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 85 then 4 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 70 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 80 then 5 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 60 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 70 then 6 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 >= 50 and(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 60 then 7 ";
            sql += "when(isnull(諸経費,0) + isnull(人件費,0) + isnull(賞与,0) + isnull(退職金,0)) / (isnull(固定売上,0) + isnull(臨時売上,0)) * 100 < 50 then 8 ";
            sql += "else 9 end as 評価, ";

            sql += "管理人件費, 管理賞与, 管理退職金, 管理諸経費, 管理人件費 + 管理賞与 + 管理退職金 + 管理諸経費 as 管理経費, 備考, 最終更新日 from dbo.yosankanri where 年月 between '" + yms + "' and '" + yme + "' and 次 = '" + zi + "' and 部門CD like '%" + bumon + "%' and 現場CD like '%" + genba + "%'";


            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);

            //wkdt = Com.replaceDataTable(dt);

            dgvyosan.DataSource = dt;
            dgvyosan.Columns["次"].Width = 30;//次
            dgvyosan.Columns["年月"].Width = 60;//年月
            dgvyosan.Columns["部門CD"].Width = 50;//部門CD
            dgvyosan.Columns["現場CD"].Width = 50;//現場CD
            dgvyosan.Columns["部門名"].Width = 100;//部門名
            dgvyosan.Columns["現場名"].Width = 150;//現場名

            dgvyosan.Columns["固定売上"].Width = 90;//固定売上
            dgvyosan.Columns["臨時売上"].Width = 90;//臨時売上
            dgvyosan.Columns["売上"].Width = 90;//売上
            dgvyosan.Columns["人件費"].Width = 90;//人件費
            dgvyosan.Columns["賞与"].Width = 90;//賞与
            dgvyosan.Columns["退職金"].Width = 90;//退職金
            dgvyosan.Columns["諸経費"].Width = 90;//諸経費
            dgvyosan.Columns["経費"].Width = 90;//経費
            dgvyosan.Columns["利益"].Width = 90;//利益
            dgvyosan.Columns["計数"].Width = 60;//計数
            dgvyosan.Columns["評価"].Width = 60;//評価
            dgvyosan.Columns["管理人件費"].Width = 120;//管理人件費
            dgvyosan.Columns["管理賞与"].Width = 120;//管理賞与
            dgvyosan.Columns["管理退職金"].Width = 120;//管理退職金
            dgvyosan.Columns["管理諸経費"].Width = 120;//管理諸経費
            dgvyosan.Columns["管理経費"].Width = 120;//管理経費
            dgvyosan.Columns["備考"].Width = 200;//備考
            dgvyosan.Columns["最終更新日"].Width = 130;//最終更新日

            //ヘッダーの中央表示
            for (int i = 0; i < dgvyosan.Columns.Count; i++)
            {
                dgvyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //三桁区切り表示
            dgvyosan.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["利益"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvyosan.Columns["評価"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理経費"].DefaultCellStyle.Format = "#,0";

            //表示位置
            dgvyosan.Columns["次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosan.Columns["年月"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosan.Columns["部門CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosan.Columns["現場CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosan.Columns["部門名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosan.Columns["現場名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosan.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["評価"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosan.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["備考"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosan.Columns["最終更新日"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            ////読み取り専用
            this.dgvyosan.Columns["年月"].ReadOnly = true;
            this.dgvyosan.Columns["売上"].ReadOnly = true;
            this.dgvyosan.Columns["経費"].ReadOnly = true;
            this.dgvyosan.Columns["利益"].ReadOnly = true;
            this.dgvyosan.Columns["計数"].ReadOnly = true;
            this.dgvyosan.Columns["評価"].ReadOnly = true;
            this.dgvyosan.Columns["管理経費"].ReadOnly = true;
            this.dgvyosan.Columns["最終更新日"].ReadOnly = true;

            ////非表示
            dgvyosan.Columns["次"].Visible = false;
            dgvyosan.Columns["部門CD"].Visible = false;
            dgvyosan.Columns["現場CD"].Visible = false;
            dgvyosan.Columns["部門名"].Visible = false;
            dgvyosan.Columns["現場名"].Visible = false;
            //dataGridView1.Columns["最終更新日"].Visible = false;

            if (genbacd.Text.Substring(1, 2) == "99")
            {
                //現場CDが事務所

                this.dgvyosan.Columns["固定売上"].Visible = false;
                this.dgvyosan.Columns["臨時売上"].Visible = false;
                this.dgvyosan.Columns["売上"].Visible = false;

                this.dgvyosan.Columns["人件費"].Visible = false;
                this.dgvyosan.Columns["賞与"].Visible = false;
                this.dgvyosan.Columns["退職金"].Visible = false;
                this.dgvyosan.Columns["諸経費"].Visible = false;
                this.dgvyosan.Columns["経費"].Visible = false;
                this.dgvyosan.Columns["利益"].Visible = false;
                this.dgvyosan.Columns["計数"].Visible = false;
                this.dgvyosan.Columns["評価"].Visible = false;

                this.dgvyosan.Columns["管理人件費"].Visible = true;
                this.dgvyosan.Columns["管理賞与"].Visible = true;
                this.dgvyosan.Columns["管理退職金"].Visible = true;
                this.dgvyosan.Columns["管理諸経費"].Visible = true;
                this.dgvyosan.Columns["管理経費"].Visible = true;
            }
            else
            {
                this.dgvyosan.Columns["固定売上"].Visible = true;
                this.dgvyosan.Columns["臨時売上"].Visible = true;
                this.dgvyosan.Columns["売上"].Visible = true;

                this.dgvyosan.Columns["人件費"].Visible = true;
                this.dgvyosan.Columns["賞与"].Visible = true;
                this.dgvyosan.Columns["退職金"].Visible = true;
                this.dgvyosan.Columns["諸経費"].Visible = true;
                this.dgvyosan.Columns["経費"].Visible = true;
                this.dgvyosan.Columns["利益"].Visible = true;
                this.dgvyosan.Columns["計数"].Visible = true;
                this.dgvyosan.Columns["評価"].Visible = true;

                this.dgvyosan.Columns["管理人件費"].Visible = false;
                this.dgvyosan.Columns["管理賞与"].Visible = false;
                this.dgvyosan.Columns["管理退職金"].Visible = false;
                this.dgvyosan.Columns["管理諸経費"].Visible = false;
                this.dgvyosan.Columns["管理経費"].Visible = false;
            }

            ////色変更
            dgvyosan.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dgvyosan.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosan.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosan.Columns["利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            //行の高さ設定
            dgvyosan.RowTemplate.Height = 20;

            #region 合計
            //合計
            DataTable copydt = dt.Clone();
            copydt = Com.GetDB("select * from dbo.y予算更新合計欄取得('" + yms + "', '" + yme + "', '" + zi + "', '" + bumon + "', '" + genba + "')");

            dgvyosansum.DataSource = copydt;
            dgvyosansum.ColumnHeadersVisible = false;

            dgvyosansum.Columns["次"].Width = 30;
            dgvyosansum.Columns["年月"].Width = 60;
            dgvyosansum.Columns["部門CD"].Width = 50;
            dgvyosansum.Columns["現場CD"].Width = 50;
            dgvyosansum.Columns["部門名"].Width = 100;
            dgvyosansum.Columns["現場名"].Width = 150;
            dgvyosansum.Columns["固定売上"].Width = 90;
            dgvyosansum.Columns["臨時売上"].Width = 90;
            dgvyosansum.Columns["売上"].Width = 90;
            dgvyosansum.Columns["人件費"].Width = 90;
            dgvyosansum.Columns["賞与"].Width = 90;
            dgvyosansum.Columns["退職金"].Width = 90;
            dgvyosansum.Columns["諸経費"].Width = 90;
            dgvyosansum.Columns["経費"].Width = 90;
            dgvyosansum.Columns["利益"].Width = 90;
            dgvyosansum.Columns["計数"].Width = 60;
            dgvyosansum.Columns["評価"].Width = 60;
            dgvyosansum.Columns["管理人件費"].Width = 120;
            dgvyosansum.Columns["管理賞与"].Width = 120;
            dgvyosansum.Columns["管理退職金"].Width = 120;
            dgvyosansum.Columns["管理諸経費"].Width = 120;
            dgvyosansum.Columns["管理経費"].Width = 120;
            dgvyosansum.Columns["備考"].Width = 200;
            dgvyosansum.Columns["最終更新日"].Width = 130;


            //ヘッダーの中央表示
            for (int i = 0; i < dgvyosansum.Columns.Count; i++)
            {
                dgvyosansum.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //三桁区切り表示
            dgvyosansum.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["利益"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvyosansum.Columns["評価"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosansum.Columns["管理経費"].DefaultCellStyle.Format = "#,0";


            //表示位置
            dgvyosansum.Columns["次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosansum.Columns["年月"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosansum.Columns["部門CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosansum.Columns["現場CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosansum.Columns["部門名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosansum.Columns["現場名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosansum.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["評価"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosansum.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosansum.Columns["備考"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvyosansum.Columns["最終更新日"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            ////非表示
            dgvyosansum.Columns["次"].Visible = false;
            dgvyosansum.Columns["部門CD"].Visible = false;
            dgvyosansum.Columns["現場CD"].Visible = false;
            dgvyosansum.Columns["部門名"].Visible = false;
            dgvyosansum.Columns["現場名"].Visible = false;
            //dataGridView1.Columns["最終更新日"].Visible = false;

            if (genbacd.Text.Substring(1, 2) == "99")
            {
                //現場CDが事務所

                this.dgvyosansum.Columns["固定売上"].Visible = false;
                this.dgvyosansum.Columns["臨時売上"].Visible = false;
                this.dgvyosansum.Columns["売上"].Visible = false;

                this.dgvyosansum.Columns["人件費"].Visible = false;
                this.dgvyosansum.Columns["賞与"].Visible = false;
                this.dgvyosansum.Columns["退職金"].Visible = false;
                this.dgvyosansum.Columns["諸経費"].Visible = false;
                this.dgvyosansum.Columns["経費"].Visible = false;
                this.dgvyosansum.Columns["利益"].Visible = false;
                this.dgvyosansum.Columns["計数"].Visible = false;
                this.dgvyosansum.Columns["評価"].Visible = false;

                this.dgvyosansum.Columns["管理人件費"].Visible = true;
                this.dgvyosansum.Columns["管理賞与"].Visible = true;
                this.dgvyosansum.Columns["管理退職金"].Visible = true;
                this.dgvyosansum.Columns["管理諸経費"].Visible = true;
                this.dgvyosansum.Columns["管理経費"].Visible = true;
            }
            else
            {
                this.dgvyosansum.Columns["固定売上"].Visible = true;
                this.dgvyosansum.Columns["臨時売上"].Visible = true;
                this.dgvyosansum.Columns["売上"].Visible = true;

                this.dgvyosansum.Columns["人件費"].Visible = true;
                this.dgvyosansum.Columns["賞与"].Visible = true;
                this.dgvyosansum.Columns["退職金"].Visible = true;
                this.dgvyosansum.Columns["諸経費"].Visible = true;
                this.dgvyosansum.Columns["経費"].Visible = true;
                this.dgvyosansum.Columns["利益"].Visible = true;
                this.dgvyosansum.Columns["計数"].Visible = true;
                this.dgvyosansum.Columns["評価"].Visible = true;

                this.dgvyosansum.Columns["管理人件費"].Visible = false;
                this.dgvyosansum.Columns["管理賞与"].Visible = false;
                this.dgvyosansum.Columns["管理退職金"].Visible = false;
                this.dgvyosansum.Columns["管理諸経費"].Visible = false;
                this.dgvyosansum.Columns["管理経費"].Visible = false;
            }

            ////色変更
            dgvyosansum.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dgvyosansum.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosansum.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosansum.Columns["利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            #endregion 合計
        }

        private void GetZisseki(string bumon, string genba)
        {
            //実績！！
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.c管理計数取得('" + zi + "', '" + yms + "','" + yme + "') where 部門CD = '" + bumon + "' and 現場CD = '" + genba + "' ");

            dgvzisseki.DataSource = dt;
            //dgvzisseki.ColumnHeadersVisible = false;

            //dgvzisseki.Columns["次"].Width = 30;
            dgvzisseki.Columns["年月"].Width = 60;
            dgvzisseki.Columns["部門CD"].Width = 50;
            dgvzisseki.Columns["現場CD"].Width = 50;
            //dgvzisseki.Columns["部門名"].Width = 100;
            //dgvzisseki.Columns["現場名"].Width = 150;
            dgvzisseki.Columns["固定売上"].Width = 90;
            dgvzisseki.Columns["臨時売上"].Width = 90;
            dgvzisseki.Columns["売上"].Width = 90;
            dgvzisseki.Columns["人件費"].Width = 90;
            dgvzisseki.Columns["賞与"].Width = 90;
            dgvzisseki.Columns["退職金"].Width = 90;
            dgvzisseki.Columns["諸経費"].Width = 90;
            dgvzisseki.Columns["経費"].Width = 90;
            dgvzisseki.Columns["利益"].Width = 90;
            dgvzisseki.Columns["計数"].Width = 60;
            dgvzisseki.Columns["評価"].Width = 60;
            dgvzisseki.Columns["管理人件費"].Width = 120;
            dgvzisseki.Columns["管理賞与"].Width = 120;
            dgvzisseki.Columns["管理退職金"].Width = 120;
            dgvzisseki.Columns["管理諸経費"].Width = 120;
            dgvzisseki.Columns["管理経費"].Width = 120;
            dgvzisseki.Columns["予算利益差"].Width = 120;
            //dgvzisseki.Columns["備考"].Width = 200;
            //dgvzisseki.Columns["最終更新日"].Width = 130;

            //dgvzisseki.Columns[0].Width = 60;
            //dgvzisseki.Columns[1].Width = 70;
            //dgvzisseki.Columns[2].Width = 70;
            //dgvzisseki.Columns[3].Width = 70;
            //dgvzisseki.Columns[4].Width = 70;
            //dgvzisseki.Columns[5].Width = 70;
            //dgvzisseki.Columns[6].Width = 70;
            //dgvzisseki.Columns[7].Width = 70;
            //dgvzisseki.Columns[8].Width = 60;
            //dgvzisseki.Columns[9].Width = 60;
            //dgvzisseki.Columns[10].Width = 60;
            //dgvzisseki.Columns[11].Width = 60;
            //dgvzisseki.Columns[12].Width = 60;
            //dgvzisseki.Columns[13].Width = 60;

            //ヘッダーの中央表示
            for (int i = 0; i < dgvzisseki.Columns.Count; i++)
            {
                dgvzisseki.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvzisseki.Columns[i].ReadOnly = true;
            }

            //三桁区切り表示
            dgvzisseki.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["利益"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvzisseki.Columns["評価"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["管理経費"].DefaultCellStyle.Format = "#,0";
            dgvzisseki.Columns["予算利益差"].DefaultCellStyle.Format = "#,0";

            //表示位置
            //dgvzisseki.Columns["次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvzisseki.Columns["年月"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvzisseki.Columns["部門CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvzisseki.Columns["現場CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //dgvzisseki.Columns["部門名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //dgvzisseki.Columns["現場名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvzisseki.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["評価"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvzisseki.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvzisseki.Columns["予算利益差"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvzisseki.Columns["備考"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //dgvzisseki.Columns["最終更新日"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            ////色変更
            dgvzisseki.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dgvzisseki.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvzisseki.Columns["利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            dgvzisseki.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;



            if (genbacd.Text.Substring(1, 2) == "99")
            {
                //現場CDが事務所

                this.dgvzisseki.Columns["固定売上"].Visible = false;
                this.dgvzisseki.Columns["臨時売上"].Visible = false;
                this.dgvzisseki.Columns["売上"].Visible = false;

                this.dgvzisseki.Columns["人件費"].Visible = false;
                this.dgvzisseki.Columns["賞与"].Visible = false;
                this.dgvzisseki.Columns["退職金"].Visible = false;
                this.dgvzisseki.Columns["諸経費"].Visible = false;
                this.dgvzisseki.Columns["経費"].Visible = false;
                this.dgvzisseki.Columns["利益"].Visible = false;
                this.dgvzisseki.Columns["計数"].Visible = false;
                this.dgvzisseki.Columns["評価"].Visible = false;

                this.dgvzisseki.Columns["管理人件費"].Visible = true;
                this.dgvzisseki.Columns["管理賞与"].Visible = true;
                this.dgvzisseki.Columns["管理退職金"].Visible = true;
                this.dgvzisseki.Columns["管理諸経費"].Visible = true;
                this.dgvzisseki.Columns["管理経費"].Visible = true;
                //this.dgvzisseki.Columns["予算利益差"].Visible = false;
            }

            else
            {
                this.dgvzisseki.Columns["固定売上"].Visible = true;
                this.dgvzisseki.Columns["臨時売上"].Visible = true;
                this.dgvzisseki.Columns["売上"].Visible = true;

                this.dgvzisseki.Columns["人件費"].Visible = true;
                this.dgvzisseki.Columns["賞与"].Visible = true;
                this.dgvzisseki.Columns["退職金"].Visible = true;
                this.dgvzisseki.Columns["諸経費"].Visible = true;
                this.dgvzisseki.Columns["経費"].Visible = true;
                this.dgvzisseki.Columns["利益"].Visible = true;
                this.dgvzisseki.Columns["計数"].Visible = true;
                this.dgvzisseki.Columns["評価"].Visible = true;

                this.dgvzisseki.Columns["管理人件費"].Visible = false;
                this.dgvzisseki.Columns["管理賞与"].Visible = false;
                this.dgvzisseki.Columns["管理退職金"].Visible = false;
                this.dgvzisseki.Columns["管理諸経費"].Visible = false;
                this.dgvzisseki.Columns["管理経費"].Visible = false;
                //this.dgvzisseki.Columns["予算利益差"].Visible = true;
            }


            ////非表示
            dgvzisseki.Columns["部門CD"].Visible = false;
            dgvzisseki.Columns["現場CD"].Visible = false;
        }

        private void GetZenki(string bumon, string genba)
        {
            //前期
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.k管理計数取得('" + zi + "', '" + exyms + "','" + exyme + "') where 部門CD = '" + bumon + "' and 現場CD = '" + genba + "' ");

            if (dt.Rows.Count == 0)
            {
                dgvex.DataSource = null;
                return;
            }

            dgvex.DataSource = dt;

            dgvex.Columns["年月"].Width = 60;//年月
            dgvex.Columns["固定売上"].Width = 90;//固定売上
            dgvex.Columns["臨時売上"].Width = 90;//臨時売上
            dgvex.Columns["売上"].Width = 90;//売上
            dgvex.Columns["人件費"].Width = 90;//人件費
            dgvex.Columns["賞与"].Width = 90;//賞与
            dgvex.Columns["退職金"].Width = 90;//退職金
            dgvex.Columns["諸経費"].Width = 90;//諸経費
            dgvex.Columns["経費"].Width = 90;//経費
            dgvex.Columns["利益"].Width = 90;//利益
            dgvex.Columns["計数"].Width = 60;//計数
            dgvex.Columns["評価"].Width = 60;//評価
            dgvex.Columns["管理人件費"].Width = 120;//管理人件費
            dgvex.Columns["管理賞与"].Width = 120;//管理賞与
            dgvex.Columns["管理退職金"].Width = 120;//管理退職金
            dgvex.Columns["管理諸経費"].Width = 120;//管理諸経費
            dgvex.Columns["管理経費"].Width = 120;//管理経費
            dgvex.Columns["予算利益差"].Width = 120;//予算利益差
            //dgvex.Columns["予算利益差率"].Width = 120;//予算利益差率


            //ヘッダーの中央表示
            for (int i = 0; i < dgvex.Columns.Count; i++)
            {
                dgvex.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvex.Columns[i].ReadOnly = true;

            }


            //処理年月の最大値と最小値を取得する
            int[] syoriym = new int[dt.Rows.Count - 1];

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int val = 0;
                if (int.TryParse(dt.Rows[i][0].ToString(), out val))
                {
                    syoriym[i] = val;
                }
            }
            
            exymin= syoriym.Min().ToString();
            exymax = syoriym.Max().ToString();



            //三桁区切り表示
            dgvex.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["利益"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvex.Columns["評価"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["管理経費"].DefaultCellStyle.Format = "#,0";
            dgvex.Columns["予算利益差"].DefaultCellStyle.Format = "#,0";
            //dgvex.Columns["予算利益差率"].DefaultCellStyle.Format = "0.00\'%\'";


            //表示位置
            dgvex.Columns["年月"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvex.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvex.Columns["評価"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvex.Columns["予算利益差"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvex.Columns["予算利益差率"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            ////色変更
            dgvex.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dgvex.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvex.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvex.Columns["利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;


            if (genbacd.Text.Substring(1, 2) == "99")
            {
                //現場CDが事務所

                this.dgvex.Columns["固定売上"].Visible = false;
                this.dgvex.Columns["臨時売上"].Visible = false;
                this.dgvex.Columns["売上"].Visible = false;

                this.dgvex.Columns["人件費"].Visible = false;
                this.dgvex.Columns["賞与"].Visible = false;
                this.dgvex.Columns["退職金"].Visible = false;
                this.dgvex.Columns["諸経費"].Visible = false;
                this.dgvex.Columns["経費"].Visible = false;
                this.dgvex.Columns["利益"].Visible = false;
                this.dgvex.Columns["計数"].Visible = false;
                this.dgvex.Columns["評価"].Visible = false;

                this.dgvex.Columns["管理人件費"].Visible = true;
                this.dgvex.Columns["管理賞与"].Visible = true;
                this.dgvex.Columns["管理退職金"].Visible = true;
                this.dgvex.Columns["管理諸経費"].Visible = true;
                this.dgvex.Columns["管理経費"].Visible = true;
                //this.dgvex.Columns["予算利益差率"].Visible = false;
            }

            else
            {
                this.dgvex.Columns["固定売上"].Visible = true;
                this.dgvex.Columns["臨時売上"].Visible = true;
                this.dgvex.Columns["売上"].Visible = true;

                this.dgvex.Columns["人件費"].Visible = true;
                this.dgvex.Columns["賞与"].Visible = true;
                this.dgvex.Columns["退職金"].Visible = true;
                this.dgvex.Columns["諸経費"].Visible = true;
                this.dgvex.Columns["経費"].Visible = true;
                this.dgvex.Columns["利益"].Visible = true;
                this.dgvex.Columns["計数"].Visible = true;
                this.dgvex.Columns["評価"].Visible = true;

                this.dgvex.Columns["管理人件費"].Visible = false;
                this.dgvex.Columns["管理賞与"].Visible = false;
                this.dgvex.Columns["管理退職金"].Visible = false;
                this.dgvex.Columns["管理諸経費"].Visible = false;
                this.dgvex.Columns["管理経費"].Visible = false;
                //this.dgvex.Columns["予算利益差率"].Visible = true;
            }

            ////非表示
            dgvex.Columns["部門CD"].Visible = false;
            dgvex.Columns["現場CD"].Visible = false;

        }



        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //dt = Com.replaceDataTable(wkdt);


                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                MessageBox.Show("更新しました。");

                GetYosan(bumoncd.Text, genbacd.Text);
                GetZisseki(bumoncd.Text, genbacd.Text);
                GetZenki(bumoncd.Text, genbacd.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }


        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                
                GetDispData();
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //編集中のセル
            var cellEditing = dgvyosan.Rows[e.RowIndex].Cells[e.ColumnIndex];

            //編集中の列が「集金状況」の場合
            if (new string[] { "固定売上", "臨時売上", "人件費", "賞与", "退職金", "諸経費", "管理人件費", "管理賞与", "管理退職金", "管理諸経費" }.Contains(cellEditing.OwningColumn.Name))
            {
                int io = 0;
                bool result = int.TryParse(dgvyosan.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out io);

                if (!result)
                {
                    MessageBox.Show("数字を入力してください。");
                    dgvyosan[e.ColumnIndex, e.RowIndex].Value = 0;

                    return;
                }

                //最終行であればスルー
                if (e.RowIndex == 11)
                {
                    //dgvyosan[21, e.RowIndex].Value = DateTime.Now;
                    //return;
                }
                else
                { 

                //金額同じであればスルー
                if (dgvyosan.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value == dgvyosan.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) return;


                if (checkBox1.Checked)
                {
                    dgvyosan.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value = dgvyosan.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                }

                }
            }
            else
            {
                return;
            }

            //固定、臨時が変更
            decimal kotei = Convert.ToDecimal(dgvyosan[6, e.RowIndex].Value);
            decimal ringi = Convert.ToDecimal(dgvyosan[7, e.RowIndex].Value);
            decimal uri = kotei + ringi;

            decimal zinkenhi = Convert.ToDecimal(dgvyosan[9, e.RowIndex].Value);
            decimal syouyo = Convert.ToDecimal(dgvyosan[10, e.RowIndex].Value);
            decimal taisyo = Convert.ToDecimal(dgvyosan[11, e.RowIndex].Value);
            decimal syokeihi = Convert.ToDecimal(dgvyosan[12, e.RowIndex].Value);

            decimal keihi = syokeihi + zinkenhi + syouyo + taisyo;

            decimal rieki = uri - keihi;
            decimal keisu = uri == 0 ? 0 : keihi / uri * 100;
            int hyouka = 0;
            //売上
            dgvyosan[8, e.RowIndex].Value = uri;
            //経費
            dgvyosan[13, e.RowIndex].Value = keihi;
            //利益
            dgvyosan[14, e.RowIndex].Value = rieki;
            //計数
            dgvyosan[15, e.RowIndex].Value = keisu;

            //評価
            if (keisu == 0) hyouka = 0;
            else if (keisu > 100) hyouka = 1;
            else if (keisu > 90) hyouka = 2;
            else if (keisu > 85) hyouka = 3;
            else if (keisu > 80) hyouka = 4;
            else if (keisu > 70) hyouka = 5;
            else if (keisu > 60) hyouka = 6;
            else if (keisu > 50) hyouka = 7;
            else  hyouka = 8;

            dgvyosan[16, e.RowIndex].Value = hyouka;


            decimal kanzinkenhi = Convert.ToDecimal(dgvyosan[17, e.RowIndex].Value);
            decimal kansyouyo = Convert.ToDecimal(dgvyosan[18, e.RowIndex].Value);
            decimal kantaisyo = Convert.ToDecimal(dgvyosan[19, e.RowIndex].Value);
            decimal kansyokeihi = Convert.ToDecimal(dgvyosan[20, e.RowIndex].Value);
            decimal kankeihi = kansyokeihi + kanzinkenhi + kansyouyo + kantaisyo;
            //経費
            dgvyosan[21, e.RowIndex].Value = kankeihi;
            

            dgvyosan[23, e.RowIndex].Value = DateTime.Now;
            dgvyosansum[23, 0].Value = DateTime.Now;


            //合計行の更新
            if (genbacd.Text.Substring(1, 2) != "99")
            {
                decimal sumkoteiuri = 0;
                decimal sumringiuri = 0;

                decimal sumzinzi = 0;
                decimal sumsyouyo = 0;
                decimal sumtaisyo = 0;
                decimal sumsyo = 0;

                decimal sumkeisu = 0;
                int sumhyouka = 0;

                for (int i = 0; i < 12; i++)
                {
                    sumkoteiuri += Convert.ToDecimal(dgvyosan[6, i].Value);
                    sumringiuri += Convert.ToDecimal(dgvyosan[7, i].Value);

                    sumzinzi += Convert.ToDecimal(dgvyosan[9, i].Value);
                    sumsyouyo += Convert.ToDecimal(dgvyosan[10, i].Value);
                    sumtaisyo += Convert.ToDecimal(dgvyosan[11, i].Value);
                    sumsyo += Convert.ToDecimal(dgvyosan[12, i].Value);
                }

                dgvyosansum[6, 0].Value = sumkoteiuri; //固定売上
                dgvyosansum[7, 0].Value = sumringiuri; //臨時売上
                dgvyosansum[8, 0].Value = sumkoteiuri + sumringiuri; //売上

                dgvyosansum[9, 0].Value = sumzinzi; //人件費
                dgvyosansum[10, 0].Value = sumsyouyo; //賞与
                dgvyosansum[11, 0].Value = sumtaisyo; //退職金
                dgvyosansum[12, 0].Value = sumsyo; //諸経費
                dgvyosansum[13, 0].Value = sumzinzi + sumsyouyo + sumtaisyo + sumsyo; //経費

                dgvyosansum[14, 0].Value = sumkoteiuri + sumringiuri - sumzinzi - sumsyouyo - sumtaisyo - sumsyo; //利益
                if (sumkoteiuri + sumringiuri > 0)
                { 
                    dgvyosansum[15, 0].Value = sumkeisu = (sumzinzi + sumsyouyo + sumtaisyo + sumsyo) / (sumkoteiuri + sumringiuri) * 100; //計数
                }
                //評価
                if (sumkeisu == 0) sumhyouka = 0;
                else if (sumkeisu > 100) sumhyouka = 1;
                else if (sumkeisu > 90) sumhyouka = 2;
                else if (sumkeisu > 85) sumhyouka = 3;
                else if (sumkeisu > 80) sumhyouka = 4;
                else if (sumkeisu > 70) sumhyouka = 5;
                else if (sumkeisu > 60) sumhyouka = 6;
                else if (sumkeisu > 50) sumhyouka = 7;
                else sumhyouka = 8;
                dgvyosansum[16, 0].Value = sumhyouka; //評価

            }
            else
            {
                //事務所分
                decimal sumkanzin = 0;
                decimal sumsyouyo = 0;
                decimal sumtaisyo = 0;
                decimal sumkansyo = 0;

                for (int i = 0; i < 12; i++)
                {
                    sumkanzin += Convert.ToDecimal(dgvyosan[17, i].Value);
                    sumsyouyo += Convert.ToDecimal(dgvyosan[18, i].Value);
                    sumtaisyo += Convert.ToDecimal(dgvyosan[19, i].Value);
                    sumkansyo += Convert.ToDecimal(dgvyosan[20, i].Value);
                }

                dgvyosansum[17, 0].Value = sumkanzin; //管理人件費
                dgvyosansum[18, 0].Value = sumsyouyo; //管理賞与
                dgvyosansum[19, 0].Value = sumtaisyo; //管理退職金
                dgvyosansum[20, 0].Value = sumkansyo; //管理諸経費
                dgvyosansum[21, 0].Value = sumkanzin + sumsyouyo + sumtaisyo + sumkansyo; //管理経費

            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }
                
                //評価
                if (e.ColumnIndex == 16)
                {
                    GetLevelDispMoku(e, val);
                }
            }
        }

        public static void GetLevelDispMoku(DataGridViewCellFormattingEventArgs e, decimal val)
        {

                if (val == 0)
                {
                    e.Value = "-";
                }
                else if (val == 1)
                {
                    e.Value = "Ｅ";
                    e.CellStyle.BackColor = Color.Black;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (val == 2)
                {
                    e.Value = "Ｄ";
                    e.CellStyle.BackColor = Color.Gray;
                }
                else if (val == 3)
                {
                    e.Value = "Ｃ";
                    e.CellStyle.BackColor = Color.Crimson;
                }
                else if (val == 4)
                {
                    e.Value = "Ｂ";
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (val == 5)
                {
                    e.Value = "Ａ";
                    e.CellStyle.BackColor = Color.CornflowerBlue;
                }
                else if (val == 6)
                {
                    e.Value = "Ｓ";
                    e.CellStyle.BackColor = Color.LawnGreen;
                }
                else if (val == 7)
                {
                    e.Value = "ＳＳ";
                    e.CellStyle.BackColor = Color.Green;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (val == 8)
                {
                    e.Value = "ＳＳＳ";
                    e.CellStyle.BackColor = Color.Indigo;
                    e.CellStyle.ForeColor = Color.White;
                }
                else
                {
                    e.Value = "Error";
                }

                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private DataRowView _drv = null;

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {

            //ソート対応
            BindingManagerBase bm = dgvlist.BindingContext[dgvlist.DataSource, dgvlist.DataMember];
            DataRowView drv = (DataRowView)bm.Current;

            //前回と同じならスルー
            if (_drv == drv) return;

            _drv = drv;

            bumonname.Text = drv.Row.ItemArray[0].ToString();
            genbaname.Text = drv.Row.ItemArray[1].ToString();
            label7.Text = drv.Row.ItemArray[5].ToString();

            bumoncd.Text = drv.Row.ItemArray[3].ToString();
            genbacd.Text = drv.Row.ItemArray[4].ToString();

            GetYosan(drv.Row.ItemArray[3].ToString(), drv.Row.ItemArray[4].ToString());
            GetZisseki(drv.Row.ItemArray[3].ToString(), drv.Row.ItemArray[4].ToString());
            GetZenki(drv.Row.ItemArray[3].ToString(), drv.Row.ItemArray[4].ToString());

            //ヘッダーの高さ設定
            dgvyosan.ColumnHeadersHeight = 18;
            //行の高さ設定
            dgvyosan.RowTemplate.Height = 22;

        }

        private void label5_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }

            SetSyokusyu();
            GetDispData();
        }

        private void label23_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
            }

            GetDispData();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSyokusyu();
            GetDispData();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetDispData();
        }

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                if (e.ColumnIndex == 11)
                {
                    GetLevelDispMoku(e, val);
                }
            }
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                if (e.ColumnIndex == 11)
                {
                    GetLevelDispMoku(e, val);
                }
            }
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            decimal result;

            string group = dgvex.Columns[dgvex.CurrentCell.ColumnIndex].HeaderCell.Value.ToString();

            if (group == "固定売上" | group == "臨時売上" | group == "売上" | group == "人件費" | group == "賞与" | group == "退職金" | group == "諸経費" | group == "管理人件費" | group == "管理賞与" | group == "管理退職金" | group == "管理諸経費")
            {
                if ((group == "管理人件費" || group == "管理賞与" || group == "管理退職金") && genbacd.Text.Substring(1, 4) == "9900")
                {
                    if (Convert.ToInt16(Program.access) < 5) //部門長未満はみれない
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + bumonname.Text + " " + genbaname.Text, "");
                        return;
                    }
                    else if (Convert.ToInt16(Program.access) < 9 && bumoncd.Text == "11000")
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + bumonname.Text + " " + genbaname.Text, "");
                        return;
                    }
                }

                //数値以外はスルー
                if (decimal.TryParse(dgvex.CurrentCell.Value.ToString(), out result))
                {
                    //ゼロはスルー
                    if (result != 0)
                    {
                        //年月取得
                        string ym = dgvex.Rows[dgvex.CurrentCell.RowIndex].Cells[0].Value.ToString();

                        string ys = "";
                        string ye = "";

                        //合計の場合はbetween対応
                        if (ym == "年間")
                        {
                            ys = exymin;
                            ye = exymax;
                        }
                        else
                        {
                            ys = ym;
                            ye = ym;
                        }

                        //別フォームで表示
                        DetailsS_Yosan Detail = new DetailsS_Yosan(group, ys, ye, bumoncd.Text, genbacd.Text, bumonname.Text, genbaname.Text);
                        Detail.Show();
                    }
                }
            }

            //固定売上、臨時売上、物品売上、売上、人件費、諸経費以外はスルー
            //201605 従業員数追加

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //表示されているコントロールがDataGridViewTextBoxEditingControlか調べる
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                DataGridView dgv = (DataGridView)sender;

                //編集のために表示されているコントロールを取得
                DataGridViewTextBoxEditingControl tb =
                    (DataGridViewTextBoxEditingControl)e.Control;

                //イベントハンドラを削除
                tb.KeyPress -=
                    new KeyPressEventHandler(dataGridViewTextBox_KeyPress);

                //該当する列か調べる
                //if (dgv.CurrentCell.OwningColumn.Name == "管理人件費" )
                if (new string[] { "固定売上", "臨時売上", "人件費", "賞与", "退職金", "諸経費", "管理人件費", "管理賞与", "管理退職金", "管理諸経費" }.Contains(dgv.CurrentCell.OwningColumn.Name))
                {
                    //KeyPressイベントハンドラを追加
                    tb.KeyPress +=
                        new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
                }
            }
        }

        //DataGridViewに表示されているテキストボックスのKeyPressイベントハンドラ
        private void dataGridViewTextBox_KeyPress(object sender,KeyPressEventArgs e)
        {
            //全角数字を半角数字に変換
            string str = e.KeyChar.ToString();
            str = Microsoft.VisualBasic.Strings.StrConv(str, VbStrConv.Narrow);
            e.KeyChar = Convert.ToChar(str);

            //string test = str.Substring(1, 5);

            //数字しか入力できないようにする
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                //コピペはOK
                if (e.KeyChar == Strings.ChrW(22) || e.KeyChar == Strings.ChrW(3))
                {

                }
                else
                {
                    //MessageBox.Show("数字しか入力できません");
                    e.Handled = true;
                }
            }
            else
            {

            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\予算入力.xlsx"); return;
        }

        private void dgvyosan_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                if (e.ColumnIndex == 16)
                {
                    GetLevelDispMoku(e, val);
                }
            }

            if (Program.loginname != "喜屋武　大祐")
            {
                //TODO　入力可能期間において、入力可能月だけをコメントアウトする
                this.dgvyosan.Rows[0].ReadOnly = true;  //  4月
                this.dgvyosan.Rows[1].ReadOnly = true;  //  5月
                //this.dgvyosan.Rows[2].ReadOnly = true;  //  6月
                //this.dgvyosan.Rows[3].ReadOnly = true;  //  7月
                //this.dgvyosan.Rows[4].ReadOnly = true;  //  8月
                //this.dgvyosan.Rows[5].ReadOnly = true;  //  9月
                //this.dgvyosan.Rows[6].ReadOnly = true;  // 10月
                //this.dgvyosan.Rows[7].ReadOnly = true;  // 11月
                //this.dgvyosan.Rows[8].ReadOnly = true;  // 12月
                //this.dgvyosan.Rows[9].ReadOnly = true;  //  1月
                //this.dgvyosan.Rows[10].ReadOnly = true; //  2月
                //this.dgvyosan.Rows[11].ReadOnly = true; //  3月
            }
        }

        private void dgvzisseki_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            decimal result;

            string group = dgvzisseki.Columns[dgvzisseki.CurrentCell.ColumnIndex].HeaderCell.Value.ToString();

            if (group == "固定売上" | group == "臨時売上" | group == "売上" | group == "人件費" | group == "賞与" | group == "退職金" | group == "諸経費" | group == "管理人件費" | group == "管理賞与" | group == "管理退職金" | group == "管理諸経費")
            {
                if ((group == "管理人件費" || group == "管理賞与" || group == "管理退職金") && genbacd.Text.Substring(1, 4) == "9900")
                {
                    if (Convert.ToInt16(Program.access) < 5) //部門長未満はみれない
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + bumonname.Text + " " + genbaname.Text, "");
                        return;
                    }
                    else if (Convert.ToInt16(Program.access) < 9 && bumoncd.Text == "11000")
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + bumonname.Text + " " + genbaname.Text, "");
                        return;
                    }
                }

                //数値以外はスルー
                if (decimal.TryParse(dgvzisseki.CurrentCell.Value.ToString(), out result))
                {
                    //ゼロはスルー
                    if (result != 0)
                    {
                        //年月取得
                        string ym = dgvzisseki.Rows[dgvzisseki.CurrentCell.RowIndex].Cells[0].Value.ToString();

                        string ys = "";
                        string ye = "";

                        //合計の場合はbetween対応
                        if (ym == "年間")
                        {
                            ys = exymin;
                            ye = exymax;
                        }
                        else
                        {
                            ys = ym;
                            ye = ym;
                        }

                        //別フォームで表示
                        DetailsS_Yosan Detail = new DetailsS_Yosan(group, ys, ye, bumoncd.Text, genbacd.Text, bumonname.Text, genbaname.Text);
                        Detail.Show();
                    }
                }
            }

            //固定売上、臨時売上、物品売上、売上、人件費、諸経費以外はスルー
            //201605 従業員数追加

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void cbzi_SelectedIndexChanged(object sender, EventArgs e)
        {
            zi = cbzi.SelectedItem.ToString().Substring(1, 1);

            if (zi == zimax)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }


            if (Program.loginname != "喜屋武　大祐")
            {
                //TODO 入力できる期間の場合は下記をコメントアウトする
                //button1.Enabled = false;
            }


            SetBumon();
            SetSyokusyu();

            //リスト表示
            GetDispData();

            //GetYosan(bumoncd.Text, genbacd.Text);
        }

        private void cbki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null)
            {
                IniSet();
            }
        }
    }
}
