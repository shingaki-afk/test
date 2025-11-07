using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using System.Data.SqlClient;

namespace ODIS.ODIS
{
    public partial class DetailsSKKSum : Form
    {
        private string _result;
        private string _group;
        private string _ys;
        private string _ye;
        private string _ymct;

        private string _count;

        private DataRowView _drv = null;

        public DetailsSKKSum()
        {
            InitializeComponent();
        }

        public DetailsSKKSum(string group, string ys, string ye, string result, string count)
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);
            //dataGridView2.Font = new Font(dataGridView2.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //諸経費じゃない場合は非表示
            if (group != "諸経費")
            {
                checkBox1.Visible = false;
            }

            _result = result;
            _group = group;
            _ys = ys;
            _ye = ye;

            _count = count;

            //年月数を取得する
            DataTable ymcount = Com.GetDB("select distinct 年月 from dbo.kanrikeisuu where 年月 between '" + ys + "' and '" + ye + "' ");
            _ymct = ymcount.Rows.Count.ToString();

            GetData(group, ys, ye, result);
        }


        private void GetData(string group, string ys, string ye, string result)
        {
            DataTable dt = new DataTable();

            string avflg = "";

            if (ys != ye) avflg = " / " + _ymct + " ";

            string sql = "";
            sql = "select 科目コード, 科目名, sum(金額) as 合計額, sum(金額) " + avflg + "as 月額, sum(金額) " + avflg + " / " + _count + " as [月額/一人当] from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql += "where left(伝票日付,6) between '" + ys + "' and '" + ye + "' ";

            if (group == "管理諸経費")
            {
                sql += "and 科目コード between '8300' and '8990' ";
            }
            else
            { 
                sql += "and 科目コード between '8250' and '8299' ";
            }

            sql += "and 摘要文 not like '給与__月支給分%' ";
            sql += "and 摘要文 not like '賞与___0%' ";
            sql += "and 科目コード not in ('8215','8346') ";
            sql += "and 摘要文 not like '%労災保険概算納付%' ";
            sql += "and 摘要文 not like '%月分社会保険料差額分（子ども子育て拠出金%' ";
            sql += "and 摘要文 not like '%全友協沖縄支部へ活動資金助成金を期末剰余金返金%' ";
            sql += "and 摘要文 not like '全友協沖縄支部より%年度資金剰余金を戻入支払' ";
            sql += "and 摘要文 not like '%全友協沖縄支部へ%月分会費として%' ";

            //材料費と外注費を除外
            if (!checkBox1.Checked)
            {
                sql += "and 科目コード not in ('8251','8281') ";
            }

            sql += result;

            sql += "group by 科目コード, 科目名 order by 科目コード";

            dt = Com.GetDB(sql);

            decimal sum = 0;
            decimal sumav = 0;
            decimal tanka = 0;

            DataTable Disp = new DataTable();

            //Disp.Columns.Add("伝票日付", typeof(string));
            Disp.Columns.Add("科目CD", typeof(string));
            Disp.Columns.Add("科目名", typeof(string));
            Disp.Columns.Add("合計額", typeof(decimal));
            Disp.Columns.Add("月額", typeof(decimal));
            Disp.Columns.Add("月額/一人当", typeof(decimal));
            //Disp.Columns.Add("摘要", typeof(string));
            //Disp.Columns.Add("伝票番号", typeof(string));

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = Disp.NewRow();
                //nr["伝票日付"] = row["伝票日付"];
                nr["科目CD"] = row["科目コード"];
                nr["科目名"] = row["科目名"];
                nr["合計額"] = row["合計額"];
                nr["月額"] = row["月額"];
                nr["月額/一人当"] = row["月額/一人当"];
                //nr["取引先"] = row["取引先名"];
                //nr["摘要"] = row["摘要文"];
                //nr["伝票番号"] = row["伝票番号"];
                sum += Convert.ToDecimal(row["合計額"]);
                sumav += Convert.ToDecimal(row["月額"]);
                tanka += Convert.ToDecimal(row["月額/一人当"]);
                Disp.Rows.Add(nr);
            }

            this.total.Text = sum.ToString("C");
            this.totalav.Text = sumav.ToString("C");
            this.count.Text = Convert.ToDecimal(_count).ToString("#人");
            this.avat.Text = tanka.ToString("C");


            dataGridView1.DataSource = Disp;

            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns[0].Width = 70;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 120;
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[4].Width = 150;
            //dataGridView1.Columns[5].Width = 500;

            //this.bumon.Text = bumon;
            //this.genba.Text = genba;
            if (ys == ye)
            {
                this.month.Text = ys;
            }
            else
            {
                this.month.Text = "年間";
            }
            this.koumoku.Text = group;


            //検索履歴登録
            Com.InHistory("51_管理計数_合計詳細表示", group + " " + ys, dataGridView1.RowCount.ToString());
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            dataGridView1.Enabled = false;


            //ソート対応
            BindingManagerBase bm = ((DataGridView)sender).BindingContext[((DataGridView)sender).DataSource, ((DataGridView)sender).DataMember];

            if (bm.Count == 0) return;

            DataRowView drv = (DataRowView)bm.Current;

            //前回と同じならスルー
            if (_drv == drv)
            {
                //マウスカーソルをデフォルトにする
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                dataGridView1.Enabled = true;

                return;
            }

            _drv = drv;

            string kamokucd = drv.Row.ItemArray[0].ToString();


            DataTable dt = new DataTable();

            string sql = "";
            sql = "select 部門コード, 部門名, 現場コード, 現場名, 伝票日付, 工種名, 金額, 摘要文 from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql += "where left(伝票日付,6) between '" + _ys + "' and '" + _ye + "' ";
            //sql += "and 科目コード between '8220' and '8600' ";
            sql += "and 摘要文 not like '給与__月支給分%' ";
            sql += "and 摘要文 not like '賞与___0%' ";
            sql += "and 科目コード not in ('8215','8346') ";
            sql += "and 摘要文 not like '%労災保険概算納付%' ";
            sql += "and 摘要文 not like '%月分社会保険料差額分（子ども子育て拠出金%' ";
            sql += "and 摘要文 not like '%全友協沖縄支部へ活動資金助成金を期末剰余金返金%' ";
            sql += "and 摘要文 not like '全友協沖縄支部より%年度資金剰余金を戻入支払' ";
            sql += "and 摘要文 not like '%全友協沖縄支部へ%月分会費として%' ";
            sql += "and 科目コード = '" + kamokucd + "' ";
            sql += _result;

            sql += "order by 部門コード, 現場コード";

            dt = Com.GetDB(sql);

            dataGridView2.DataSource = dt;


            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            dataGridView1.Enabled = true;
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            GetData(_group, _ys, _ye, _result);
        }
    }
}
