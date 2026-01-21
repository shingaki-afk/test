using Microsoft.VisualBasic;
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
    public partial class TaisyokuK : Form
    {
        private DataTable dt = new DataTable();
        public TaisyokuK()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            list.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            syousai.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            if (Convert.ToInt16(Program.access) >= 8)
            {
                GetData();
                DispData();
                Com.InHistory("47_退職金", "", "");
            }
            else
            {
                MessageBox.Show("権限ないすー");
                Com.InHistory("47_退職金_権限無表示", "", "");
            }
        }

        private void DispData()
        {
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');


            string result = "";

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }
            
            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            DataTable Disp = new DataTable();
            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("氏名", typeof(string));
            Disp.Columns.Add("地区名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));
            Disp.Columns.Add("入社年月日", typeof(string));
            Disp.Columns.Add("本採用年月", typeof(string));
            Disp.Columns.Add("退職金_自己都合", typeof(int));
            Disp.Columns.Add("退職金_会社都合", typeof(int));
            Disp.Columns.Add("基本給", typeof(int));
            Disp.Columns.Add("支給率", typeof(decimal));
            Disp.Columns.Add("自己都合計数", typeof(decimal));
            Disp.Columns.Add("在籍月数", typeof(int));
            Disp.Columns.Add("対象月数", typeof(int));
            Disp.Columns.Add("欠格数合計", typeof(int));
            Disp.Columns.Add("60才越数", typeof(int));
            Disp.Columns.Add("月給以外", typeof(int));
            Disp.Columns.Add("日給", typeof(int));
            Disp.Columns.Add("締変更月", typeof(int));
            Disp.Columns.Add("勤務数不足", typeof(int));
            Disp.Columns.Add("差", typeof(int));
            Disp.Columns.Add("日給加算", typeof(int));

            int zikosum = 0;
            int kaisum = 0;

            foreach (DataRow row in dtrow)
            {
                DataRow nr = Disp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["氏名"] = row["氏名"];
                nr["地区名"] = row["地区名"];
                nr["現場名"] = row["現場名"];
                nr["入社年月日"] = row["入社年月日"];
                nr["本採用年月"] = row["本採用年月"];
                nr["退職金_自己都合"] = row["退職金_自己都合"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["退職金_自己都合"]);
                nr["退職金_会社都合"] = row["退職金_会社都合"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["退職金_会社都合"]);
                nr["基本給"] = row["基本給"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["基本給"]);
                nr["支給率"] = row["支給率"].Equals(DBNull.Value) ? 0 : Convert.ToDecimal(row["支給率"]);
                nr["自己都合計数"] = row["自己都合計数"].Equals(DBNull.Value) ? 0 : Convert.ToDecimal(row["自己都合計数"]);
                nr["在籍月数"] = Convert.ToInt64(row["在籍月数"]);
                //nr["対象月数"] = Convert.ToInt64(row["対象月数"]);
                //nr["欠格数合計"] = Convert.ToInt64(row["欠格数合計"]);
                //nr["60才越数"] = Convert.ToInt64(row["60才越数"]);
                //nr["月給以外"] = Convert.ToInt64(row["月給以外"]);
                //nr["日給"] = Convert.ToInt64(row["日給"]);
                //nr["締変更月"] = Convert.ToInt64(row["締変更月"]);
                //nr["勤務数不足"] = Convert.ToInt64(row["勤務数不足"]);
                //nr["差"] = Convert.ToInt64(row["差"]);
                //nr["日給加算"] = Convert.ToInt64(row["日給加算"]);
                nr["対象月数"] = row["対象月数"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["対象月数"]);
                nr["欠格数合計"] = row["欠格数合計"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["欠格数合計"]);
                nr["60才越数"] = row["60才越数"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["60才越数"]);
                nr["月給以外"] = row["月給以外"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["月給以外"]);
                nr["日給"] = row["日給"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["日給"]);
                nr["締変更月"] = row["締変更月"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["締変更月"]);
                nr["勤務数不足"] = row["勤務数不足"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["勤務数不足"]);
                nr["差"] = row["差"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["差"]);
                nr["日給加算"] = row["日給加算"].Equals(DBNull.Value) ? 0 : Convert.ToInt64(row["日給加算"]);
                Disp.Rows.Add(nr);

                zikosum += Convert.ToInt32(nr["退職金_自己都合"]);
                kaisum += Convert.ToInt32(nr["退職金_会社都合"]);
            }

            label1.Text = String.Format("退職金_自己都合 合計　{0:#,0}", zikosum);
            label2.Text = String.Format("退職金_会社都合 合計　{0:#,0}", kaisum); 

            list.DataSource = Disp;

            list.Columns[0].Width = 80;
            list.Columns[1].Width = 120;
            list.Columns[2].Width = 50;
            list.Columns[3].Width = 150;
            list.Columns[4].Width = 70;
            list.Columns[5].Width = 70;
            list.Columns[6].Width = 60;
            list.Columns[7].Width = 60;
            list.Columns[8].Width = 60;
            list.Columns[9].Width = 30;
            list.Columns[10].Width = 30;
            list.Columns[11].Width = 30;
            list.Columns[12].Width = 30;
            list.Columns[13].Width = 30;
            list.Columns[14].Width = 30;
            list.Columns[15].Width = 30;
            list.Columns[16].Width = 30;
            list.Columns[17].Width = 30;
            list.Columns[18].Width = 30;
            list.Columns[19].Width = 30;
            list.Columns[20].Width = 30;
            //list.Columns[21].Visible = false;

            list.Columns[6].DefaultCellStyle.Format = "#,0";
            list.Columns[7].DefaultCellStyle.Format = "#,0";
            list.Columns[8].DefaultCellStyle.Format = "#,0";

            list.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            list.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void GetData()
        {

            dt = Com.GetDB("select * from dbo.t退職金リスト order by 社員番号");
        }

        private void list_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void list_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            //ソートエラー対応
            DataGridViewRow dgr = list.CurrentRow;
            if (dgr == null) return;

            syousai.DataSource = "";

            string syainno = list.Rows[list.CurrentCell.RowIndex].Cells[0].Value.ToString();
            name.Text = list.Rows[list.CurrentCell.RowIndex].Cells[1].Value.ToString();
            syousai.DataSource = Com.GetDB("select 処理年月, 基本給, 給与区分, 所定 + 法休 + 所休 + 有給 + 特休 + 無特 as 基準日数, 対象 from t退職金元データ取得 where 社員番号 = '" + syainno + "' order by 処理年月");

            syousai.Columns[0].Width = 80;
            syousai.Columns[1].Width = 80;
            syousai.Columns[2].Width = 80;
            syousai.Columns[3].Width = 80;
            syousai.Columns[4].Width = 70;
            syousai.Columns[4].Width = 120;

            syousai.Columns[1].DefaultCellStyle.Format = "#,0";
            syousai.Columns[3].DefaultCellStyle.Format = "#,0";

            syousai.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            syousai.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //カーソル変更
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DispData();
        }

        private void TaisyokuK_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\17_総務部\04_給与\毎月給与計算業務\90_対応マニュアル\46_退職金_未対応一覧.xlsx"); return;
        }
    }
}
