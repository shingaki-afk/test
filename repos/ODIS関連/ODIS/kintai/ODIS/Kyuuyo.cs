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
    public partial class Kyuuyo : Form
    {
        private string sDate = ""; 
        private string eDate = "";

        public Kyuuyo()
        {

            if (Convert.ToInt16(Program.access) < 8)
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("48_月支給額平均_権限制限", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //境界線
            splitContainer1.SplitterDistance = 300;

            //日付設定
            DateTime today = DateTime.Today;
            if (today.Day >= 15)
            {
                yms.Value = today.AddYears(-1);
                yme.Value = today;
            }
            else
            {
                yms.Value = today.AddYears(-1).AddMonths(-1);
                yme.Value = today.AddMonths(-1);
            }

            GetData();

            Com.InHistory("48_月支給額平均", "", "");
        }

        private void GetData()
        {
            Cursor.Current = Cursors.WaitCursor;

            sDate = yms.Value.ToString("yyyyMM");
            eDate = yme.Value.ToString("yyyyMM");

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

            //除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                foreach (string s in ar2)
                {
                    result += " and (reskey not like '%" + s + "%' and reskey not like '%" + Com.isOneByteChar(s) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            ////先頭が「and」の場合、削除する
            //if (result.StartsWith(" and"))
            //{
            //    result = result.Remove(0, 4);
            //}

            DataTable dt = new DataTable();



            dt = Com.GetDB("select * from dbo.s支給額平均取得('" + sDate + "','" + eDate + "') where reskey like '%%' " + result + " order by 開始年月");
            dataGridView1.DataSource = dt;

            //0社員番号
            //1氏名
            //2役職名
            //3支給区分
            //4地区名
            //5組織名
            //6現場名
            //7開始年月
            //8最終年月

            //9労働時間平均
            //10月数
            //11固定平均
            //12変動平均
            //13支給平均
            //14reskey

            //右寄左寄
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[11].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[12].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[13].DefaultCellStyle.Format = "#,##0";

            dataGridView1.Columns[0].Width = 90;
            dataGridView1.Columns[1].Width = 130;
            dataGridView1.Columns[2].Width = 90;
            dataGridView1.Columns[3].Width = 90;
            dataGridView1.Columns[4].Width = 90;
            dataGridView1.Columns[5].Width = 150;
            dataGridView1.Columns[6].Width = 200;
            dataGridView1.Columns[7].Width = 60;
            dataGridView1.Columns[8].Width = 60;
            
            dataGridView1.Columns[9].Width = 60;
            dataGridView1.Columns[10].Width = 60;

            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].Width = 80;


            dataGridView1.Columns[14].Visible = false; //reskey

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void yms_ValueChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void yme_ValueChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

            //ソートエラー対応
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;

            string tantouk = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            DataTable dt = new DataTable();

            string sql = "";

            sql = "select 処理年 +処理月 as 支給年月,社員番号, 氏名, 地区名, 組織名, 現場名, 役職名, 支給区分, 週労働数, 勤務時間, ";
            //固定合計、変動合計
            sql += "本給+職務技能給+調整手当+特別手当+皆勤手当+役職手当+現場手当+免許手当+離島手当+扶養手当+転勤手当+通勤非課税+通勤課税+登録手当+通信手当+車両手当+持株奨励金 as 固定合計,";
            sql += "延長手当+法休出手当+所休出手当+残業手当+[60超残手当]+深夜手当+回数手当１+回数手当２+臨時手当+臨作業手当+正月期末+[前払金(+)]+臨休業手当+欠勤控除 as 変動合計,";
            sql += "本給+職務技能給+調整手当+特別手当+皆勤手当+役職手当+現場手当+免許手当+離島手当+扶養手当+転勤手当+通勤非課税+通勤課税+登録手当+通信手当+車両手当+持株奨励金+延長手当+法休出手当+所休出手当+残業手当+[60超残手当]+深夜手当+回数手当１+回数手当２+臨時手当+臨作業手当+正月期末+[前払金(+)]+臨休業手当+欠勤控除 as 支給合計,";
            //支給額
            sql += "本給,職務技能給, 調整手当,特別手当, 皆勤手当,役職手当,現場手当,免許手当,離島手当,扶養手当,転勤手当,通勤非課税,通勤課税,登録手当,通信手当,車両手当,退職積立金,持株奨励金,延長手当,法休出手当,所休出手当,残業手当,[60超残手当],深夜手当,回数手当１,回数手当２,臨時手当,臨作業手当,正月期末,[前払金(+)],臨休業手当,欠勤控除,支給合計額,";
            //控除
            sql += "健保,介保,厚年,雇保,所得税,住民税,財形積立,生命保険,友の会,固定他１,固定他２,積立金,[前払金(-)],変動他１,変動他２,差押金,年調過不足額,";
            sql += "控除合計額,";
            //勤怠
            sql += "延長時間,法休時間,所休時間,残業時間,[60超残Ｈ],深夜時間,遅刻回数,遅刻時間,時給,所定,法休,所休,有給,特休,無特,振休,公休,調休,届欠,無届,回数１,回数２,";
            //その他情報
            sql += "通勤1日単価,標準報酬月額,有給残日数,差引支給額 ";
            sql += "from dbo.KM_給与明細 where 社員番号 = '" + tantouk + "' and 処理年 +処理月 between '" + sDate + "' and '" + eDate + "' order by 処理年 +処理月";

           dt = Com.GetDB(sql);

            dataGridView2.DataSource = dt;

            for (int i = 8; i < dt.Columns.Count; i++)
            {
                dataGridView2.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) GetData();
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) GetData();
        }
    }
}
