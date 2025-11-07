using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class TeateKoujoCheck : Form
    {
        private SqlConnection Cn;
        //private SqlCommand Cmd;

        //サンエー控除対象者一覧
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        /// <summary>
        /// 対象期間インスタンス
        /// </summary>
        private TargetDays td = new TargetDays();

        public TeateKoujoCheck()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            checkedListBox1.Items.Add("固定他１");
            checkedListBox1.Items.Add("固定他２");
            checkedListBox1.Items.Add("変動他１");
            checkedListBox1.Items.Add("変動他２");
            checkedListBox1.Items.Add("臨作業手当");
            checkedListBox1.Items.Add("臨時手当");
            checkedListBox1.Items.Add("前払(+)");
            checkedListBox1.Items.Add("持株奨励金");
            checkedListBox1.Items.Add("琉銀財形");
            checkedListBox1.Items.Add("沖銀財形");
            checkedListBox1.Items.Add("日本生命");
            checkedListBox1.Items.Add("住友生命");
            checkedListBox1.Items.Add("日動火災");
            checkedListBox1.Items.Add("ＡＦＬＡＣ");
            checkedListBox1.Items.Add("団体傷害");
            checkedListBox1.Items.Add("アライアンス");
            checkedListBox1.Items.Add("朝日生命");
            checkedListBox1.Items.Add("フコク生命");
            checkedListBox1.Items.Add("ＴＯＰ２１");
            checkedListBox1.Items.Add("積立金");
            checkedListBox1.SetItemChecked(0, true);
            checkedListBox1.SetItemChecked(1, true);
            checkedListBox1.SetItemChecked(2, true);
            checkedListBox1.SetItemChecked(3, true);
            checkedListBox1.SetItemChecked(4, true);
            checkedListBox1.SetItemChecked(5, true);
            checkedListBox1.SetItemChecked(6, true);
            checkedListBox1.SetItemChecked(7, true);
            checkedListBox1.SetItemChecked(8, true);
            checkedListBox1.SetItemChecked(9, true);
            checkedListBox1.SetItemChecked(10, true);
            checkedListBox1.SetItemChecked(11, true);
            checkedListBox1.SetItemChecked(12, true);
            checkedListBox1.SetItemChecked(13, true);
            checkedListBox1.SetItemChecked(14, true);
            checkedListBox1.SetItemChecked(15, true);
            checkedListBox1.SetItemChecked(16, true);
            checkedListBox1.SetItemChecked(17, true);
            checkedListBox1.SetItemChecked(18, true);
            checkedListBox1.SetItemChecked(19, true);
            GetMainData();
            GetData();

            memo.Text = "変01_作業靴代(自己負担分)" + Com.nl;
            memo.Text += "変02_寮変動費" + Com.nl;
            memo.Text += "変03_携帯個人負担分" + Com.nl;
            memo.Text += "変04_社宅変動費" + Com.nl;
            memo.Text += "変05_OIC食堂食事券購入" + Com.nl;
            memo.Text += "変06_濃度オゾン発生装置購入代" + Com.nl;
            memo.Text += "変07_食堂利用代" + Com.nl;
            memo.Text += "変08_モノレール定期券代" + Com.nl;

            //memo.Text += "臨01現場代理人手当" + Com.nl;
            //memo.Text += "臨02現場代理人補佐手当" + Com.nl;
            memo2.Text += "臨04_責任者手当" + Com.nl;
            memo2.Text += "臨06_遠隔監視業務" + Com.nl;
            memo2.Text += "臨07_緊急連絡先当番業務" + Com.nl;
            memo2.Text += "臨08_夜間待機手当" + Com.nl;
            memo2.Text += "臨09_応援時移動手当" + Com.nl;
            memo2.Text += "臨11_労災休業補償額" + Com.nl;
            memo2.Text += "臨12_入社祝金" + Com.nl;
            memo2.Text += "臨13_送迎手当" + Com.nl;
            memo2.Text += "臨90_下水道3種 登録手当" + Com.nl; 
            memo2.Text += "臨99_誤計算差額" + Com.nl;

            memo3.Text += "前01_紹介謝礼金" + Com.nl;
            memo3.Text += "前02_入社祝金" + Com.nl;
            memo3.Text += "現場手当充当分" + Com.nl;
            memo3.Text += "未設定" + Com.nl;
        }

        private void GetMainData()
        {
            dataGridView1.DataSource = null;
            dt.Clear();


            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select a.項目, a.内容, a.社員番号, a.氏名, No, 金額, 備考, (select b.退職年月日 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 退職年月日, (select b.組織名 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 組織名, (select b.現場名 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 現場名, a.管理No from dbo.固定控除 a ";

            //TODO ダミー
            sql += " where 項目 <> 'ダミー' ";

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    //sql += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    sql += " and (内容 like '%" + s + "%' or 内容 like '%" + Com.isOneByteChar(s) + "%' or 内容 like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or 内容 like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or 内容 like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or 内容 like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%' ";
                    sql += " or 氏名 like '%" + s + "%' or 氏名 like '%" + Com.isOneByteChar(s) + "%' or 氏名 like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or 氏名 like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or 氏名 like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or 氏名 like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //契約区分
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        sql += " and 項目 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            sql += " order by No";

            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }

            GetData();
        }

        private void GetData()
        {
            dataGridView3.DataSource = null;

            DataTable dt3 = new DataTable();
            string y = td.StartYMD.AddMonths(1).ToString("yyyy");
            string m = td.StartYMD.AddMonths(1).ToString("MM");
            string s = td.StartYMD.ToString("yyyy/MM/dd");
            string e = td.EndYMD.ToString("yyyy/MM/dd");
            //dataGridView3.DataSource = Com.GetDB("exec k固定変動控除と臨時手当差分チェック '" + y + "', '" + m + "', '" + r + "'");

            string sql = "select * from k固定変動臨時ZeeM比較('" + y + "', '" + m + "', '" + s + "', '" + e + "') ";

            sql += " where 項目 <> 'ダミー' ";

            //契約区分
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        sql += " and 項目 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            dataGridView3.DataSource = Com.GetDB(sql);
            label1.Text = dataGridView3.RowCount.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetMainData();
            GetData();

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;


        }

        private void tikulbl_Click(object sender, EventArgs e)
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
        }
    }
}
