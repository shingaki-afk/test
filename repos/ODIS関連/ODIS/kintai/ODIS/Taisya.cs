using ODIS;
using ODIS.ODIS;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using C1.C1Excel;
using System.IO;

namespace ODIS.ODIS
{
    public partial class Taisya : Form
    {
        //表示者の登録有無フラグ
        string entryFlg = "";

        //元データ
        DataTable dt = new DataTable();

        //選択一覧データ
        DataTable SelectDisp = new DataTable();

        //エントリー一覧データ
        DataTable EntryDisp = new DataTable();

        //退職入力の制限
        TargetDays td = new TargetDays();

        public Taisya()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            //フォントサイズの変更
            dataGridView2.Font = new Font(dataGridView2.Font.Name, 10);

            //グリッドビューのコピーで
            //dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //デフォルト値
            taisyoku.Value = td.EndYMD;

            //有効期間
            taisyoku.MinDate = td.StartYMD.AddDays(-1);

            //未来日付だと表示しない問題
            taisyoku.MaxDate = td.EndYMD.AddMonths(3);

            //2021/02/05コメントアウト
            //taisyoku.MaxDate = td.StartYMD.AddMonths(1).AddDays(-1);

            //button3.Text = td.StartYMD.Month.ToString() + "月退職印刷";

            riyuu.Items.Add("定年");
            riyuu.Items.Add("自己都合退職");
            riyuu.Items.Add("会社都合退職");
            riyuu.Items.Add("雇用期間満了");
            riyuu.Items.Add("死亡");

            //初期表示値
            riyuu.SelectedIndex = 1;

            shiharai.Items.Add("振込");
            shiharai.Items.Add("現金");
            //初期表示値
            shiharai.SelectedIndex = 0;

            //選択一覧データの枠
            SelectDisp.Columns.Add("社員番号", typeof(string));
            SelectDisp.Columns.Add("氏名", typeof(string));
            SelectDisp.Columns.Add("組織名", typeof(string));
            SelectDisp.Columns.Add("現場名", typeof(string));
            SelectDisp.Columns.Add("登録状況", typeof(string));
            SelectDisp.Columns.Add("社保", typeof(string));
            SelectDisp.Columns.Add("友の会退職記念品", typeof(string));
            SelectDisp.Columns.Add("会社退職記念品", typeof(string));
            SelectDisp.Columns.Add("退職積立金", typeof(string));
            //SelectDisp.Columns.Add("メール区分", typeof(string));
            //SelectDisp.Columns.Add("メール備考", typeof(string));

            //SelectDisp.Columns.Add("携帯端末", typeof(string));
            //SelectDisp.Columns.Add("携帯区分", typeof(string));
            //SelectDisp.Columns.Add("携帯番号", typeof(string));

            //SelectDisp.Columns.Add("アカウント", typeof(string));


            //エントリー一覧データの枠
            EntryDisp.Columns.Add("社員番号", typeof(string));
            EntryDisp.Columns.Add("氏名", typeof(string));
            EntryDisp.Columns.Add("現場名", typeof(string));
            EntryDisp.Columns.Add("組織名", typeof(string));
            EntryDisp.Columns.Add("給与区分", typeof(string));
            EntryDisp.Columns.Add("支払方法", typeof(string));
            EntryDisp.Columns.Add("退職日", typeof(string));
            EntryDisp.Columns.Add("退職理由", typeof(string));
            EntryDisp.Columns.Add("入社日", typeof(string));
            EntryDisp.Columns.Add("特記事項", typeof(string));

            EntryDisp.Columns.Add("登録日時", typeof(string));
            EntryDisp.Columns.Add("登録者", typeof(string));
            //EntryDisp.Columns.Add("社保", typeof(string));
            EntryDisp.Columns.Add("友の会退職記念品", typeof(string));
            EntryDisp.Columns.Add("会社退職記念品", typeof(string));
            //EntryDisp.Columns.Add("更新日時", typeof(string));
            //EntryDisp.Columns.Add("更新者", typeof(string));
            //EntryDisp.Columns.Add("承認日時", typeof(string));
            //EntryDisp.Columns.Add("承認者", typeof(string));

            EntryDisp.Columns.Add("退職積立金", typeof(string));
            EntryDisp.Columns.Add("在籍年月", typeof(string));
            //EntryDisp.Columns.Add("メール区分", typeof(string));
            //EntryDisp.Columns.Add("メール備考", typeof(string));

            //EntryDisp.Columns.Add("携帯端末", typeof(string));
            //EntryDisp.Columns.Add("携帯区分", typeof(string));
            //EntryDisp.Columns.Add("携帯番号", typeof(string));

            //EntryDisp.Columns.Add("アカウント", typeof(string));
            EntryDisp.Columns.Add("担当区分", typeof(string));
            EntryDisp.Columns.Add("離職票", typeof(string));
            EntryDisp.Columns.Add("社保喪失", typeof(string));

            risyoku.Items.Add("");
            risyoku.Items.Add("〇");

            syaho.Items.Add("");
            syaho.Items.Add("〇");

            GetData();
            Refine();
            GetEntryData();

            Com.InHistory("32_退職入力", "", "");
        }

        //データ取得
        private void GetData()
        {
            dt.Clear();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        DataTable dtcb = new DataTable();
                        string sql = "";

                        if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
                        {
                            Cmd.CommandText = " select * from dbo.t退職データ取得 where 担当区分 in ('03_施設', '04_エンジ') order by 組織CD, 現場CD";
                            sql = "select distinct 担当区分 from dbo.t退職データ取得 where 担当区分 in ('03_施設', '04_エンジ') and 登録フラグ <> '0' order by 担当区分";
                        }
                        else if (Program.loginname == "金城　智之")
                        {
                            Cmd.CommandText = " select * from dbo.t退職データ取得 where 担当区分 like '%%' order by 組織CD, 現場CD";
                            sql = "select distinct 担当区分 from dbo.t退職データ取得 where 担当区分 like '%%' and 登録フラグ <> '0' order by 担当区分";
                        }
                        //TODO 2503大濱さん宮古島応援のため
                        else if (Program.loginname == "大浜　綾希子")
                        {
                            Cmd.CommandText = " select * from dbo.t退職データ取得 where 担当区分 like '%" + Program.loginbusyo + "%' or (担当区分 = '15_久米島' and 組織名 = '現業（久）') order by 組織CD, 現場CD";
                            sql = "select distinct 担当区分 from dbo.t退職データ取得 where (担当区分 like '%" + Program.loginbusyo + "%' or (担当区分 = '15_久米島' and 組織名 = '現業（久）')) and 登録フラグ <> '0' order by 担当区分";
                        }
                        else
                        { 
                            Cmd.CommandText = " select * from dbo.t退職データ取得 where 担当区分 like '%" + Program.loginbusyo + "%' order by 組織CD, 現場CD";
                            sql = "select distinct 担当区分 from dbo.t退職データ取得 where 担当区分 like '%" + Program.loginbusyo + "%' and 登録フラグ <> '0' order by 担当区分";
                        }
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);

                        //チェックボックス
                        dtcb = Com.GetDB(sql);

                        foreach (DataRow row in dtcb.Rows)
                        {
                            checkedListBox1.Items.Add(row["担当区分"]);
                        }

                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            checkedListBox1.SetItemChecked(i, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
        }

        //対象データ取得
        private void Refine()
        {
            SelectDisp.Clear();

            string res = search.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            foreach (DataRow row in dtrow)
            {
                DataRow nr = SelectDisp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["氏名"] = row["氏名"];
                nr["組織名"] = row["組織名"];
                nr["現場名"] = row["現場名"];

                if (row["登録フラグ"].ToString() == "1")
                {
                    nr["登録状況"] = "登録済 未承認";
                }
                else if (row["登録フラグ"].ToString() == "2")
                {
                    nr["登録状況"] = "承認済";
                }
                else
                {
                    nr["登録状況"] = "";
                }
                nr["社保"] = row["社保"];
                nr["友の会退職記念品"] = row["友の会退職記念品"];
                nr["会社退職記念品"] = row["会社退職記念品"];
                nr["退職積立金"] = row["退職積立金"];
                //nr["メール区分"] = row["メール区分"];
                //nr["メール備考"] = row["メール備考"];

                //nr["携帯端末"] = row["携帯端末"];
                //nr["携帯区分"] = row["携帯区分"];
                //nr["携帯番号"] = row["携帯番号"];

                //nr["アカウント"] = row["アカウント"];
                //nr["担当区分"] = row["担当区分"];

                SelectDisp.Rows.Add(nr);
            }

            dataGridView1.DataSource = SelectDisp;
        }

        //エントリーデータ取得
        private void GetEntryData()
        {
            EntryDisp.Clear();

            DataRow[] dtrow2;

            string result = "";

            //担当区分
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        result += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            dtrow2 = dt.Select("登録フラグ <> '0'" + result, "登録日時");



            //TODO 絞り込み設定をする
            //if (Program.loginID == "19300309")
            //{
            //    dtrow2 = dt.Select("登録フラグ <> '0' and 更新者 = '金城　智之'", "登録日時");
            //}
            //else
            //{
            //    dtrow2 = dt.Select("登録フラグ <> '0'", "登録日時");
            //}

            int ct1 = 0;
            int ct2 = 0;

            foreach (DataRow row in dtrow2)
            {
                DataRow nr = EntryDisp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["氏名"] = row["氏名"];
                nr["現場名"] = row["現場名"];
                nr["組織名"] = row["組織名"];
                nr["給与区分"] = row["給与区分"];
                nr["支払方法"] = row["支払方法"];
                nr["離職票"] = row["離職票"];
                nr["社保喪失"] = row["社保喪失"];
                nr["退職日"] = Convert.ToDateTime(row["退職日"]).ToString("yyyy/MM/dd");
                nr["退職理由"] = row["退職理由"];
                nr["入社日"] = row["入社年月日"];
                nr["特記事項"] = row["特記事項"];
                nr["登録日時"] = row["登録日時"];
                nr["登録者"] = row["登録者"];
                //nr["社保"] = row["社保"];
                nr["友の会退職記念品"] = row["友の会退職記念品"];
                nr["会社退職記念品"] = row["会社退職記念品"];
                //nr["更新日時"] = row["更新日時"];
                //nr["更新者"] = row["更新者"];
                //nr["承認日時"] = row["承認日時"];
                //nr["承認者"] = row["承認者"];

                nr["退職積立金"] = row["退職積立金"];
                nr["在籍年月"] = row["在籍年月"];
                //nr["メール区分"] = row["メール区分"];
                //nr["メール備考"] = row["メール備考"];

                //nr["携帯端末"] = row["携帯端末"];
                //nr["携帯区分"] = row["携帯区分"];
                //nr["携帯番号"] = row["携帯番号"];
                //nr["アカウント"] = row["アカウント"];
                nr["担当区分"] = row["担当区分"];

                if (row["登録フラグ"].ToString() == "1")
                {
                    ct1++;
                }
                else if (row["登録フラグ"].ToString() == "2")
                {
                    ct2++;
                }

                EntryDisp.Rows.Add(nr);
            }

            dataGridView2.DataSource = EntryDisp;

            //エントリー数を表示
            this.label8.Text = EntryDisp.Rows.Count.ToString() + "人登録" + "/" + ct1.ToString() + "/" + ct2.ToString();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            SyousaiData(drv[0].ToString());
        }

        private void SyousaiData(string str)
        {

            shiharai.SelectedIndex = 0;
            risyoku.SelectedIndex = 0;
            syaho.SelectedIndex = 0;

            //対象者のみに絞込
            DataRow[] targetDr = dt.Select("社員番号 = " + str.Substring(0, 8), "");

            foreach (DataRow row in targetDr)
            {
                number.Text = row["社員番号"].Equals(DBNull.Value) ? "" : row["社員番号"].ToString();　 //社員番号
                name.Text = row["氏名"].Equals(DBNull.Value) ? "" : row["氏名"].ToString();　　 //漢字氏名
                syokusyu.Text = row["組織名"].Equals(DBNull.Value) ? "" : row["組織名"].ToString(); //職種名
                genba.Text = row["現場名"].Equals(DBNull.Value) ? "" : row["現場名"].ToString();　　//現場名
                kubun.Text = row["給与区分"].Equals(DBNull.Value) ? "" : row["給与区分"].ToString();　　//給与支給区分名
                nyuusya.Text = row["入社年月日"].Equals(DBNull.Value) ? "" : row["入社年月日"].ToString();　　//入社年月日

                taisyoku.Text = row["退職日"].Equals(DBNull.Value) ? td.EndYMD.ToString() : row["退職日"].ToString();
                riyuu.SelectedItem = row["退職理由"].Equals(DBNull.Value) ? "" : row["退職理由"].ToString();

                shiharai.SelectedItem = row["支払方法"].Equals(DBNull.Value) ? "" : row["支払方法"].ToString();

                risyoku.SelectedItem = row["離職票"].Equals(DBNull.Value) ? "" : row["離職票"].ToString();
                syaho.SelectedItem = row["社保喪失"].Equals(DBNull.Value) ? "" : row["社保喪失"].ToString();

                etc.Text = row["特記事項"].Equals(DBNull.Value) ? "" : row["特記事項"].ToString();

                entryFlg = row["登録フラグ"].Equals(DBNull.Value) ? "" : row["登録フラグ"].ToString();　　//登録フラグ
                //label5.Text = "マイナンバー：　" + row["マイナンバー"].Equals(DBNull.Value) ? "" : row["マイナンバー"].ToString();　　//口座種類
                label5.Text = row["マイナンバー"].ToString();　　  //口座種類
                label6.Text = row["支払方法_ZEEM"].ToString();　　 //支払方法
                label7.Text = row["在籍年月"].ToString();　　      //在籍年月
                tomo.Text = row["友の会退職記念品"].ToString();　　//在籍年月
                kaisya.Text = row["会社退職記念品"].ToString();
                label22.Text = row["退職積立金"].ToString();
            }
        }

        //登録ボタンクリック
        private void button1_Click(object sender, EventArgs e)
        {
            //登録有無分岐必要
            DataInsert();

            GetData();
            Refine();
            GetEntryData();

            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();

            //初期表示値
            riyuu.SelectedIndex = 1;
            shiharai.SelectedIndex = 0;
            risyoku.SelectedIndex = 0;
            syaho.SelectedIndex = 0;

            etc.Text = "";

        }

        // データ登録・更新処理
        private void DataInsert()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            string cmd = "";
            if (entryFlg == "1")
            {
                cmd = "[dbo].[UpdateTaisyokuData]";
            }
            else
            {
                cmd = "[dbo].[InsertTaisyokuData]";

            }

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = cmd;

                    Cmd.Parameters.Add(new SqlParameter("対象年月", SqlDbType.VarChar));
                    Cmd.Parameters["対象年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("支払方法", SqlDbType.VarChar));
                    Cmd.Parameters["支払方法"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("退職日", SqlDbType.Date));
                    Cmd.Parameters["退職日"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("退職理由", SqlDbType.VarChar));
                    Cmd.Parameters["退職理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特記事項", SqlDbType.VarChar));
                    Cmd.Parameters["特記事項"].Direction = ParameterDirection.Input;


                    //if (entryFlg == "1")
                    //{ 
                    //    Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.SmallDateTime));
                    //    Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;

                    //    Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar));
                    //    Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;
                    //}
                    //else
                    //{
                    //    Cmd.Parameters.Add(new SqlParameter("登録日時", SqlDbType.SmallDateTime));
                    //    Cmd.Parameters["登録日時"].Direction = ParameterDirection.Input;

                    //    Cmd.Parameters.Add(new SqlParameter("登録者", SqlDbType.VarChar));
                    //    Cmd.Parameters["登録者"].Direction = ParameterDirection.Input;
                    //}

                    Cmd.Parameters.Add(new SqlParameter("登録日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["登録日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("登録者", SqlDbType.VarChar));
                    Cmd.Parameters["登録者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("離職票", SqlDbType.VarChar));
                    Cmd.Parameters["離職票"].Direction = ParameterDirection.Input;
                    
                    Cmd.Parameters.Add(new SqlParameter("社保喪失", SqlDbType.VarChar));
                    Cmd.Parameters["社保喪失"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD.ToString("yyyyMM");
                    Cmd.Parameters["社員番号"].Value = number.Text;
                    Cmd.Parameters["支払方法"].Value = shiharai.Text;
                    Cmd.Parameters["退職日"].Value = taisyoku.Text;
                    Cmd.Parameters["退職理由"].Value = riyuu.Text;
                    Cmd.Parameters["特記事項"].Value = etc.Text;

                    Cmd.Parameters["登録日時"].Value = DateTime.Now;
                    Cmd.Parameters["登録者"].Value = Program.loginname;

                    Cmd.Parameters["離職票"].Value = risyoku.Text;
                    Cmd.Parameters["社保喪失"].Value = syaho.Text;

                    //if (entryFlg == "1")
                    //{
                    //    Cmd.Parameters["更新日時"].Value = DateTime.Now;
                    //    Cmd.Parameters["更新者"].Value = Program.loginname;
                    //}
                    //else
                    //{
                    //    Cmd.Parameters["登録日時"].Value = DateTime.Now;
                    //    Cmd.Parameters["登録者"].Value = Program.loginname;
                    //}

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        //検索ボタン
        private void button3_Click(object sender, EventArgs e)
        {
            Refine();
        }

        //検索ボタンをエンター押下で実行
        private void search_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                Refine();
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            SyousaiData(drv[0].ToString());
        }

        //クリア処理
        private void button2_Click(object sender, EventArgs e)
        {
            if (entryFlg == "0") return;

            DialogResult result = MessageBox.Show("取消してよろしいですか？",
                        "警告",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) return;

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "delete from dbo.退職データ WHERE 社員番号 = '" + number.Text + "'";
                    using (dr = Cmd.ExecuteReader())
                    {
                        //TODO
                    }
                }
            }

            GetData();
            Refine();
            GetEntryData();

            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();
        }

        //印刷ボタン
        private void button3_Click_1(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            //button3.Enabled = false;



            this.TopMost = false;

            string fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\退職発令原票.xlsx";

            //エクセルオブジェクト
            Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();

            //エクセルを非表示
            m_MyExcel.Visible = false;

            //アラート非表示
            m_MyExcel.DisplayAlerts = false;

            //ブックオブジェクト
            Microsoft.Office.Interop.Excel.Workbook m_MyBook;

            //シートオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet m_MySheet;

            //ブックを開く
            m_MyBook = m_MyExcel.Workbooks.Open(Filename: fileName);

            //シート取得
            m_MySheet = m_MyBook.Worksheets[1];
            m_MySheet.Select();

            //EXCEL側の個人データが二行に分かれているため、+2となる変数を用意
            int ii = 0;

            //申請日
            m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
            //給与計算月
            m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
            //申請者
            m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
            //人数
            m_MySheet.Cells[34, 7] = EntryDisp.Rows.Count.ToString() + "名";



            //行別出力
            for (int i = 0; i < EntryDisp.Rows.Count; i++)
            {
                //10人分を区切りにシート分
                if (i == 10)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[2];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 20)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[3];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 30)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[4];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 40)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[5];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 50)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[6];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 60)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[7];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 70)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[8];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 80)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[9];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 90)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[10];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }
                else if (i == 100)
                {
                    ii = 0;
                    m_MySheet = m_MyBook.Worksheets[11];
                    m_MySheet.Select();
                    //申請日
                    m_MySheet.Cells[27, 1] = "申請日 ：   " + DateTime.Now.ToString("yyyy年MM月dd日");
                    //給与計算月
                    m_MySheet.Cells[29, 1] = "給与計算月 ：   " + td.EndYMD.AddDays(1).ToString("yyyy年MM月支給");
                    //申請者
                    m_MySheet.Cells[29, 5] = "申請者 ：   " + Program.loginname;
                }


                //列別出力
                for (int j = 0; j <= 9; j++)
                {
                    switch (j)
                    {
                        case 0:
                        case 1:
                        case 2:
                            //社員番号、氏名、現場名
                            m_MySheet.Cells[ii + 6, j + 1] = EntryDisp.Rows[i][j].ToString();
                            break;
                        case 3:
                        case 4:
                        case 5:
                            //職種名、給与支給区分、支払方法
                            m_MySheet.Cells[ii + 6, j + 2] = EntryDisp.Rows[i][j].ToString();
                            break;
                        case 6:
                            //発令日(退職日)
                            m_MySheet.Cells[ii + 6 + 1, j - 5] = Convert.ToDateTime(EntryDisp.Rows[i][6]).ToString("yyyy年MM月dd日");
                            break;
                        case 7:
                        case 8:
                        case 9:
                            //退職理由、入社日付、特記事項
                            m_MySheet.Cells[ii + 6 + 1, j - 4] = EntryDisp.Rows[i][j].ToString(); //"入社日付";
                            break;
                        default:
                            break;
                    }
                }

                ii = ii + 2;
            }



            string localPass = @"C:\ODIS\TAISYOKU\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_") ;

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            //保存前にシートを先頭に戻す
            m_MySheet = m_MyBook.Worksheets[1];
            m_MySheet.Select();

            //excel保存 ローカルへ
            m_MyBook.SaveAs(exlName + ".xlsx");

            m_MyBook.Close(false);
            m_MyExcel.Quit();


            //PDF
            //==========================================
            //Excel Change PDF           
            //Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
            m_MyExcel.Visible = false; //エクセルを非表示
                m_MyExcel.DisplayAlerts = false; //アラート非表示
                //Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
                                                                  //Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト


                //ブックを開く
                m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");
                //label6.Text = "もう少しですが、こっから長いです。。";

                //PDF保存
                m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");

                m_MyBook.Close(false);
                m_MyExcel.Quit();

                //excel出力
                //System.Diagnostics.Process.Start(@"c:\temp\test2.xlsx");
                //PDF出力
                System.Diagnostics.Process.Start(exlName + ".pdf");

            //==========================================

            //excel出力
            //System.Diagnostics.Process.Start(exlName + ".xlsx");




            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            //button3.Enabled = true;
        }

        private void label6_TextChanged(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Syounin("承認");

            GetData();
            Refine();
            GetEntryData();

            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();

            //初期表示値
            riyuu.SelectedIndex = 1;
            shiharai.SelectedIndex = 0;
            risyoku.SelectedIndex = 0;
            syaho.SelectedIndex = 0;
            etc.Text = "";
        }

        // データ登録・更新処理
        private void Syounin(string flg)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[退職承認]";

                    Cmd.Parameters.Add(new SqlParameter("対象年月", SqlDbType.VarChar));
                    Cmd.Parameters["対象年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("支払方法", SqlDbType.VarChar));
                    Cmd.Parameters["支払方法"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("退職日", SqlDbType.Date));
                    Cmd.Parameters["退職日"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("退職理由", SqlDbType.VarChar));
                    Cmd.Parameters["退職理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特記事項", SqlDbType.VarChar));
                    Cmd.Parameters["特記事項"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("承認日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["承認日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("承認者", SqlDbType.VarChar));
                    Cmd.Parameters["承認者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD.ToString("yyyyMM");
                    Cmd.Parameters["社員番号"].Value = number.Text;
                    Cmd.Parameters["支払方法"].Value = shiharai.Text;
                    Cmd.Parameters["退職日"].Value = taisyoku.Text;
                    Cmd.Parameters["退職理由"].Value = riyuu.Text;
                    Cmd.Parameters["特記事項"].Value = etc.Text;

                    if (flg == "承認")
                    { 
                        Cmd.Parameters["承認日時"].Value = DateTime.Now;
                        Cmd.Parameters["承認者"].Value = Program.loginname;
                    }
                    else
                    {
                        Cmd.Parameters["承認日時"].Value = DBNull.Value;
                        Cmd.Parameters["承認者"].Value = DBNull.Value;
                    }

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Syounin("取消");

            GetData();
            Refine();
            GetEntryData();

            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();

            //初期表示値
            riyuu.SelectedIndex = 1;
            shiharai.SelectedIndex = 0;
            risyoku.SelectedIndex = 0;
            syaho.SelectedIndex = 0;
            etc.Text = "";
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button7.Enabled = false;

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\60_退職発令原票.xlsx");

            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];

            int rows = EntryDisp.Rows.Count;
            int cols = EntryDisp.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[i + 1, j + 1].Value = EntryDisp.Rows[i][j].ToString();
                }
            }

            XLSheet ws = c1XLBook1.Sheets["原票"];


            ws[68, 0].Value = "　作成日　　：　" + DateTime.Now.ToString("yyyy年MM月dd日");
            ws[70, 0].Value = "給与計算月　：　" + td.StartYMD.ToString("yyyy年M月 ") + td.StartYMD.AddMonths(1).ToString(" (M月給与分)");
            ws[70, 4].Value = "作成者　：　" + Program.loginname;

            int ct = rows / 10 + 1;
            int ct2 = 1;
            for (int i = 1; i <= ct; i++)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = i.ToString();   // クローンをリネーム
                newSheet[0, 6].Value = ct2.ToString();      // 値の変更
                c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
                ct2 = ct2 + 10;
            }

            c1XLBook1.Sheets.Remove("原票");

            //TODO いる？
            //空
            //if (rows == 0)
            //{
            //    XLSheet newSheet = ws.Clone();
            //    newSheet.Name = "空";   // クローンをリネーム
            //    newSheet[0, 15].Value = "";      // 値の変更
            //    c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
            //}

            string localPass = @"C:\ODIS\TAISYOKU\";
            string exlName = localPass + td.StartYMD.ToString("yyyy年M月 ") + td.StartYMD.AddMonths(1).ToString(" (M月給与分)") + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");

            //Excel Change PDF           
            Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
            m_MyExcel.Visible = false; //エクセルを非表示
            m_MyExcel.DisplayAlerts = false; //アラート非表示
            Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
                                                              //Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト

            //ブックを開く
            m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");

            //PDF保存
            m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");

            m_MyBook.Close(false);
            m_MyExcel.Quit();

            FileInfo file = new FileInfo(exlName + ".xlsx");
            file.Delete();

            //PDF出力
            //System.Diagnostics.Process.Start(exlName + ".xlsx");
            System.Diagnostics.Process.Start(@"C:\ODIS\TAISYOKU\");
            System.Diagnostics.Process.Start(exlName + ".pdf");



            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button7.Enabled = true;
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
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

        private void checkedListBox1_MouseCaptureChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            GetEntryData();
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }
    }
}
