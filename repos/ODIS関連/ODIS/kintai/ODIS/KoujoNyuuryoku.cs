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

namespace ODIS.ODIS
{
    public partial class KoujoNyuuryoku : Form
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

        public KoujoNyuuryoku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            //フォントサイズの変更
            dataGridView2.Font = new Font(dataGridView2.Font.Name, 10);

            //グリッドビューのコピーで
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            koumoku.Items.Add("固定他１");
            koumoku.Items.Add("固定他２");
            koumoku.Items.Add("変動他１");
            koumoku.Items.Add("変動他２");

            //初期表示値
            koumoku.SelectedIndex = 1;

            naiyou.Items.Add("固01_浦添駐車場代");
            naiyou.Items.Add("固02_借入金/立替返済"); 
            naiyou.Items.Add("固03_北部駐車場代");
            naiyou.Items.Add("固04_サンエー駐車場代");
            naiyou.Items.Add("固05_新都心駐車代");
            naiyou.Items.Add("固06_ダブルツリー駐車場代");
            naiyou.Items.Add("固07_二十日会賄費");
            naiyou.Items.Add("固08_寮固定費"); 

            naiyou.Items.Add("変01_作業靴代(自己負担分)");
            naiyou.Items.Add("変02_寮変動費");
            naiyou.Items.Add("変03_携帯個人負担分");
            naiyou.Items.Add("変04_社宅変動費");
            naiyou.Items.Add("OIC食堂食事券購入");
            naiyou.Items.Add("低濃度オゾン発生装置購入代");
            //初期表示値
            naiyou.SelectedIndex = 0;

            //選択一覧データの枠
            SelectDisp.Columns.Add("社員番号", typeof(string));
            SelectDisp.Columns.Add("氏名", typeof(string));
            SelectDisp.Columns.Add("組織名", typeof(string));
            SelectDisp.Columns.Add("現場名", typeof(string));
            


            //エントリー一覧データの枠
            EntryDisp.Columns.Add("社員番号", typeof(string));
            EntryDisp.Columns.Add("氏名", typeof(string));
            EntryDisp.Columns.Add("現場名", typeof(string));
            EntryDisp.Columns.Add("組織名", typeof(string));
            //EntryDisp.Columns.Add("給与区分", typeof(string));
            EntryDisp.Columns.Add("項目", typeof(string));
            EntryDisp.Columns.Add("内容", typeof(string));
            EntryDisp.Columns.Add("金額", typeof(int));
            EntryDisp.Columns.Add("備考", typeof(string));



            GetData();
            Refine();
            GetEntryData();

            Com.InHistory("控除入力画面", "", "");
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
                        if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
                        {
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 in ('03_施設', '04_エンジ') order by 組織CD, 現場CD"; 
                        }
                        else if (Program.loginname == "金城　智之")
                        {
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 like '%%' order by 組織CD, 現場CD";
                        }
                        else
                        { 
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 like '%" + Program.loginbusyo + "%' order by 組織CD, 現場CD";
                        }
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
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

            foreach (DataRow row in dtrow)
            {
                DataRow nr = SelectDisp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["氏名"] = row["氏名"];
                nr["組織名"] = row["組織名"];
                nr["現場名"] = row["現場名"];
                
                SelectDisp.Rows.Add(nr);
            }

            dataGridView1.DataSource = SelectDisp;
        }

        //エントリーデータ取得
        private void GetEntryData()
        {
            EntryDisp.Clear();

            string sql = "";
            sql = " select a.項目, a.内容, a.社員番号, a.氏名, b.退職年月日, b.組織名, No, 金額, 備考 from dbo.固定控除 a left join dbo.s社員基本情報 b on a.社員番号 = b.社員番号 where 項目 in ('固定他１','固定他２','変動他１','変動他２') ";


            //DataTable EntryDisp = new DataTable();
            EntryDisp = Com.GetDB(sql);

            //foreach (DataRow row in Com.GetDB(sql).Rows)
            //{
            //    DataRow nr = EntryDisp.NewRow();
            //    nr["社員番号"] = row["社員番号"];
            //    nr["氏名"] = row["氏名"];
            //    nr["現場名"] = row["現場名"];
            //    nr["組織名"] = row["組織名"];
            //    //nr["給与区分"] = row["給与支給区分名"]; 
            //    nr["項目"] = row["項目"];
            //    nr["内容"] = row["内容"];
            //    nr["金額"] = row["金額"];
            //    nr["備考"] = row["備考"];
            //    EntryDisp.Rows.Add(nr);
            //}

            dataGridView2.DataSource = EntryDisp;
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

            naiyou.SelectedIndex = 0;
            //対象者のみに絞込
            DataRow[] targetDr = dt.Select("社員番号 = " + str.Substring(0, 8), "");

            foreach (DataRow row in targetDr)
            {
                number.Text = row["社員番号"].Equals(DBNull.Value) ? "" : row["社員番号"].ToString();　 //社員番号
                name.Text = row["氏名"].Equals(DBNull.Value) ? "" : row["氏名"].ToString();　　 //漢字氏名
                syokusyu.Text = row["組織名"].Equals(DBNull.Value) ? "" : row["組織名"].ToString(); //職種名
                genba.Text = row["現場名"].Equals(DBNull.Value) ? "" : row["現場名"].ToString();　　//現場名
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
            koumoku.SelectedIndex = 1;
            naiyou.SelectedIndex = 0;
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

  
                    Cmd.Parameters.Add(new SqlParameter("登録日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["登録日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("登録者", SqlDbType.VarChar));
                    Cmd.Parameters["登録者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD.ToString("yyyyMM");
                    Cmd.Parameters["社員番号"].Value = number.Text;
                    Cmd.Parameters["支払方法"].Value = naiyou.Text;
                    Cmd.Parameters["退職理由"].Value = koumoku.Text;
                    Cmd.Parameters["特記事項"].Value = etc.Text;

                    Cmd.Parameters["登録日時"].Value = DateTime.Now;
                    Cmd.Parameters["登録者"].Value = Program.loginname;

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
            SyousaiData(drv[2].ToString());
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
            button3.Enabled = false;

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

            //excel出力
            System.Diagnostics.Process.Start(exlName + ".xlsx");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button3.Enabled = true;
        }


    }
}
