using C1.C1Excel;
using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class HendouKoujo : Form
    {
        //元データ
        DataTable dt = new DataTable();

        //選択一覧データ
        DataTable SelectDisp = new DataTable();

        //エントリー一覧データ
        DataTable EntryDisp = new DataTable();

        //退職入力の制限
        TargetDays td = new TargetDays();

        //総務のみ表示する内容
        //string soumuonly = "";

        public HendouKoujo()
        {
            if (Convert.ToInt16(Program.access) == 1)
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("変動控除入力権限無", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            //フォントサイズの変更
            dataGridView2.Font = new Font(dataGridView2.Font.Name, 10);

            //グリッドビューのコピーで
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //値の変更不可
            dataGridView2.ReadOnly = true;


            //koumoku.Items.Add("固定他１");
            //koumoku.Items.Add("固定他２");
            //koumoku.Items.Add("臨時手当");
            //koumoku.Items.Add("臨作業手当");
            koumoku.Items.Add("変動他１");
            koumoku.Items.Add("変動他２");
            //koumoku.Items.Add("前払(+)");
            //koumoku.Items.Add("前払(-)");

            //初期表示値
            koumoku.SelectedIndex = 1;

            naiyou.Items.Add("変01_作業靴代(自己負担分)");
            naiyou.Items.Add("変02_寮変動費");
            naiyou.Items.Add("変03_携帯個人負担分");
            naiyou.Items.Add("変04_社宅変動費");
            naiyou.Items.Add("変05_OIC食堂食事券購入");
            naiyou.Items.Add("変06_濃度オゾン発生装置購入代");
            naiyou.Items.Add("変07_食堂利用代");
            naiyou.Items.Add("変08_モノレール定期券代");
            naiyou.Items.Add("その他");

            //初期表示値
            naiyou.SelectedIndex = 0;

            //選択一覧データの枠
            SelectDisp.Columns.Add("社員番号", typeof(string));
            SelectDisp.Columns.Add("氏名", typeof(string));
            SelectDisp.Columns.Add("組織名", typeof(string));
            SelectDisp.Columns.Add("現場名", typeof(string));
            SelectDisp.Columns.Add("給与支給区分名", typeof(string));

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

            //自分が登録した分チェックボックスにデフォルトチェック
            //checkBox1.Checked = false;

            //新規、変更の分だけ表示
            //checkBox2.Checked = false;

            GetData();
            Refine();
            GetEntryData();

            label1.Text = td.StartYMD.AddMonths(1).ToString("yyyy年MM月支給分");

            Com.InHistory("変動控除入力画面", "", "");
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

                        DataTable dtcb2 = new DataTable();
                        string sql2 = "";

                        DataTable dtcb3 = new DataTable();
                        string sql3 = "";

                        string y = td.StartYMD.ToString("yyyy");
                        string m = td.StartYMD.ToString("MM");

                        if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
                        {
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 in ('03_施設', '04_エンジ') order by 組織CD, 現場CD";
                            sql = " select distinct 担当区分 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 in ('03_施設', '04_エンジ') order by 担当区分";
                            sql2 = " select distinct 項目 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 in ('03_施設', '04_エンジ') order by 項目 ";
                            sql3 = " select distinct 内容 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 in ('03_施設', '04_エンジ') order by 内容 ";

                        }
                        else if (Program.loginname == "金城　智之" || Program.loginname == "喜屋武　大祐")
                        {
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 like '%%' order by 組織CD, 現場CD";
                            sql = " select distinct 担当区分 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%%' order by 担当区分";
                            sql2 = " select distinct 項目 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%%' ";
                            sql3 = " select distinct 内容 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%%' order by 内容 ";
                        }
                        else
                        {
                            Cmd.CommandText = " select * from dbo.s社員検索 where 担当区分 like '%" + Program.loginbusyo + "%' order by 組織CD, 現場CD";
                            sql = " select distinct 担当区分 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%" + Program.loginbusyo + "%' order by 担当区分";
                            sql2 = " select distinct 項目 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%" + Program.loginbusyo + "%' order by 項目 ";
                            sql3 = " select distinct 内容 from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%" + Program.loginbusyo + "%' order by 内容 ";
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

                        //チェックボックス
                        dtcb2 = Com.GetDB(sql2);

                        foreach (DataRow row in dtcb2.Rows)
                        {
                            checkedListBox2.Items.Add(row["項目"]);
                        }

                        for (int i = 0; i < checkedListBox2.Items.Count; i++)
                        {
                            checkedListBox2.SetItemChecked(i, true);
                        }

                        //チェックボックス
                        dtcb3 = Com.GetDB(sql3);

                        foreach (DataRow row in dtcb3.Rows)
                        {
                            checkedListBox3.Items.Add(row["内容"]);
                        }

                        for (int i = 0; i < checkedListBox3.Items.Count; i++)
                        {
                            checkedListBox3.SetItemChecked(i, true);
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
                nr["給与支給区分名"] = row["給与支給区分名"];
                SelectDisp.Rows.Add(nr);
            }

            dataGridView1.DataSource = SelectDisp;
        }

        //エントリーデータ取得
        private void GetEntryData()
        {
            EntryDisp.Clear();

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

            //項目
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    if (!checkedListBox2.GetItemChecked(i))
                    {
                        result += " and 項目 <> '" + checkedListBox2.Items[i].ToString() + "'";
                    }
                }
            }

            //内容
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    if (!checkedListBox3.GetItemChecked(i))
                    {
                        result += " and 内容 <> '" + checkedListBox3.Items[i].ToString() + "'";
                    }
                }
            }


            //自分が更新
            if (checkBox1.Checked)
            {
                result += " and (更新者 = '" + Program.loginname + "' or 登録者 = '" + Program.loginname + "') ";
            }

            //if (checkBox2.Checked)
            //{
            //    result += " and 状況項目 <> ''";
            //}

            string sql = "";
            string y = td.StartYMD.ToString("yyyy");
            string m = td.StartYMD.ToString("MM");



            if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
            {
                sql = " select * from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 in ('03_施設', '04_エンジ') " + result + " order by 内容, 組織CD, 現場CD";
            }
            else if (Program.loginname == "金城　智之")
            {
                sql = " select * from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%%' " + result + " order by 内容, 組織CD, 現場CD";
            }
            else
            {
                sql = " select * from dbo.h変動控除データ取得_前月比('" + y + "','" + m + "') where 担当区分 like '%" + Program.loginbusyo + "%' " + result + " order by 内容, 組織CD, 現場CD";
            }

            EntryDisp = Com.GetDB(sql);

            dataGridView2.DataSource = EntryDisp;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            //SyousaiData(drv[0].ToString());

            number.Text = drv[0].ToString();  //社員番号
            name.Text = drv[1].ToString();   //漢字氏名
            syokusyu.Text = drv[2].ToString(); //職種名
            genba.Text = drv[3].ToString();  //現場名
            kubun.Text = drv[4].ToString();  //給与支給区分名

            //koumoku.SelectedItem = "";  //項目
            //naiyou.Text = "";  //内容
            //kingaku.Text = "";  //金額
            //etc.Text = "";  //備考
            kanriNo.Text = "";　　//管理No
        }

        private void SyousaiData(string str)
        {

            naiyou.SelectedIndex = 0;
            //対象者のみに絞込
            DataRow[] targetDr = EntryDisp.Select("管理No = " + str, "");

            foreach (DataRow row in targetDr)
            {
                number.Text = row["社員番号"].Equals(DBNull.Value) ? "" : row["社員番号"].ToString();　 //社員番号
                name.Text = row["氏名"].Equals(DBNull.Value) ? "" : row["氏名"].ToString();　　 //漢字氏名
                syokusyu.Text = row["組織名"].Equals(DBNull.Value) ? "" : row["組織名"].ToString(); //職種名
                genba.Text = row["現場名"].Equals(DBNull.Value) ? "" : row["現場名"].ToString();　　//現場名

                kubun.Text = row["給与支給区分名"].Equals(DBNull.Value) ? "" : row["給与支給区分名"].ToString();　　//給与支給区分名
                koumoku.SelectedItem = row["項目"].Equals(DBNull.Value) ? "" : row["項目"].ToString();　　//項目
                naiyou.Text = row["内容"].Equals(DBNull.Value) ? "" : row["内容"].ToString(); //内容
                kingaku.Text = row["金額"].Equals(DBNull.Value) ? "" : row["金額"].ToString(); //金額
                etc.Text = row["備考"].Equals(DBNull.Value) ? "" : row["備考"].ToString();  //備考
                kanriNo.Text = row["管理No"].Equals(DBNull.Value) ? "" : row["管理No"].ToString(); //管理No

                //状況項目が新規の場合だけが、削除ボタンを有効にする。
                //if (row["状況項目"].ToString() == "新規")
                //{
                //    button2.Visible = true;
                //}
                //else
                //{
                //    button2.Visible = false;
                //}
            }
        }

        //登録ボタンクリック
        private void button1_Click(object sender, EventArgs e)
        {
            UpdateInsert();
        }

        private void UpdateInsert()
        {
            //登録有無分岐必要
            DataInsert();

            //GetData();
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
            if (kanriNo.Text != "")
            {
                cmd = "[dbo].[UpdateRingiData]";
            }
            else
            {
                cmd = "[dbo].[InsertRingiData]";

            }

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = cmd;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名", SqlDbType.VarChar));
                    Cmd.Parameters["氏名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("項目", SqlDbType.VarChar));
                    Cmd.Parameters["項目"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("内容", SqlDbType.VarChar));
                    Cmd.Parameters["内容"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("金額", SqlDbType.VarChar));
                    Cmd.Parameters["金額"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                    Cmd.Parameters["備考"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("管理No", SqlDbType.VarChar));
                    Cmd.Parameters["管理No"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("登録者", SqlDbType.VarChar));
                    Cmd.Parameters["登録者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("登録日時", SqlDbType.DateTime));
                    Cmd.Parameters["登録日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar));
                    Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.DateTime));
                    Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = number.Text;
                    Cmd.Parameters["氏名"].Value = name.Text;
                    Cmd.Parameters["項目"].Value = koumoku.SelectedItem.ToString();
                    Cmd.Parameters["内容"].Value = naiyou.Text;
                    Cmd.Parameters["金額"].Value = kingaku.Value.ToString();
                    Cmd.Parameters["備考"].Value = etc.Text;
                    Cmd.Parameters["管理No"].Value = kanriNo.Text;
                    Cmd.Parameters["登録者"].Value = Program.loginname;
                    Cmd.Parameters["登録日時"].Value = DateTime.Now;
                    Cmd.Parameters["更新者"].Value = Program.loginname;
                    Cmd.Parameters["更新日時"].Value = DateTime.Now;

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
            if (kanriNo.Text == "") return;

            DialogResult result = MessageBox.Show("削除してよろしいですか？",
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
                    Cmd.CommandText = "delete from dbo.固定控除 WHERE 管理No = '" + kanriNo.Text + "'";
                    using (dr = Cmd.ExecuteReader())
                    {
                        //TODO
                    }
                }
            }

            //GetData();
            Refine();
            GetEntryData();

            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //強制的にチェック
            //checkBox2.Checked = true;

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button1.Enabled = false;

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\61_臨時手当申請.xlsx");

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

            ws[69, 0].Value = DateTime.Now.ToString("yyyy年MM月dd日");
            ws[69, 2].Value = td.StartYMD.ToString("yyyy年M月 ") + td.StartYMD.AddMonths(1).ToString(" (M月給与分)");
            ws[69, 4].Value = Program.loginname;

            int ct = rows / 10 + 1;
            int ct2 = 1;
            for (int i = 1; i <= ct; i++)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = i.ToString();   // クローンをリネーム
                newSheet[1, 6].Value = ct2.ToString();      // 値の変更
                c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
                ct2 = ct2 + 10;
            }

            c1XLBook1.Sheets.Remove("原票");

            string localPass = @"C:\ODIS\KKoujo\";
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

            //exccelファイル削除
            //FileInfo file = new FileInfo(exlName + ".xlsx");
            //file.Delete();

            //PDFopen
            //System.Diagnostics.Process.Start(exlName + ".xlsx");
            System.Diagnostics.Process.Start(@"C:\ODIS\KKoujo\");
            System.Diagnostics.Process.Start(exlName + ".pdf");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void kanriNo_TextChanged(object sender, EventArgs e)
        {
            if (kanriNo.Text == "")
            {
                button1.Text = "新規登録";
                button2.Visible = false;
                button5.Visible = false;
            }
            else
            {
                button1.Text = "更新";
                button2.Visible = true;
                button5.Visible = true;
            }
        }

        private void checkedListBox1_MouseCaptureChanged(object sender, EventArgs e)
        {
            //GetEntryData();
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

            GetEntryData();
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
        }

        //private void button4_Click(object sender, EventArgs e)
        //{
        //    GetEntryData();
        //}

        private void button5_Click(object sender, EventArgs e)
        {
            kanriNo.Text = "";
            UpdateInsert();
        }

        private void checkedListBox2_ItemCheck(object sender, ItemCheckEventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
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

            GetEntryData();
        }

        private void label8_Click(object sender, EventArgs e)
        {
            if (checkedListBox3.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, false);
                }
            }

            GetEntryData();
        }

        private void dataGridView2_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 16)
            {
                //セルの値により、背景色を変更する
                //if (e.Value.ToString() != "" && e.Value.ToString() == "変更")
                if (e.Value.ToString() == "金額変更")
                {
                    dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.AliceBlue;
                }
                else if (e.Value.ToString() == "新規")
                {
                    dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Cornsilk;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            GetEntryData();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            GetEntryData();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetEntryData();
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetEntryData();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetEntryData();
        }
    }
}
