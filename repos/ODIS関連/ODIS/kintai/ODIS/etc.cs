using ODIS.ODIS;
using Npgsql;
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
using Microsoft.VisualBasic.FileIO;

namespace ODIS.ODIS
{
    public partial class etc : Form
    {
        private TargetDays td = new TargetDays();

        public etc()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            #region チェックリストボックスの初期設定
            comboBox3.Items.Add("両方");
            comboBox3.Items.Add("男性");
            comboBox3.Items.Add("女性");
            comboBox3.SelectedIndex = 0;

            checkedListBox1.Items.Add("那覇");
            checkedListBox1.Items.Add("八重山");
            checkedListBox1.Items.Add("北部");
            checkedListBox1.Items.Add("本社");
            #endregion

            GetData();

            //トラックバー
            trackBar1.Minimum = 0;
            trackBar1.Maximum = 100;

            //初期値
            trackBar1.Value = 30;
            button1.Text = "月30人以上の入社or退職があった現場";

            // 描画される目盛りの刻みを設定
            trackBar1.TickFrequency = 10;

            // スライダーをキーボードやマウス、
            // PageUp,Downキーで動かした場合の移動量設定
            trackBar1.SmallChange = 1;
            trackBar1.LargeChange = 10;

            // 値が変更された際のイベントハンドらーを追加
            trackBar1.ValueChanged += new EventHandler(trackBar1_ValueChanged);

            if (Program.loginname == "喜屋武　大祐")
            {
                button2.Enabled = true;
            }

            Com.InHistory("テスト中", "", "");
        }

        void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            // TrackBarの値が変更されたらラベルに表示
            button1.Text = "月" + trackBar1.Value.ToString() + "人以上の入社or退職があった現場";
        }

        private DataTable GetData(string sql)
        {
            DataTable dt = new DataTable();
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = sql;
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

            return dt;
        }



        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = GetData("select * from dbo.s社員基本情報_期間指定('" + dateTimePicker1.Value.ToString("yyyy/MM/dd") + "') where 退職年月日 is null");
            dataGridView1.DataSource = dt;
            label1.Text = "上記日付時点の従業員数：" + dt.Rows.Count.ToString();
            Com.InHistory("当時の従業員一覧", "", "");
        }


        private void button6_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button6.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string str = "";
            //int ki = 46;
            //int start = 2017;
            //int end = 2018;
            //int ct = 12;
            //int ki = 47;
            //int start = 2018;
            //int end = 2019;
            //int ct = 13;
            int ki = 48;
            int start = 2019;
            int end = 2020;
            int ct = 14;

            for (int i = 0; i <= ct; i++)
            {
                str += " select ";
                str += " 期, ";
                str += " [在籍数(4月1日時点)] ";
                str += " , case when [在籍数(4月1日時点)] = 0 then 0 else [在籍年齢合計] / CONVERT(decimal, [在籍数(4月1日時点)]) end as 在籍平均年齢 ";
                str += " , case when [在籍数(4月1日時点)] = 0 then '-' else CONVERT(varchar(10), [在籍月合計] / [在籍数(4月1日時点)] / 12) + '年' + CONVERT(varchar(10), [在籍月合計] / [在籍数(4月1日時点)] % 12) + 'ヶ月' end as 在籍平均勤続年数 ";
                str += " , [入社数] ";
                str += " , case when [入社数] = 0 then 0 else [入社年齢合計] / CONVERT(decimal, [入社数]) end as 入社平均年齢 ";
                str += " , case when [在籍数(4月1日時点)] = 0 then 0 else [入社数]  * 100 / CONVERT(decimal, [在籍数(4月1日時点)]) end as [入職率(％)] ";
                str += " , [退職数] ";
                str += " , case when [退職数] = 0 then 0 else [退職年齢合計] / CONVERT(decimal, [退職数]) end as 退職平均年齢 ";
                str += " , case when [退職数] = 0 then '-' else CONVERT(varchar(10), [退職月合計] / [退職数] / 12) + '年' + CONVERT(varchar(10), [退職月合計] / [退職数] % 12) + 'ヶ月' end as 退職平均勤続年数 ";
                str += " , case when [在籍数(4月1日時点)] = 0 then 0 else [退職数] * 100 / CONVERT(decimal, [在籍数(4月1日時点)]) end as [離職率(％)] ";
                str += " from (select ";
                str += "' " + ki.ToString() + " (" + start.ToString() + "/04/01 ～ " + end.ToString() + "/03/31)' as 期, ";
                str += "sum(case when 入社年月日 <= '" + start.ToString() + "/04/01' and(退職年月日 >= '" + start.ToString() + "/04/01' or 退職年月日 is null) then 1 else 0 end) as [在籍数(4月1日時点)], ";
                str += "sum(case when 入社年月日 <= '" + start.ToString() + "/04/01' and(退職年月日 >= '" + start.ToString() + "/04/01' or 退職年月日 is null) then 年齢 else 0 end) as [在籍年齢合計], ";
                str += "sum(case when 入社年月日 <= '" + start.ToString() + "/04/01' and(退職年月日 >= '" + start.ToString() + "/04/01' or 退職年月日 is null) then 在籍月 else 0 end) as [在籍月合計], ";
                str += "sum(case when 入社年月日 between '" + start.ToString() + "/04/01' and '" + end.ToString() + "/03/31' then 1 else 0 end) as [入社数],  ";
                str += "sum(case when 入社年月日 between '" + start.ToString() + "/04/01' and '" + end.ToString() + "/03/31' then 年齢 else 0 end) as [入社年齢合計],  ";
                str += "sum(case when 退職年月日 between '" + start.ToString() + "/04/01' and '" + end.ToString() + "/03/31' then 1 else 0 end) as [退職数], ";
                str += "sum(case when 退職年月日 between '" + start.ToString() + "/04/01' and '" + end.ToString() + "/03/31' then 年齢 else 0 end) as [退職年齢合計], ";
                str += "sum(case when 退職年月日 between '" + start.ToString() + "/04/01' and '" + end.ToString() + "/03/31' then 在籍月 else 0 end) as [退職月合計] ";
                str += "from dbo.従業員情報_期間指定('" + end.ToString() + "/03/31') where 入社年月日 is not null ";

                //地区
                for (int ii = 0; ii < checkedListBox1.Items.Count; ii++)
                {
                    if (!checkedListBox1.GetItemChecked(ii))
                    {
                        str += " and 地区名 <> '" + checkedListBox1.Items[ii].ToString() + "'";
                    }
                }

                //職種
                for (int ii = 0; ii < checkedListBox2.Items.Count; ii++)
                {
                    if (!checkedListBox2.GetItemChecked(ii))
                    {
                        str += " and 部門名 <> '" + checkedListBox2.Items[ii].ToString() + "'";
                    }
                }

                //給与区分
                for (int ii = 0; ii < checkedListBox3.Items.Count; ii++)
                {
                    if (!checkedListBox3.GetItemChecked(ii))
                    {
                        str += " and 給与支給区分名 <> '" + checkedListBox3.Items[ii].ToString() + "'";
                    }
                }

                //性別
                if (comboBox3.SelectedItem.ToString() == "男性")
                {
                    str += " and 性別区分 = '1' ";
                }
                else if (comboBox3.SelectedItem.ToString() == "女性")
                {
                    str += " and 性別区分 = '2' ";
                }
                else
                {

                }


                str += " ) temp ";
                if (i != ct) str += "union all ";

                ki--; start--; end--;
            }

            DataTable dt = new DataTable();
            dt = Com.GetDB(str);
            dataGridView1.DataSource = dt;

            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[2].Width = 130;
            dataGridView1.Columns[3].Width = 130;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 120;
            dataGridView1.Columns[6].Width = 100;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].Width = 130;
            dataGridView1.Columns[9].Width = 130;
            dataGridView1.Columns[10].Width = 100;

            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.Snow; //売上
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.Snow; //売上
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.Snow; //売上
            dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.Linen; //売上
            dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.Linen; //売上
            dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.Linen; //売上
            dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.Linen; //売上

            dataGridView1.Columns[0].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[2].DefaultCellStyle.Format = "0.00\'才\'";
            dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[5].DefaultCellStyle.Format = "0.00\'才\'";
            dataGridView1.Columns[6].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "0.00\'才\'";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "0.00\'%\'";

            //dataGridView1.Columns[9].DefaultCellStyle.Format = "0.00\'%\'";

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button6.Enabled = true;
        }

        //データ取得
        private void GetData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable dtSikyuu = new DataTable();
            DataTable dtBumon = new DataTable();
            //DataTable dtYakusyoku = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {

                        //Cmd.CommandText = "select distinct 部門CD, 部門名 from dbo.accessNew order by 部門CD";
                        Cmd.CommandText = "select distinct 部門名 from dbo.accessNew order by 部門名";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dtBumon);

                        Cmd.CommandText = "select distinct 給与支給区分, 給与支給名称 from dbo.accessNew order by 給与支給区分";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dtSikyuu);

                        //Cmd.CommandText = "select distinct 役職CD, 役職名 from dbo.accessNew order by 役職CD";
                        //da = new SqlDataAdapter(Cmd);
                        //da.Fill(dtYakusyoku);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            foreach (DataRow row in dtBumon.Rows)
            {
                checkedListBox2.Items.Add(row["部門名"]);
            }

            foreach (DataRow row in dtSikyuu.Rows)
            {
                checkedListBox3.Items.Add(row["給与支給名称"]);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }

        }

        private void label22_Click(object sender, EventArgs e)
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
        }

        private void label24_Click(object sender, EventArgs e)
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;
            string str = "";

            str += "select 現場CD, 現場名, '退職' as 項目, substring(退職年月日, 1, 7) as 年月, count(*) as 人数 from dbo.社員基本情報 group by 現場CD, 現場名, substring(退職年月日, 1, 7) having count(*) > " + trackBar1.Value + " and substring(退職年月日, 1, 7) is not null ";
            str += " union all ";
            str += " select 現場CD, 現場名, '入社' as 項目, substring(入社年月日, 1, 7) as 年月, count(*) as 人数 from dbo.社員基本情報 group by 現場CD, 現場名, substring(入社年月日, 1, 7) having count(*) > " + trackBar1.Value + " and substring(入社年月日, 1, 7) is not null";

            DataTable dt = new DataTable();
            dt = Com.GetDB(str);
            dataGridView1.DataSource = dt;

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            SendMail sm = new SendMail("test", "test", "test");
        }

        


        //private void button8_Click(object sender, EventArgs e)
        //{
        //    //TextFieldParser parser = new TextFieldParser(@"C:\temp\勤怠テンプレート.csv", Encoding.GetEncoding("Shift_JIS"));
        //    TextFieldParser parser = new TextFieldParser(@textBox1.Text, Encoding.GetEncoding("Shift_JIS"));
        //    parser.TextFieldType = FieldType.Delimited;
        //    parser.SetDelimiters(","); // 区切り文字はコンマ

        //    // データをすべてクリア
        //    dataGridView1.Rows.Clear();

        //    //dataGridView1.ColumnCount = 6;
        //    //dataGridView1.Columns[0].HeaderText = "処理年";
        //    //dataGridView1.Columns[1].HeaderText = "処理月";
        //    //dataGridView1.Columns[2].HeaderText = "社員番号";
        //    //dataGridView1.Columns[3].HeaderText = "姓";
        //    //dataGridView1.Columns[4].HeaderText = "名";
        //    //dataGridView1.Columns[5].HeaderText = "名";

        //    bool flg = true;

        //    while (!parser.EndOfData)
        //    {
        //        string[] row = parser.ReadFields(); // 1行読み込み

        //        //列追加
        //        if (flg)
        //        {
        //            dataGridView1.ColumnCount = row.Length;
        //            for (int i = 0; i < row.Length; i++)
        //            {
        //                dataGridView1.Columns[i].HeaderText = row[i];
        //            }
        //            flg = false;
        //        }
        //        else
        //        {
        //            // 読み込んだデータ(1行をDataGridViewに表示する)
        //            dataGridView1.Rows.Add(row);
        //        }
        //    }
        //}

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            TextBox TargetTextBox = sender as TextBox;

            if (TargetTextBox == null)
            {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグ中のファイルやディレクトリの取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop);

            //ファイルがドラッグされている場合、
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 配列分ループ
                foreach (string sTemp in sFileName)
                {
                    // ファイルパスかチェック
                    if (File.Exists(sTemp) == false)
                    {
                        // ファイルパス以外なので何もしない
                        return;
                    }
                    else
                    {
                        break;
                    }
                }

                // カーソルを[+]へ変更する
                // ここでEffectを変更しないと、以降のイベント（Drop）は発生しない
                e.Effect = DragDropEffects.Copy;
            }
        }

       

        private void button7_Click_1(object sender, EventArgs e)
        {
            ////Form2に送るテキスト
            //string sendText = idouday.Value.ToString("yyyy/MM/dd");

            ////Form2から送られてきたテキストを受け取る。
            //string[] receiveText = SelectEmpNyu.ShowMiniForm(sendText);　//Form2を開く

            //if (receiveText == null) return;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            string ym = "";

            //CSV読み込み
            // CSVファイルの読み込み
            //string filePath = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\keisu202107.csv";
            // StreamReaderクラスをインスタンス化
            StreamReader reader = new StreamReader(txtcsv.Text, Encoding.GetEncoding("shift_jis"));

            int i=0;
            // 最後まで読み込む
            while (reader.Peek() >= 0)
            {
                string[] cols = reader.ReadLine().Split(',');
                if (i==1) ym = cols[2];
                i++;
            }
            reader.Close();

            DialogResult result = MessageBox.Show("対象年月全データdelete後にinsertしてもよいですか？" + Environment.NewLine + "対象年月:" + ym,
                        "警告",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) return;

            string filePath = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\imp\" + ym + "_" + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒") + ".csv";
            

            //コピーするー。
            File.Copy(txtcsv.Text, filePath);

            DataTable dt = new DataTable();
            //delete
            dt = Com.GetDB("delete from genbakeisuu where 処理年月 = '" + ym + "'");
            //insert
            dt = Com.GetDB("BULK INSERT genbakeisuu FROM '" + filePath + "' WITH ( FIELDTERMINATOR = ',', ROWTERMINATOR = '\n', FIRSTROW = 2 )");
        }

        private void txtcsv_DragDrop(object sender, DragEventArgs e)
        {

            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            TextBox TargetTextBox = sender as TextBox;

            if (TargetTextBox == null)
            {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定
        }

        private void txtcsv_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグ中のファイルやディレクトリの取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop);

            //ファイルがドラッグされている場合、
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 配列分ループ
                foreach (string sTemp in sFileName)
                {
                    // ファイルパスかチェック
                    if (File.Exists(sTemp) == false)
                    {
                        // ファイルパス以外なので何もしない
                        return;
                    }
                    else
                    {
                        break;
                    }
                }

                // カーソルを[+]へ変更する
                // ここでEffectを変更しないと、以降のイベント（Drop）は発生しない
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //出力
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.k給与改定一覧");

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\給与改定通知.xlsx");

            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                }
            }

            //string sheetname = ym.ToString("yyyyMM");

            XLSheet ws = c1XLBook1.Sheets["通知フォーム"];

            for (int i = 1; i <= rows; i++)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = i.ToString();   // クローンをリネーム
                newSheet[1, 7].Value = i;      // 値の変更
                c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
            }

            c1XLBook1.Sheets.Remove("通知フォーム");


            string localPass = @"C:\ODIS\TSUUCHI\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");

            //if (excel)
            //{
            //    System.Diagnostics.Process.Start(exlName + ".xlsx");
            //}
            //else
            //{
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


                //excel出力
                //System.Diagnostics.Process.Start(@"c:\temp\test2.xlsx");
                //PDF出力
                System.Diagnostics.Process.Start(exlName + ".pdf");
            //}
        }
    }
}
