using C1.C1Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Emp : Form
    {
        //全データ格納テーブル
        private DataTable dt = new DataTable();

        //濱川用 検索条件
        //private string hama = " where 在籍区分 = '1' ";
        private string hama = " where 在籍区分 = '1' ";

        //固定一覧
        private DataTable dtkotei = new DataTable();
        //private string taisyoku = "";

        //null対応
        private DateTime zerodt = new System.DateTime(2022, 12, 31, 0, 0, 0, 0);

        public Emp()
        {
            InitializeComponent();

            //フォームのKeyPreviewを切り替える
            //this.KeyPreview = !this.KeyPreview;

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvkintai.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvkouza.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvmeisai.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            shikakudgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView9.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //濱さん用
            dataGridView4.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            dataGridView2.RowHeadersVisible = false;

            //行ヘッダを非表示
            shikakudgv.RowHeadersVisible = false;

            // 選択モードを行単位での選択のみにする
            shikakudgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;


            //データ取得
            GetData("");

            //#region チェックリストボックスの初期設定
            //checkedListBox1.Items.Add("那覇");
            //checkedListBox1.Items.Add("八重山");
            //checkedListBox1.Items.Add("北部");
            //checkedListBox1.Items.Add("本社");
            //checkedListBox1.Items.Add("広域");
            //checkedListBox1.Items.Add("宮古島");
            //checkedListBox1.Items.Add("久米島");
            //#endregion


            #region コンボボックスの初期設定
            //年齢
            for (int i = 10; i <= 90; i++) comboBox1.Items.Add(i);
            for (int i = 10; i <= 90; i++) comboBox2.Items.Add(i);

            //在籍数
            for (int i = 0; i <= 50; i++) comboBox3.Items.Add(i);
            for (int i = 0; i <= 50; i++) comboBox4.Items.Add(i);

            //誕生月
            for (int i = 1; i <= 12; i++) comboBox5.Items.Add(i);
            #endregion

            //出勤簿 TODO 
            comboBox8.Items.Add("2025/01");
            comboBox8.Items.Add("2025/02");
            comboBox8.Items.Add("2025/03");
            comboBox8.Items.Add("2025/04");
            comboBox8.Items.Add("2025/05");
            comboBox8.Items.Add("2025/06");
            comboBox8.Items.Add("2025/07");
            comboBox8.Items.Add("2025/08");
            comboBox8.Items.Add("2025/09");
            comboBox8.Items.Add("2025/10");
            comboBox8.Items.Add("2025/11");
            comboBox8.Items.Add("2025/12");
            comboBox8.Items.Add("2026/01");
            comboBox8.Items.Add("2026/02");
            comboBox8.Items.Add("2026/03");

            wareki.Items.Add("");

            for (int i = 20; i < 31; i++)
            {
                wareki.Items.Add("平成" + i + "年");
            }

            wareki.Items.Add("令和元年(平成31年)");

            for (int i = 2; i < 30; i++)
            {
                wareki.Items.Add("令和" + i + "年(平成" + (i + 30) + "年)");
            }

            //労働条件の設定
            //TODO 運用がうまくいけば空白はなくなる！
            koyoukubun.Items.Add("");
            koyoukubun.Items.Add("1 期間の定めあり");
            koyoukubun.Items.Add("2 期間の定めなし");

            koushinkubun.Items.Add("");
            koushinkubun.Items.Add("1 自動更新");
            koushinkubun.Items.Add("2 更新する場合があり得る");
            koushinkubun.Items.Add("3 契約の更新はしない");

            //休日勤務
            kyuusyustu.Items.Add("");
            kyuusyustu.Items.Add("1 有");
            kyuusyustu.Items.Add("2 無");

            //時間外労働
            zikangairoudou.Items.Add("");
            zikangairoudou.Items.Add("1 有");
            zikangairoudou.Items.Add("2 無");

            //夜間勤務
            yakankinmu.Items.Add("");
            yakankinmu.Items.Add("1 有");
            yakankinmu.Items.Add("2 無");

            //定年
            teinen.Items.Add("");
            teinen.Items.Add("1 該当無");
            teinen.Items.Add("2 満65才");

            //賞与
            syouyo.Items.Add("");
            syouyo.Items.Add("1 有");
            syouyo.Items.Add("2 無");

            //退職金
            taisyokukin.Items.Add("");
            taisyokukin.Items.Add("1 有");
            taisyokukin.Items.Add("2 無");

            //休日回数 
            kyuujitsukaisuu.Items.Add("");
            kyuujitsukaisuu.Items.Add("1 年間107日　月変形週労40時間  ※2月が29日の暦日の場合は108日");
            kyuujitsukaisuu.Items.Add("2 1ヶ月につき  4日～10日  (週5以上勤務)"); //週5
            kyuujitsukaisuu.Items.Add("3 1ヶ月につき 12日～15日  (週4勤務)"); //週4
            kyuujitsukaisuu.Items.Add("4 1ヶ月につき 16日～20日  (週3勤務)"); //週3
            kyuujitsukaisuu.Items.Add("5 1ヶ月につき 20日～25日  (週2勤務)"); //週2
            kyuujitsukaisuu.Items.Add("6 1ヶ月につき 23日～27日  (週1勤務)"); //週1

            //初期値設定
            Clear();
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from accessNew where 在籍区分 <> '9' order by 担当区分 ";
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

        private void Clear()
        {
            //文字列で絞込
            textBox1.Text = "";
            textBox2.Text = "";

            //文字列指定
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;

            //数字で絞込
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;

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

            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                checkedListBox4.SetItemChecked(i, true);
            }

            comboBox1.SelectedIndex = 10;
            comboBox2.SelectedIndex = 11;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 1;
            comboBox5.SelectedIndex = 0;

            //生年月日
            dateTimePicker1.Value = DateTime.Today.AddYears(-80);
            dateTimePicker2.Value = DateTime.Today.AddYears(-10);

            //入社年月日
            dateTimePicker3.Value = new DateTime(DateTime.Today.AddMonths(-1).Year, DateTime.Today.AddMonths(-1).Month, 1);
            dateTimePicker4.Value = DateTime.Today;

            //入社年月日
            dateTimePicker5.Value = new DateTime(DateTime.Today.AddMonths(-1).Year, DateTime.Today.AddMonths(-1).Month, 1);
            dateTimePicker6.Value = DateTime.Today;

            #region 非表示対応

            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;
            label29.Visible = false;

            dateTimePicker3.Visible = false;
            dateTimePicker4.Visible = false;
            label30.Visible = false;

            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            label31.Visible = false;

            comboBox1.Visible = false;
            comboBox2.Visible = false;
            label25.Visible = false;
            label26.Visible = false;

            comboBox3.Visible = false;
            comboBox4.Visible = false;
            label27.Visible = false;
            label28.Visible = false;

            comboBox5.Visible = false;

            #endregion

            //退職者含める
            checkBox1.Checked = true;
            checkBox2.Checked = false;
            checkBox6.Checked = false;
        }

        //データ取得
        private void GetData(string str)
        {
            //部門設定
            SetTiku();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable dtSikyuu = new DataTable();
            DataTable dtBumon = new DataTable();
            DataTable dtYakusyoku = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        if (dt.Rows.Count > 0)
                        {

                        }
                        else
                        { 
                            Cmd.CommandText = "select * from dbo.accessNew";
                            Cmd.CommandTimeout = 600;
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dt);
                        }

                        if (str == "")
                        {
                            //Cmd.CommandText = "select distinct 部門CD, 部門名 from dbo.accessNew where 在籍区分 <> '9' order by 部門CD";
                            Cmd.CommandText = "select distinct 担当事務 from dbo.accessNew where 在籍区分 <> '9' order by 担当事務";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtBumon);

                            Cmd.CommandText = "select distinct 給与支給区分, 給与支給区分名 from dbo.社員基本情報 where 在籍区分 <> '9' order by 給与支給区分";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtSikyuu);

                            Cmd.CommandText = "select distinct 役職CD, 役職名 from dbo.社員基本情報 where 在籍区分 <> '9' order by 役職CD";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtYakusyoku);
                        }
                        else
                        {
                            //Cmd.CommandText = "select distinct 部門CD, 部門名 from dbo.accessNew order by 部門CD";
                            Cmd.CommandText = "select distinct 担当事務 from dbo.accessNew order by 担当事務";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtBumon);

                            Cmd.CommandText = "select distinct 給与支給区分, 給与支給区分名 from dbo.社員基本情報 order by 給与支給区分";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtSikyuu);

                            Cmd.CommandText = "select distinct 役職CD, 役職名 from dbo.社員基本情報 order by 役職CD";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dtYakusyoku);
                        }
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
                checkedListBox2.Items.Add(row["担当事務"]);
            }

            foreach (DataRow row in dtSikyuu.Rows)
            {
                checkedListBox3.Items.Add(row["給与支給区分名"]);
            }

            foreach (DataRow row in dtYakusyoku.Rows)
            {
                checkedListBox4.Items.Add(row["役職名"]);
            }

            comboBox6.Items.Add("全対象");
            comboBox7.Items.Add("全対象");
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "生年月日絞込用") break;
                comboBox6.Items.Add(col.ColumnName);
                comboBox7.Items.Add(col.ColumnName);
            }

        }

        private void DataView()
        {

            #region チェックボックスが１つ以上入っているかチェック
            string errStr = "";
            if (checkedListBox1.CheckedItems.Count == 0) errStr += "【地区】に一個以上のチェックが必要です。\n";
            if (checkedListBox2.CheckedItems.Count == 0) errStr += "【部門】に一個以上のチェックが必要です。\n";
            if (checkedListBox3.CheckedItems.Count == 0) errStr += "【給与区分】に一個以上のチェックが必要です。\n";
            if (checkedListBox4.CheckedItems.Count == 0) errStr += "【役職】に一個以上のチェックが必要です。\n";
            #endregion

            #region 期間チェック
            if (dateTimePicker1.Value > dateTimePicker2.Value) errStr += "「生年月日」が「開始日 > 終了日」になってます。\n";
            if (dateTimePicker3.Value > dateTimePicker4.Value) errStr += "「入社年月日」が「開始日 > 終了日」になってます。\n";
            if (dateTimePicker5.Value > dateTimePicker6.Value) errStr += "「退職年月日」が「開始日 > 終了日」になってます。\n";
            if (Convert.ToInt16(comboBox1.SelectedItem) >= Convert.ToInt16(comboBox2.SelectedItem)) errStr += "「年齢」が「開始年 >= 終了年」になってます。\n";
            if (Convert.ToInt16(comboBox3.SelectedItem) >= Convert.ToInt16(comboBox4.SelectedItem)) errStr += "「在籍年数」が「開始日 >= 終了日」になってます。\n";
            #endregion

            if (errStr != "")
            {
                MessageBox.Show(errStr);
                return;
            }


            DataTable Disp = new DataTable();

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string conditions = "";

            string result = "";
            if (ar[0] != "")
            {
                conditions = "【含】" + textBox1.Text;

                foreach (string s in ar)
                {
                    if (comboBox6.SelectedIndex == 0)
                    {
                        result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                    else
                    {
                        result += " and (" + comboBox6.SelectedItem.ToString() + " like '%" + s + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Com.isOneByteChar(s) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }
            }

            //test 除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                conditions += "　【除】" + textBox2.Text;
                foreach (string s in ar2)
                {
                    if (comboBox7.SelectedIndex == 0)
                    {
                        result += " and reskey not like '%" + s + "%' ";
                        //result += " and (reskey not like '%" + s + "%' or reskey not like '%" + Com.isOneByteChar(s) + "%' or reskey not like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey not like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey not like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey not like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                    else
                    {
                        result += " and (" + comboBox7.SelectedItem.ToString() + " not like '%" + s + "%' or " + comboBox7.SelectedItem.ToString() + " not like '%" + Com.isOneByteChar(s) + "%' or " + comboBox7.SelectedItem.ToString() + " not like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or " + comboBox7.SelectedItem.ToString() + " not like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or " + comboBox7.SelectedItem.ToString() + " not like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or " + comboBox7.SelectedItem.ToString() + " not like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }
            }

            #region チェックボックスの絞込文字列を設定

            //在籍者のみ
            if (!checkBox6.Checked)
            {
                result += " and 在籍区分 = '1'";
            }
            else
            {
                conditions += "　【含】退職";
            }

            //地区

            bool flg = false;

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    result += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    flg = true;
                }
            }

            if (flg)
            {
                conditions += "　【絞】地区";
                flg = false;
            }


            //部門
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 担当事務 <> '" + checkedListBox2.Items[i].ToString() + "'";
                    flg = true;
                }
            }

            if (flg)
            {
                conditions += "　【絞】部門";
                flg = false;
            }

            //給与区分
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    result += " and 給与支給名称 <> '" + checkedListBox3.Items[i].ToString() + "'";
                    flg = true;
                }
            }

            if (flg)
            {
                conditions += "　【絞】給与";
                flg = false;
            }

            //役職
            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                if (!checkedListBox4.GetItemChecked(i))
                {
                    result += " and 役職名 <> '" + checkedListBox4.Items[i].ToString() + "'";
                    flg = true;
                }
            }

            if (flg)
            {
                conditions += "　【絞】役職";
                flg = false;
            }

            //生年月日
            if (checkBox3.Checked)
            {
                string sDate = dateTimePicker1.Value.ToString("yyyyMMdd");
                string eDate = dateTimePicker2.Value.ToString("yyyyMMdd");
                result += " and 生年月日絞込用 >= " + sDate + " and 生年月日絞込用 <= " + eDate;
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】生年";
                flg = false;
            }

            //入社年月日
            if (checkBox4.Checked)
            {
                string sDate = dateTimePicker3.Value.ToString("yyyyMMdd");
                string eDate = dateTimePicker4.Value.ToString("yyyyMMdd");
                result += " and 入社年月日絞込用 >= " + sDate + " and 入社年月日絞込用 <= " + eDate;
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】入社";
                flg = false;
            }

            //退職年月日
            if (checkBox5.Checked)
            {
                string sDate = dateTimePicker5.Value.ToString("yyyyMMdd");
                string eDate = dateTimePicker6.Value.ToString("yyyyMMdd");
                result += " and 退職年月日絞込用 >= " + sDate + " and 退職年月日絞込用 <= " + eDate;
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】退職";
                flg = false;
            }

            //年齢
            if (checkBox7.Checked)
            {

                result += "and 年齢 >= " + comboBox1.SelectedItem.ToString() + " and 年齢 < " + comboBox2.SelectedItem.ToString();
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】年齢";
                flg = false;
            }

            //在籍年数
            if (checkBox8.Checked)
            {

                result += "and 在籍月 >= " + (Convert.ToInt16(comboBox3.SelectedItem) * 12).ToString() + " and 在籍月 < " + (Convert.ToInt16(comboBox4.SelectedItem) * 12).ToString();
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】在籍年数";
                flg = false;
            }

            //誕生月
            if (checkBox9.Checked)
            {

                result += " and 誕生月 = '" + comboBox5.SelectedItem.ToString() + "'";
                flg = true;
            }

            if (flg)
            {
                conditions += "　【絞】誕生月";
                flg = false;
            }


            #endregion

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            //濱川用
            hama = "";
            if (result != "")
            {
                hama = " where ";
            }

            hama += result;

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //女性数
            int lady = 0;

            //年齢平均用
            double agesum = 0;
            int ageHi = 0;
            int ageLow = 0;

            //在籍平均用
            int zaisum = 0;
            int zaiHi = 0;
            int zaiLow = 0;

            if (checkBox2.Checked)
            {
                if (checkBox1.Checked)
                {
                    Disp.Columns.Add("社員番号", typeof(string));
                    Disp.Columns.Add("漢字氏名", typeof(string));
                    Disp.Columns.Add("カナ氏名", typeof(string));
                    Disp.Columns.Add("地区CD", typeof(string));
                    Disp.Columns.Add("地区名", typeof(string));
                    Disp.Columns.Add("組織CD", typeof(string));
                    Disp.Columns.Add("組織名", typeof(string));
                    Disp.Columns.Add("現場CD", typeof(string));
                    Disp.Columns.Add("現場名", typeof(string));
                    Disp.Columns.Add("役職CD", typeof(string));
                    Disp.Columns.Add("役職名", typeof(string));
                    Disp.Columns.Add("給与CD", typeof(string));
                    Disp.Columns.Add("給与名", typeof(string));

                    //Disp.Columns.Add("等級CD", typeof(int));
                    //Disp.Columns.Add("等級名", typeof(string));
                    //Disp.Columns.Add("号棒CD", typeof(int));
                    //Disp.Columns.Add("号棒名", typeof(string));
                    Disp.Columns.Add("入社年月日", typeof(string));
                    Disp.Columns.Add("退職年月日", typeof(string));
                    Disp.Columns.Add("生年月日", typeof(string));
                    Disp.Columns.Add("郵便番号", typeof(string));
                    Disp.Columns.Add("住所1", typeof(string));
                    Disp.Columns.Add("住所2", typeof(string));
                    Disp.Columns.Add("住所3", typeof(string));
                    Disp.Columns.Add("住所4", typeof(string));
                    Disp.Columns.Add("電話番号", typeof(string));
                    Disp.Columns.Add("年齢", typeof(string));
                    Disp.Columns.Add("在籍年月", typeof(string));
                    Disp.Columns.Add("在籍月", typeof(int));
                    Disp.Columns.Add("時給", typeof(int));
                    Disp.Columns.Add("日給", typeof(int));
                    Disp.Columns.Add("労働時間", typeof(int));
                    Disp.Columns.Add("週労働数", typeof(string));
                    Disp.Columns.Add("メール", typeof(string));
                    Disp.Columns.Add("性別", typeof(string));
                    //Disp.Columns.Add("血液型", typeof(string));
                    Disp.Columns.Add("在籍状況", typeof(string));
                    Disp.Columns.Add("友の会", typeof(string));
                    Disp.Columns.Add("健保", typeof(string));
                    Disp.Columns.Add("雇保", typeof(string));
                    Disp.Columns.Add("試用", typeof(string));
                    Disp.Columns.Add("契約", typeof(string));
                    Disp.Columns.Add("国籍", typeof(string));
                    Disp.Columns.Add("職種", typeof(string));
                    //Disp.Columns.Add("内線番号", typeof(string));
                    Disp.Columns.Add("担当区分", typeof(string));
                    Disp.Columns.Add("通勤1日単価", typeof(int));
                    Disp.Columns.Add("回数１単価", typeof(int));

                    foreach (DataRow row in dtrow)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号"] = row["社員番号"];
                        nr["漢字氏名"] = row["漢字氏名"];
                        nr["カナ氏名"] = row["カナ氏名"];
                        nr["地区CD"] = row["地区CD"];
                        nr["地区名"] = row["地区名"];
                        nr["組織CD"] = row["組織CD"];
                        nr["組織名"] = row["組織名"];
                        nr["現場CD"] = row["現場CD"];
                        nr["現場名"] = row["現場名"];
                        nr["役職CD"] = row["役職CD"];
                        nr["役職名"] = row["役職名"];
                        nr["給与CD"] = row["給与支給区分"];
                        nr["給与名"] = row["給与支給名称"];
                        nr["入社年月日"] = row["入社年月日"];
                        nr["退職年月日"] = row["退職年月日"];
                        nr["生年月日"] = row["生年月日"];
                        nr["性別"] = row["性別区分"].ToString() == "1" ? "男性" : "女性";

                        nr["在籍月"] = row["在籍月"];
                        nr["時給"] = row["時給"];
                        nr["日給"] = row["日給"];
                        nr["労働時間"] = row["労働時間"];
                        nr["週労働数"] = row["週労働数"];
                        nr["メール"] = row["ml"];

                        //nr["血液型"] = row["血液型"].Equals(DBNull.Value) ? "" : GetKetsueki[row["血液型"].ToString()];
                        nr["郵便番号"] = row["郵便番号"];
                        nr["住所1"] = row["住所1"];
                        nr["住所2"] = row["住所2"];
                        nr["住所3"] = row["住所3"];
                        nr["住所4"] = row["住所4"];
                        nr["電話番号"] = row["電話番号"];
                        nr["年齢"] = row["年齢"];
                        nr["在籍年月"] = row["在籍年月"];
                        nr["在籍状況"] = row["在籍状況"];
                        nr["友の会"] = row["友の会"];
                        nr["健保"] = row["健保"];
                        nr["雇保"] = row["雇保"];
                        nr["試用"] = row["試用"];
                        nr["契約"] = row["契約"];
                        nr["国籍"] = row["国籍"];
                        nr["職種"] = row["職種"];
                        //nr["内線番号"] = row["内線番号"];
                        nr["担当区分"] = row["担当区分"];
                        nr["通勤1日単価"] = row["通勤1日単価"];
                        nr["回数１単価"] = row["回数１単価"];
                        Disp.Rows.Add(nr);
                    }
                }
                else
                {
                    Disp.Columns.Add("社員番号\nカナ氏名\n漢字氏名", typeof(string));
                    Disp.Columns.Add("CD\n地区\n組織\n現場", typeof(string));
                    Disp.Columns.Add("名\n地区\n組織\n現場", typeof(string));
                    //Disp.Columns.Add("CD\n役職\n給与\n等級-号棒", typeof(string));
                    //Disp.Columns.Add("名\n役職\n給与\n等級-号棒", typeof(string));
                    Disp.Columns.Add("CD\n役職\n給与", typeof(string));
                    Disp.Columns.Add("名\n役職\n給与", typeof(string));
                    Disp.Columns.Add("入社日\n退職日\n生年月日", typeof(string));
                    Disp.Columns.Add("郵便番号\n住所\n電話番号", typeof(string));
                    Disp.Columns.Add("年齢\n在籍年月\n在籍状況", typeof(string));

                    foreach (DataRow row in dtrow)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号\nカナ氏名\n漢字氏名"] = row["社員番号"] + "\n" + row["カナ氏名"] + "\n" + row["漢字氏名"];
                        nr["CD\n地区\n組織\n現場"] = row["地区CD"] + "\n" + row["組織CD"] + "\n" + row["現場CD"];
                        nr["名\n地区\n組織\n現場"] = row["地区名"] + "\n" + row["組織名"] + "\n" + row["現場名"];
                        //nr["CD\n役職\n給与\n等級-号棒"] = row["役職CD"] + "\n" + row["給与支給区分"] + "\n" + row["等級コード"] + "-" + row["号棒コード"];
                        //nr["名\n役職\n給与\n等級-号棒"] = row["役職名"] + "\n" + row["給与支給名称"] + "\n" + row["等級名"] + "-" + row["号棒名"];
                        nr["CD\n役職\n給与"] = row["役職CD"] + "\n" + row["給与支給区分"];
                        nr["名\n役職\n給与"] = row["役職名"] + "\n" + row["給与支給名称"];
                        nr["入社日\n退職日\n生年月日"] = row["入社年月日"] + "\n" + row["退職年月日"] + "\n" + row["生年月日"];
                        nr["郵便番号\n住所\n電話番号"] = row["郵便番号"] + "\n" + row["住所1"] + " " + row["住所2"] + " " + row["住所3"] + " " + row["住所4"] + "\n" + row["電話番号"];
                        nr["年齢\n在籍年月\n在籍状況"] = row["年齢"] + "\n" + row["在籍年月"] + "\n" + row["在籍状況"];
                        Disp.Rows.Add(nr);
                    }
                }
            }
            else
            {
                if (checkBox1.Checked)
                {
                    Disp.Columns.Add("社員番号", typeof(string));
                    Disp.Columns.Add("漢字氏名", typeof(string));
                    Disp.Columns.Add("カナ氏名", typeof(string));
                    Disp.Columns.Add("地区名", typeof(string));
                    Disp.Columns.Add("組織名", typeof(string));
                    Disp.Columns.Add("現場名", typeof(string));
                    Disp.Columns.Add("役職名", typeof(string));
                    Disp.Columns.Add("給与名", typeof(string));
                    //Disp.Columns.Add("等級名", typeof(string));
                    //Disp.Columns.Add("号棒名", typeof(string));
                    Disp.Columns.Add("入社年月日", typeof(string));
                    Disp.Columns.Add("退職年月日", typeof(string));
                    Disp.Columns.Add("生年月日", typeof(string));
                    Disp.Columns.Add("郵便番号", typeof(string));
                    Disp.Columns.Add("住所1", typeof(string));
                    Disp.Columns.Add("住所2", typeof(string));
                    Disp.Columns.Add("住所3", typeof(string));
                    Disp.Columns.Add("住所4", typeof(string));
                    Disp.Columns.Add("電話番号", typeof(string));
                    Disp.Columns.Add("年齢", typeof(string));
                    Disp.Columns.Add("在籍年月", typeof(string));
                    Disp.Columns.Add("在籍月", typeof(int));
                    Disp.Columns.Add("時給", typeof(int));
                    Disp.Columns.Add("日給", typeof(int));
                    Disp.Columns.Add("労働時間", typeof(int));
                    Disp.Columns.Add("週労働数", typeof(string));
                    Disp.Columns.Add("メール", typeof(string));
                    Disp.Columns.Add("性別", typeof(string));
                    //Disp.Columns.Add("血液型", typeof(string));
                    Disp.Columns.Add("在籍状況", typeof(string));
                    Disp.Columns.Add("友の会", typeof(string));
                    Disp.Columns.Add("健保", typeof(string));
                    Disp.Columns.Add("雇保", typeof(string));
                    Disp.Columns.Add("試用", typeof(string));
                    Disp.Columns.Add("契約", typeof(string));
                    Disp.Columns.Add("国籍", typeof(string));
                    Disp.Columns.Add("職種", typeof(string));
                    //Disp.Columns.Add("内線番号", typeof(string));
                    Disp.Columns.Add("担当区分", typeof(string));
                    Disp.Columns.Add("通勤1日単価", typeof(int));
                    Disp.Columns.Add("回数１単価", typeof(int));
                    Disp.Columns.Add("休日区分名", typeof(string));

                    foreach (DataRow row in dtrow)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号"] = row["社員番号"];
                        nr["漢字氏名"] = row["漢字氏名"];
                        nr["カナ氏名"] = row["カナ氏名"];
                        nr["地区名"] = row["地区名"];
                        nr["組織名"] = row["組織名"];
                        nr["現場名"] = row["現場名"];
                        nr["役職名"] = row["役職名"];
                        nr["給与名"] = row["給与支給名称"];
                        //nr["等級名"] = row["等級名"];
                        //nr["号棒名"] = row["号棒名"];
                        nr["入社年月日"] = row["入社年月日"];
                        nr["退職年月日"] = row["退職年月日"];
                        nr["生年月日"] = row["生年月日"];
                        nr["性別"] = row["性別区分"].ToString() == "1" ? "男性" : "女性";

                        nr["在籍月"] = row["在籍月"];
                        nr["時給"] = row["時給"];
                        nr["日給"] = row["日給"];
                        nr["労働時間"] = row["労働時間"];
                        nr["週労働数"] = row["週労働数"];
                        nr["メール"] = row["ml"];

                        //TODO:Viewにて血液型名を表示するに変更させる必要がある　※二か所以上存在
                        //nr["血液型"] = row["血液型"].Equals(DBNull.Value) ? "" : GetKetsueki[row["血液型"].ToString()];
                        //nr["血液型"] = row["血液型"];

                        nr["郵便番号"] = row["郵便番号"];
                        nr["住所1"] = row["住所1"];
                        nr["住所2"] = row["住所2"];
                        nr["住所3"] = row["住所3"];
                        nr["住所4"] = row["住所4"];
                        nr["電話番号"] = row["電話番号"];
                        nr["年齢"] = row["年齢"];
                        nr["在籍年月"] = row["在籍年月"];
                        nr["在籍状況"] = row["在籍状況"];
                        nr["友の会"] = row["友の会"];
                        nr["健保"] = row["健保"];
                        nr["雇保"] = row["雇保"];
                        nr["試用"] = row["試用"];
                        nr["契約"] = row["契約"];
                        nr["国籍"] = row["国籍"];
                        nr["職種"] = row["職種"];
                        //nr["入社時年齢"] = row["入社時年齢"];
                        nr["担当区分"] = row["担当区分"];
                        nr["通勤1日単価"] = row["通勤1日単価"];
                        nr["回数１単価"] = row["回数１単価"];
                        nr["休日区分名"] = row["休日区分名"];

                        Disp.Rows.Add(nr);

                        //女性数
                        if (row["性別区分"].ToString() == "2") lady++;

                        //平均年齢用合計
                        agesum += Convert.ToDouble(row["年齢"]);

                        //最高年齢
                        if (ageHi < Convert.ToInt16(row["年齢"])) ageHi = Convert.ToInt16(row["年齢"]);

                        //最少年齢
                        if (ageLow == 0)
                        {
                            ageLow = ageLow = Convert.ToInt16(row["年齢"]);
                        }
                        else
                        {
                            if (ageLow > Convert.ToInt16(row["年齢"])) ageLow = Convert.ToInt16(row["年齢"]);
                        }

                        int zaim = row["在籍月"].Equals(DBNull.Value) ? 0 : Convert.ToInt32(row["在籍月"]);

                        //平均在籍用合計
                        zaisum += zaim;

                        //最高在籍年数
                        if (zaiHi < zaim) zaiHi = zaim;

                        //最少在籍年数
                        if (zaiLow == 0)
                        {
                            zaiLow = zaiLow = zaim;
                        }
                        else
                        {
                            if (zaiLow > zaim) zaiLow = zaim;
                        }
                    }
                }
                else
                {
                    Disp.Columns.Add("社員番号\nカナ氏名\n漢字氏名", typeof(string));
                    Disp.Columns.Add("地区\n組織\n現場", typeof(string));
                    //Disp.Columns.Add("役職\n給与\n等級-号棒", typeof(string));
                    Disp.Columns.Add("役職\n給与", typeof(string));
                    Disp.Columns.Add("入社日\n退職日\n生年月日", typeof(string));
                    Disp.Columns.Add("郵便番号\n住所\n電話番号", typeof(string));
                    Disp.Columns.Add("年齢\n在籍年月\n在籍状況", typeof(string));

                    foreach (DataRow row in dtrow)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号\nカナ氏名\n漢字氏名"] = row["社員番号"] + "\n" + row["カナ氏名"] + "\n" + row["漢字氏名"];
                        nr["地区\n組織\n現場"] = row["地区名"] + "\n" + row["組織名"] + "\n" + row["現場名"];
                        //nr["役職\n給与\n等級-号棒"] = row["役職名"] + "\n" + row["給与支給名称"] + "\n" + row["等級名"] + "-" + row["号棒名"];
                        nr["役職\n給与"] = row["役職名"] + "\n" + row["給与支給名称"];
                        nr["入社日\n退職日\n生年月日"] = row["入社年月日"] + "\n" + row["退職年月日"] + "\n" + row["生年月日"];
                        nr["郵便番号\n住所\n電話番号"] = row["郵便番号"] + "\n" + row["住所1"] + " " + row["住所2"] + " " + row["住所3"] + " " + row["住所4"] + "\n" + row["電話番号"];
                        nr["年齢\n在籍年月\n在籍状況"] = row["年齢"] + "\n" + row["在籍年月"] + "\n" + row["在籍状況"];
                        Disp.Rows.Add(nr);
                    }
                }
            }


            //データグリッドビューの高さ指定　※セット前にすること！
            if (checkBox1.Checked)
            {
                dataGridView1.RowTemplate.Height = 20;
            }
            else
            {
                dataGridView1.RowTemplate.Height = 60;
            }

            label13.Text = dtrow.Length.ToString() + " 名 (男性: " + (dtrow.Length - lady).ToString() + "名 / 女性: " + lady + "名)";

            double avgage = agesum / dtrow.Length;

            label14.Text = Math.Round(avgage, 2, MidpointRounding.AwayFromZero).ToString() + "才";

            int avgzaiyy = zaisum == 0 ? 0 : zaisum / dtrow.Length / 12;
            int avgzaimm = zaisum == 0 ? 0 : (zaisum / dtrow.Length) % 12;
            label21.Text = avgzaiyy.ToString() + "年" + avgzaimm.ToString() + "ヶ月";
            //label10.Text = "合計売上　\\" + uriageAll.ToString("#,0") + "円";

            label41.Text = ageHi.ToString() + "才";
            label42.Text = ageLow.ToString() + "才";
            label43.Text = (zaiHi / 12).ToString() + "年" + (zaiHi % 12).ToString() + "ヶ月";
            label44.Text = (zaiLow / 12).ToString() + "年" + (zaiLow % 12).ToString() + "ヶ月";

            dataGridView1.DataSource = Disp;

            //検索履歴登録
            Com.InHistory("21_従業員検索", conditions, dtrow.Length.ToString());

            if (checkBox2.Checked)
            {
                if (checkBox1.Checked)
                {
                    //一列表示でコード表示

                    // セル内で文字列を折り返えさない
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                }
                else
                {
                    //複数列表示でコード表示

                    // セル内で文字列を折り返す
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                    dataGridView1.Columns[0].Width = 130;
                    dataGridView1.Columns[1].Width = 80;
                    dataGridView1.Columns[2].Width = 210;
                    dataGridView1.Columns[3].Width = 130;
                    dataGridView1.Columns[4].Width = 130;
                    dataGridView1.Columns[5].Width = 110;
                    dataGridView1.Columns[6].Width = 300;
                    dataGridView1.Columns[7].Width = 110;
                }
            }
            else
            {
                if (checkBox1.Checked)
                {
                    //一列表示でコード表示しない

                    // セル内で文字列を折り返さない
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                }
                else
                {
                    //複数列表示でコード表示しない

                    // セル内で文字列を折り返す
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                    dataGridView1.Columns[0].Width = 150;
                    dataGridView1.Columns[1].Width = 200;
                    dataGridView1.Columns[2].Width = 120;
                    dataGridView1.Columns[3].Width = 100;
                    dataGridView1.Columns[4].Width = 450;
                    dataGridView1.Columns[5].Width = 100;
                }
            }
            //}

            System.GC.Collect();

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            SyousaiData(drv[0].ToString());
        }


        private void KinkyuuReset()
        {
            //緊急連絡先のリセット
            honkeitai.Text = "";
            honkotei.Text = "";

            kaz1name.Text = "";
            kaz1kana.Text = "";
            kaz1gara.Text = "";
            kaz1no.Text = "";

            kaz2name.Text = "";
            kaz2kana.Text = "";
            kaz2gara.Text = "";
            kaz2no.Text = "";

            kinkyuuday.Text = "";
            kinkyuusya.Text = "";
        }

        private void RoudouReset()
        {
            //労働条件
            koyoukaishibi.Visible = true;
            koyousyuuryoubi.Visible = true;

            keiyakunengetsu.Value = null;
            koyoukubun.Text = "";
            koyoukaishibi.Value = null;
            koyousyuuryoubi.Value = null;
            koushinkubun.Text = "";
            syuugyoubasyo.Text = "";
            gyoumunaiyou.Text = "";
            dtps0.Value = zerodt;
            dtpe0.Value = zerodt;
            dtps1.Value = zerodt;
            dtpe1.Value = zerodt;
            dtps2.Value = zerodt;
            dtpe2.Value = zerodt;
            dtps3.Value = zerodt;
            dtpe3.Value = zerodt;
            dtps4.Value = zerodt;
            dtpe4.Value = zerodt;
            dtps5.Value = zerodt;
            dtpe5.Value = zerodt;
            zikangairoudou.Text = "";
            yakankinmu.Text = "";
            kyuujitsukaisuu.Text = "";
            kyuusyustu.Text = "";
            teinen.Text = "";
            syouyo.Text = "";
            taisyokukin.Text = "";
            kinmuH.Text = "";
        }

        public Dictionary<string, string> GetKetsueki = new Dictionary<string, string>()
        {
            {"A1", "Ａ Rh+"},
            {"A2", "Ａ Rh-"},
            {"B1", "Ｂ Rh+"},
            {"B2", "Ｂ Rh-"},
            {"C1", "AB Rh+"},
            {"C2", "AB Rh-"},
            {"O1", "Ｏ Rh+"},
            {"O2", "Ｏ Rh-"},
        };


        private void GetKinkyuu(string str)
        {
            //緊急連絡先情報
            DataTable dtrenraku = new DataTable();
            dtrenraku = Com.GetDB("select* from dbo.k緊急連絡先 where 社員番号 = '" + str.Substring(0, 8) + "' ");
            //dataGridView8.DataSource = dtrenraku;

            if (dtrenraku.Rows.Count > 0)
            {
                honkeitai.Text = dtrenraku.Rows[0][1].ToString();
                honkotei.Text = dtrenraku.Rows[0][2].ToString();

                kaz1name.Text = dtrenraku.Rows[0][3].ToString();
                kaz1kana.Text = dtrenraku.Rows[0][4].ToString();
                kaz1gara.Text = dtrenraku.Rows[0][5].ToString();
                kaz1no.Text = dtrenraku.Rows[0][6].ToString();

                kaz2name.Text = dtrenraku.Rows[0][7].ToString();
                kaz2kana.Text = dtrenraku.Rows[0][8].ToString();
                kaz2gara.Text = dtrenraku.Rows[0][9].ToString();
                kaz2no.Text = dtrenraku.Rows[0][10].ToString();

                kinkyuuday.Text = dtrenraku.Rows[0][11].ToString();
                kinkyuusya.Text = dtrenraku.Rows[0][12].ToString();

            }
            else
            {
                KinkyuuReset();
            }
        }

        private void SyousaiData(string str)
        {
            //対象者のみに絞込
            DataRow[] targetDr = dt.Select("社員番号 = " + str.Substring(0, 8), "");

            //西暦表示の為 TODO:上に移動
            CultureInfo ci = new CultureInfo("ja-JP");
            ci.DateTimeFormat.Calendar = new JapaneseCalendar();

            foreach (DataRow row in targetDr)
            {
                syainNo.Text = row[0].Equals(DBNull.Value) ? "" : row[0].ToString();//社員番号
                kanzishimei.Text = row[1].Equals(DBNull.Value) ? "" : row[1].ToString();//漢字氏名
                kanashimei.Text = row[2].Equals(DBNull.Value) ? "" : row[2].ToString();//カナ氏名
                tikucode.Text = row[3].Equals(DBNull.Value) ? "" : row[3].ToString();//地区CD
                syokusyucode.Text = row[4].Equals(DBNull.Value) ? "" : row[4].ToString();//組織CD
                genbacode.Text = row[5].Equals(DBNull.Value) ? "" : row[5].ToString();//現場CD
                yakusyokucode.Text = row[6].Equals(DBNull.Value) ? "" : row[6].ToString();//役職CD
                kyuuyokubuncode.Text = row[7].Equals(DBNull.Value) ? "" : row[7].ToString();//給与支給区分
                tikumei.Text = row[8].Equals(DBNull.Value) ? "" : row[8].ToString();//地区名
                syokusyumei.Text = row[9].Equals(DBNull.Value) ? "" : row[9].ToString();//組織名
                genbamei.Text = row[10].Equals(DBNull.Value) ? "" : row[10].ToString();//現場名
                yakusyoku.Text = row[11].Equals(DBNull.Value) ? "" : row[11].ToString();//役職名
                kyuuyokubun.Text = row[12].Equals(DBNull.Value) ? "" : row[12].ToString();//給与支給名称
                //toukyuu.Text = row[13].Equals(DBNull.Value) ? "" : row[13].ToString();//等級コード
                //toukyuumei.Text = row[14].Equals(DBNull.Value) ? "" : row[14].ToString();//等級名
                //goubou.Text = row[15].Equals(DBNull.Value) ? "" : row[15].ToString();//号棒コード
                //gouboumei.Text = row[16].Equals(DBNull.Value) ? "" : row[16].ToString();//号棒名
                nyuusyabi.Text = row[17].Equals(DBNull.Value) ? "" : "(" + Convert.ToDateTime(row[17]).ToString("ggyy年", ci) + ") " + row[17].ToString();//入社年月日
                taisyokubi.Text = row[18].Equals(DBNull.Value) ? "" : "(" + Convert.ToDateTime(row[18]).ToString("ggyy年", ci) + ") " + row[18].ToString();//退職年月日
                seinengappi.Text = row[19].Equals(DBNull.Value) ? "" : "(" + Convert.ToDateTime(row[19]).ToString("ggyy年", ci) + ") " + row[19].ToString();//生年月日
                seibetsu.Text = row[20].ToString() == "1" ? "男性" : "女性";//性別区分
                //ketsuekigata.Text = row[22].Equals(DBNull.Value) ? "" : GetKetsueki[row[22].ToString()];//血液型
                yuubinNo.Text = row[23].Equals(DBNull.Value) ? "" : row[23].ToString();//郵便番号
                zyuusyo.Text = row[24].Equals(DBNull.Value) ? "" : row[24].ToString();//住所1
                zyuusyo.Text += row[25].Equals(DBNull.Value) ? "" : row[25].ToString();//住所2
                zyuusyo.Text += row[26].Equals(DBNull.Value) ? "" : row[26].ToString();//住所3
                zyuusyo.Text += row[27].Equals(DBNull.Value) ? "" : row[27].ToString();//住所4
                tel.Text = row[28].Equals(DBNull.Value) ? "" : row[28].ToString();//電話番号
                nenrei.Text = row[29].Equals(DBNull.Value) ? "" : row[29].ToString() + "才";//年齢
                kinnzoku.Text = row[30].Equals(DBNull.Value) ? "" : row[30].ToString();//在籍年月
                zaiseki.Text = row[32].Equals(DBNull.Value) ? "" : row[32].ToString();//在籍状況

                zikyuu.Text = row[39].Equals(DBNull.Value) ? "-" : Convert.ToInt16(row[39]).ToString();//時給
                nikyuu.Text = row[40].Equals(DBNull.Value) ? "-" : Convert.ToInt16(row[40]).ToString();//日給
                roudouzikan.Text = row[41].Equals(DBNull.Value) ? "-" : Convert.ToDecimal(row[41]).ToString("G4") + " 時間";//労働時間

                mail.Text = row[42].Equals(DBNull.Value) ? "" : row[42].ToString();//メールアドレス

                roudouD.Text = row[43].ToString();//週労働日数

                tomonokai.Text = row[44].ToString();//友の会
                syaho.Text = row[45].ToString();//健保
                koyou.Text = row[46].ToString();//雇保
                shiyou.Text = row[47].ToString();//試用期間
                keiyaku.Text = row[48].ToString();//契約
                kokuseki.Text = row[49].ToString();//国籍
                syokusyu.Text = row[50].ToString();//職種
                nyusyaold.Text = row[51].ToString() + "才";//入社時年齢
                kubun.Text = row[52].ToString();//担当区分
                //sanzikyuu.Text = Math.Round(Convert.ToDecimal(row[54])).ToString();//算出時給
                zeihyoukubun.Text = row[55].ToString();//税表区分
                syougai.Text = row[56].ToString();//障
                kahu.Text = row[57].ToString();//寡フ
                kinnrou.Text = row[58].ToString();//勤労
                gaikoku.Text = row[59].ToString();//外国人
                saigai.Text = row[60].ToString();//災害
                gakureki.Text = row[61].ToString();//最終学歴;
                syagai.Text = row[62].ToString();//社外経験;

                nyuusyanenngaku.Text = row[63].ToString();//年齢給;
                syagaigaku.Text = row[64].ToString();//経験給;
                gakurekigaku.Text = row[65].ToString();//学歴給;
                syokusyugaku.Text = row[66].ToString();//職務給;

                kizyungaigaku.Text = row[67].ToString();//基準外額;
                hyoukagaku.Text = row[70].ToString();//評価
                zinzi.Text = row[69].ToString(); //職種(担当事務)
                kounen.Text = row[71].ToString(); //厚生年金追加
                //zairyuuno.Text = row[72].ToString(); //在留カードNO
                zairyuuyuukou.Text = row[73].ToString(); //在留カード有効期限
                kyuuzitsu.Text = row[74].ToString(); //休日区分名
            }

            //担当区分一致したら異動ボタン、労働条件更新ボタン、労働条件出力ボタンを表示する
            if (Program.loginname == "親泊　美和子" || Program.loginname == "小園　玲奈" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" )
            {
                if ((zinzi.Text == "03_施設" || zinzi.Text == "04_エンジ" || zinzi.Text == "03_警備") && Convert.ToInt16(yakusyokucode.Text) > 130)
                {
                    idoubtn.Visible = true;
                    //roujex.Visible = true;
                    rowjkoushin.Visible = true;
                }
                else
                {
                    idoubtn.Visible = false;
                    //roujex.Visible = false;
                    rowjkoushin.Visible = false;
                }
            }
            //TODO 2503大濱さん宮古島応援のため

            else if (Program.loginname == "大浜　綾希子")
            //else if (Program.loginname == "大浜　綾希子" || Program.loginname == "佐久間　みどり")
            //else if (Program.loginname == "佐久間　みどり")
            {
                if ((zinzi.Text == "01_現業" || zinzi.Text == "02_客室") && Convert.ToInt16(yakusyokucode.Text) > 130)
                {
                    idoubtn.Visible = true;
                    //roujex.Visible = true;
                    rowjkoushin.Visible = true;
                }
                else
                {
                    idoubtn.Visible = false;
                    //roujex.Visible = false;
                    rowjkoushin.Visible = false;
                }
            }
            else if (Program.loginname == "喜屋武　大祐" || Program.loginname == "金城　智之" || Program.loginname == "佐久川　昌佳")
            {
                idoubtn.Visible = true;
                //roujex.Visible = true;
                rowjkoushin.Visible = true;
            }
            else
            {

                if (kubun.Text == Program.loginbusyo)
                {
                    idoubtn.Visible = true;
                    //roujex.Visible = true;
                    rowjkoushin.Visible = true;
                }
                else
                {
                    idoubtn.Visible = false;
                    //roujex.Visible = false;
                    rowjkoushin.Visible = false;
                }
            }





            GetShikaku();

            //登録手当情報
            DataTable dttouroku = new DataTable();
            dttouroku = Com.GetDB("select 登録コード, 登録手当名称, 登録情報, 有効期限, 規定額, 手当対象額 from dbo.t登録手当一覧 where 社員番号 = '" + str.Substring(0, 8) + "'");

            dataGridView7.DataSource = dttouroku;
            dataGridView7.Columns[4].DefaultCellStyle.Format = "#,0";
            dataGridView7.Columns[5].DefaultCellStyle.Format = "#,0";
            dataGridView7.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView7.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            GetKinkyuu(str);

            //家族情報
            DataTable dtkazoku = new DataTable();
            dtkazoku = Com.GetDB("select* from dbo.k家族情報取得('" + str.Substring(0, 8) + "') order by ソート順");
            dataGridView8.DataSource = dtkazoku;

            dataGridView8.Columns[14].Visible = false;
            //dataGridView7.Columns[4].DefaultCellStyle.Format = "#,0";
            //dataGridView7.Columns[5].DefaultCellStyle.Format = "#,0";
            //dataGridView7.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView7.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //通勤管理情報
            DataTable dttuukin = new DataTable();
            dttuukin = Com.GetDB("select * from dbo.t通勤管理テーブル where 社員番号 = '" + str.Substring(0, 8) + "'");
            dataGridView10.DataSource = dttuukin;

            //研修情報取得
            DataTable kensyuudt = new DataTable();
            kensyuudt = Com.GetDB("exec 研修情報取得 '" + str.Substring(0, 8) + "'");
            dataGridView6.DataSource = kensyuudt;

            //端末情報取得
            dgvtanko.DataSource = Com.GetDB("select * from dbo.t端末管理テーブル where 社員番号 = '" + str.Substring(0, 8) + "'");
            dgvtang.DataSource = Com.GetDB("select * from dbo.t端末管理テーブル where (社員番号 is null or 社員番号 = '') and 組織CD = '" + syokusyucode.Text + "' and 現場CD = '" + genbacode.Text + "'");

            //労働条件情報
            DataTable roudoudt = new DataTable();
            roudoudt = Com.GetDB("select * from dbo.r労働条件 where 社員番号 = '" + str.Substring(0, 8) + "'");

            if (roudoudt.Rows.Count < 1)
            {
                //退職者の対応
                RoudouReset();
  
            }
            else
            {


                //null対応
                DateTime zerodt = new System.DateTime(2022, 12, 31, 0, 0, 0, 0);

                if (roudoudt.Rows[0]["契約年月"].Equals(DBNull.Value) || roudoudt.Rows[0]["契約年月"].Equals(""))
                {
                    keiyakunengetsu.Value = null;
                }
                else
                {
                    keiyakunengetsu.Value = Convert.ToDateTime(roudoudt.Rows[0]["契約年月"].ToString());
                }

                koyoukubun.SelectedIndex = koyoukubun.FindString(roudoudt.Rows[0]["雇用区分"].ToString());

                if (roudoudt.Rows[0]["雇用開始日"].Equals(DBNull.Value) || roudoudt.Rows[0]["雇用開始日"].Equals(""))
                {
                    koyoukaishibi.Value = null;
                }
                else
                {
                    koyoukaishibi.Value = Convert.ToDateTime(roudoudt.Rows[0]["雇用開始日"].ToString());
                }

                if (roudoudt.Rows[0]["雇用終了日"].Equals(DBNull.Value) || roudoudt.Rows[0]["雇用終了日"].Equals(""))
                {
                    koyousyuuryoubi.Value = null;
                }
                else
                {
                    koyousyuuryoubi.Value = Convert.ToDateTime(roudoudt.Rows[0]["雇用終了日"].ToString());
                }

                //koyoukaishibi.Value = roudoudt.Rows[0]["雇用開始日"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["雇用開始日"].ToString());
                //koyousyuuryoubi.Value = roudoudt.Rows[0]["雇用終了日"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["雇用終了日"].ToString());

                koushinkubun.SelectedIndex = koushinkubun.FindString(roudoudt.Rows[0]["更新区分"].ToString());

                syuugyoubasyo.Text = roudoudt.Rows[0]["就業場所"].ToString();
                gyoumunaiyou.Text = roudoudt.Rows[0]["業務内容"].ToString();

                dtps0.Value = roudoudt.Rows[0]["定始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["定始業"].ToString());
                dtpe0.Value = roudoudt.Rows[0]["定終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["定終業"].ToString());
                dtps1.Value = roudoudt.Rows[0]["S1始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S1始業"].ToString());
                dtpe1.Value = roudoudt.Rows[0]["S1終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S1終業"].ToString());
                dtps2.Value = roudoudt.Rows[0]["S2始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S2始業"].ToString());
                dtpe2.Value = roudoudt.Rows[0]["S2終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S2終業"].ToString());
                dtps3.Value = roudoudt.Rows[0]["S3始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S3始業"].ToString());
                dtpe3.Value = roudoudt.Rows[0]["S3終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S3終業"].ToString());
                dtps4.Value = roudoudt.Rows[0]["S4始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S4始業"].ToString());
                dtpe4.Value = roudoudt.Rows[0]["S4終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S4終業"].ToString());
                dtps5.Value = roudoudt.Rows[0]["S5始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S5始業"].ToString());
                dtpe5.Value = roudoudt.Rows[0]["S5終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S5終業"].ToString());

                zikangairoudou.SelectedIndex = koushinkubun.FindString(roudoudt.Rows[0]["時間外労働区分"].ToString());
                yakankinmu.SelectedIndex = yakankinmu.FindString(roudoudt.Rows[0]["夜間勤務区分"].ToString());

                kinmuH.Text = "【" + roudouzikan.Text + "】";

                //TODO 週労働数との連動処理
                syuuroucopy.Text = roudouD.Text;
                switch (roudouD.Text)
                {
                    case "5日以上":

                        if (kyuuyokubuncode.Text == "E1")
                        {
                            //パート
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("2");
                        }
                        else if (kyuuyokubuncode.Text == "F1")
                        {
                            //アルバイト
                            //TODO 設定無で
                        }
                        else if (kyuuyokubuncode.Text == "C1")
                        {
                            //正社員
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("1");
                        }
                        else if (kyuuyokubuncode.Text == "B1")
                        {
                            //兼務役員
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("1");
                        }
                        break;
                    case "4日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("3");
                        break;
                    case "3日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("4");
                        break;
                    case "2日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("5");
                        break;
                    case "1日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("6");
                        break;
                    default:
                        //役員
                        kyuujitsukaisuu.SelectedIndex = -1;
                        break;
                }

                if (roudoudt.Rows[0]["休日回数"].ToString() != "")
                {
                    kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString(roudoudt.Rows[0]["休日回数"].ToString());
                }

                kyuusyustu.SelectedIndex = kyuusyustu.FindString(roudoudt.Rows[0]["休出有無"].ToString());
                teinen.SelectedIndex = teinen.FindString(roudoudt.Rows[0]["定年区分"].ToString());
                syouyo.SelectedIndex = syouyo.FindString(roudoudt.Rows[0]["賞与区分"].ToString());
                taisyokukin.SelectedIndex = taisyokukin.FindString(roudoudt.Rows[0]["退職金区分"].ToString());

            }

            //係長以上は参照不可。
            if ((Program.dispZinzi == 8 && kubun.Text == "11_北部")  //北部
                || (Program.dispZinzi == 7 && kubun.Text == "12_八重山") //八重山
                //|| (Program.dispZinzi == 5 && (kubun.Text == "エンジ" || kubun.Text == "施設" || kubun.Text == "サービス" || kubun.Text == "警備")) //エンジ・施設
                || (Program.dispZinzi == 5 && (zinzi.Text == "03_施設" || zinzi.Text == "04_エンジ" || zinzi.Text == "03_警備") && Convert.ToInt16(yakusyokucode.Text) > 130) //エンジ・施設
                || (Program.dispZinzi == 4 && (kubun.Text == "01_現業" || kubun.Text == "02_客室" || kyuuyokubuncode.Text == "E1" || kyuuyokubuncode.Text == "F1") && Convert.ToInt16(yakusyokucode.Text) > 130) //現業・客室
                || (Program.dispZinzi == 3 && kubun.Text == "05_PPP/PFI" && Convert.ToInt16(yakusyokucode.Text) > 130) //指定        
                || Program.dispZinzi == 99
                )
            {

                SqlConnection Cn;
                SqlCommand Cmd;
                SqlDataAdapter da;
                //DataTable dtkotei = new DataTable();
                dtkotei.Clear();

                DataTable dttanka = new DataTable();


                if (Program.dispZinzi == 5 && Convert.ToInt16(yakusyokucode.Text) <= 135)
                {
                    //施設・エンジで係長以上は固定給与非表示
                    dataGridView2.DataSource = null;
                    dataGridView9.DataSource = null;

                    dgvmeisai.DataSource = null;
                    dgvkouza.DataSource = null;
                    dgvkintai.DataSource = null;

                    pictureBox1.Image = null;

                }
                else
                {

                    //固定給表示
                    using (Cn = new SqlConnection(Com.SQLConstr))
                    {
                        Cn.Open();
                        Cmd = Cn.CreateCommand();
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[k固定給取得]";
                        Cmd.Parameters.Add(new SqlParameter("Num", SqlDbType.VarChar));
                        Cmd.Parameters["Num"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["Num"].Value = str.Substring(0, 8);

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dtkotei);
                    }

                    dataGridView2.DataSource = dtkotei;
                    dataGridView2.Columns[1].DefaultCellStyle.Format = "#,0";
                    dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    if (tabControl1.SelectedIndex == 2)
                    {
                        //経歴情報
                        DataTable dtkeireki = new DataTable();
                        dtkeireki = Com.GetDB("select * from dbo.k雇用給与変更履歴表示('" + str.Substring(0, 8) + "') order by 適用開始日");

                        dataGridView9.DataSource = dtkeireki;

                        dataGridView9.Columns[6].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[7].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[8].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[9].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[10].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[11].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[12].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[13].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[14].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[15].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[16].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[17].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[18].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[19].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[20].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[21].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[22].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[23].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[24].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[25].DefaultCellStyle.Format = "#,0";
                        dataGridView9.Columns[26].DefaultCellStyle.Format = "#,0";

                        dataGridView9.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[21].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[24].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView9.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    else if (tabControl1.SelectedIndex == 9)
                    {
                        //
                        //口座情報
                        DataTable dtkouza = new DataTable();
                        dtkouza = Com.GetDB("select * from dbo.口座情報取得 where 社員番号 = " + str.Substring(0, 8));

                        dgvkouza.DataSource = dtkouza;


                        //顔写真
                        DataTable photodt = new DataTable();
                        photodt = Com.GetDB("select 顔写真 from QUATRO.dbo.KJMTKIHON a left join dbo.社員基本情報 b on a.個人識別ＩＤ = b.個人識別ＩＤ where b.社員番号 = " + str.Substring(0, 8));

                        if (photodt.Rows[0]["顔写真"].Equals(DBNull.Value))
                        {
                            pictureBox1.Image = null;
                        }
                        else
                        {
                            Byte[] byteBLOBData = new Byte[0];
                            byteBLOBData = (Byte[])(photodt.Rows[0]["顔写真"]);
                            MemoryStream stmBLOBData = new MemoryStream(byteBLOBData);
                            pictureBox1.Image = Image.FromStream(stmBLOBData);

                            pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;

                            // 幅0.5倍、高さ0.5倍のイメージを作成する
                            Bitmap bmp = new Bitmap(
                                pictureBox1.Image,
                                (int)(pictureBox1.Image.Width * 0.25),
                                (int)(pictureBox1.Image.Height * 0.25));

                            pictureBox1.Image = bmp;

                            if (Program.loginname == "喜屋武　大祐")
                            {
                                Bitmap bmp2 = new Bitmap(
                                pictureBox1.Image,
                                (int)(pictureBox1.Image.Width * 1),
                                (int)(pictureBox1.Image.Height * 1));
                                bmp2.Save(@"C:\Users\21151800\Documents\photo\test2.jpg");
                            }
                        }
                    }
                    else if (tabControl1.SelectedIndex == 10)
                    {
                        //
                        //勤怠情報
                        DataTable dtkintai = new DataTable();
                        dtkintai = Com.GetDB("select * from k勤怠履歴 where 社員番号 = " + str.Substring(0, 8) + " order by 処理年月 desc");

                        dgvkintai.DataSource = dtkintai;
                    }
                    else if (tabControl1.SelectedIndex == 11)
                    {
                        //
                        //明細情報
                        DataTable dtmeisai = new DataTable();
                        dtmeisai = Com.GetDB("select * from m明細履歴 where 社員番号 = " + str.Substring(0, 8) + " order by 処理年月 desc");

                        dgvmeisai.DataSource = dtmeisai;
                    }

                }
            }
            else
            {
                //1908 追加
                sanzikyuu.Text = "";

                //zitan.Text = "";
                //zangyou.Text = "";
                //shinya.Text = "";
                //syokyuu.Text = "";
                //houkyuu.Text = "";
                //entyou.Text = "";

                dataGridView2.DataSource = null;

                dataGridView9.DataSource = null;

                dgvmeisai.DataSource = null;
                dgvkouza.DataSource = null;
                dgvkintai.DataSource = null;
                pictureBox1.Image = null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ClickGetData();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                ClickGetData();
            }
        }

        private void ClickGetData()
        {
            //離職率タブ対応
            if (tabControl1.SelectedIndex == 8 && !checkBox6.Checked)
            {
                MessageBox.Show("退職率タブ表示では「退職者も含める」にチェックを入れてください");
                return;
            }

            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ表示
            DataView();

            //離職率タブ対応
            if (tabControl1.SelectedIndex == 8) GetTaisyokuData();


            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
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

        private void label58_Click(object sender, EventArgs e)
        {
            if (checkedListBox4.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox4.Items.Count; i++)
                {
                    checkedListBox4.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox4.Items.Count; i++)
                {
                    checkedListBox4.SetItemChecked(i, false);
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                label29.Visible = true;
            }
            else
            {
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                label29.Visible = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                dateTimePicker3.Visible = true;
                dateTimePicker4.Visible = true;
                label30.Visible = true;
            }
            else
            {
                dateTimePicker3.Visible = false;
                dateTimePicker4.Visible = false;
                label30.Visible = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                dateTimePicker5.Visible = true;
                dateTimePicker6.Visible = true;
                label31.Visible = true;
            }
            else
            {
                dateTimePicker5.Visible = false;
                dateTimePicker6.Visible = false;
                label31.Visible = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                label25.Visible = true;
                label26.Visible = true;
            }
            else
            {
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                label25.Visible = false;
                label26.Visible = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                label27.Visible = true;
                label28.Visible = true;
            }
            else
            {
                comboBox3.Visible = false;
                comboBox4.Visible = false;
                label27.Visible = false;
                label28.Visible = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked)
            {
                comboBox5.Visible = true;
            }
            else
            {
                comboBox5.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //クリア処理
            Clear();
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.C)
            {
                GetClip();
            }
        }

        private void GetClip()
        {
            Clipboard.SetDataObject(dataGridView1.GetClipboardContent());

            CopyReason cr = new CopyReason("従業員");
            cr.ShowDialog();

            dataGridView1.ClearSelection();
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetClip();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //TODO 異動申請権限がありません

            Genpyou Genpyou = new Genpyou(syainNo.Text);
            Genpyou.Show();
        }

        private void GetDataBetsu()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable dt = new DataTable();
            string sql = "";

            //sql = "select '【役職別】' as 項目, null as 対象者数, null as 最高齢, null as 最小齢, null as 平均齢, null as 最高在籍年数, null as 最小在籍年数, null as 平均在籍年数 union all ";
            sql = "select '【役職別】' as 項目, null as 対象者数, null as 平均齢, null as 平均在籍年数 union all ";
            //sql += "select 役職名 as 項目, count(*) as 対象者数, max(年齢) as 最高齢, min(年齢) as 最小齢, avg(年齢) as 平均齢, max(在籍年月) as 最高在籍年数, min(在籍年月) as 最小在籍年数, ";
            sql += "select 役職名 as 項目, count(*) as 対象者数, avg(年齢) as 平均齢, ";
            sql += "RIGHT('00' + CONVERT(varchar(10), avg(在籍月)/12), 2) + '年' + RIGHT('00' + CONVERT(varchar(10), avg(在籍月) % 12), 2) + 'ヶ月' as 平均在籍年数 ";
            sql += "from dbo.accessNew ";
            sql += hama;
            sql += " group by 役職名, 役職CD union all ";
            //sql += "select '【給与別】' as 項目, null as 対象者数, null as 最高齢, null as 最小齢, null as 平均齢, null as 最高在籍年数, null as 最小在籍年数, null as 平均在籍年数 union all ";
            sql += "select '【給与別】' as 項目, null as 対象者数, null as 平均齢, null as 平均在籍年数 union all ";
            sql += "select 給与支給名称 as 項目, count(*) as 対象者数, avg(年齢) as 平均齢, ";
            sql += "RIGHT('00' + CONVERT(varchar(10), avg(在籍月)/12), 2) + '年' + RIGHT('00' + CONVERT(varchar(10), avg(在籍月) % 12), 2) + 'ヶ月' as 平均在籍年数 ";
            sql += "from dbo.accessNew ";
            sql += hama;
            sql += " group by 給与支給名称, 給与支給区分 union all ";
            //sql += "select '【組織別】' as 項目, null as 対象者数, null as 最高齢, null as 最小齢, null as 平均齢, null as 最高在籍年数, null as 最小在籍年数, null as 平均在籍年数 union all ";
            sql += "select '【組織別】' as 項目, null as 対象者数, null as 平均齢, null as 平均在籍年数 union all ";
            sql += "select 組織名 as 項目, count(*) as 対象者数, avg(年齢) as 平均齢, ";
            sql += "RIGHT('00' + CONVERT(varchar(10), avg(在籍月)/12), 2) + '年' + RIGHT('00' + CONVERT(varchar(10), avg(在籍月) % 12), 2) + 'ヶ月' as 平均在籍年数 ";
            sql += "from dbo.accessNew ";
            sql += hama;
            sql += " group by 組織名, 組織CD union all ";
            //sql += "select '【現場別】' as 項目, null as 対象者数, null as 最高齢, null as 最小齢, null as 平均齢, null as 最高在籍年数, null as 最小在籍年数, null as 平均在籍年数 union all ";
            sql += "select '【現場別】' as 項目, null as 対象者数, null as 平均齢, null as 平均在籍年数 union all ";
            sql += "select 現場名 as 項目, count(*) as 対象者数, avg(年齢) as 平均齢, ";
            sql += "RIGHT('00' + CONVERT(varchar(10), avg(在籍月)/12), 2) + '年' + RIGHT('00' + CONVERT(varchar(10), avg(在籍月) % 12), 2) + 'ヶ月' as 平均在籍年数 ";
            sql += "from dbo.accessNew ";
            sql += hama;
            sql += " group by 現場名, 現場CD ";
            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
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

            dataGridView4.DataSource = dt;

        }

        private void GetTaisyokuData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable dt = new DataTable();

            string str = "";
            str = "select ";
            str += "sum(case when 退職年月日 between '2005/04/01' and '2006/03/31' then 1 else 0 end ) as [34期退職], ";
            str += "sum(case when 退職年月日 between '2006/04/01' and '2007/03/31' then 1 else 0 end ) as [35期退職], ";
            str += "sum(case when 退職年月日 between '2007/04/01' and '2008/03/31' then 1 else 0 end ) as [36期退職], ";
            str += "sum(case when 退職年月日 between '2008/04/01' and '2009/03/31' then 1 else 0 end ) as [37期退職], ";
            str += "sum(case when 退職年月日 between '2009/04/01' and '2010/03/31' then 1 else 0 end ) as [38期退職], ";
            str += "sum(case when 退職年月日 between '2010/04/01' and '2011/03/31' then 1 else 0 end ) as [39期退職], ";
            str += "sum(case when 退職年月日 between '2011/04/01' and '2012/03/31' then 1 else 0 end ) as [40期退職], ";
            str += "sum(case when 退職年月日 between '2012/04/01' and '2013/03/31' then 1 else 0 end ) as [41期退職], ";
            str += "sum(case when 退職年月日 between '2013/04/01' and '2014/03/31' then 1 else 0 end ) as [42期退職], ";
            str += "sum(case when 退職年月日 between '2014/04/01' and '2015/03/31' then 1 else 0 end ) as [43期退職], ";
            str += "sum(case when 退職年月日 between '2015/04/01' and '2016/03/31' then 1 else 0 end ) as [44期退職], ";
            str += "sum(case when 退職年月日 between '2016/04/01' and '2017/03/31' then 1 else 0 end ) as [45期退職], ";
            str += "sum(case when 退職年月日 between '2017/04/01' and '2018/03/31' then 1 else 0 end ) as [46期退職], ";
            str += "sum(case when 退職年月日 between '2018/04/01' and '2019/03/31' then 1 else 0 end ) as [47期退職], ";
            str += "sum(case when 退職年月日 between '2019/04/01' and '2020/03/31' then 1 else 0 end ) as [48期退職], ";
            str += "sum(case when 退職年月日 between '2020/04/01' and '2021/03/31' then 1 else 0 end ) as [49期退職], ";
            str += "sum(case when 退職年月日 between '2021/04/01' and '2022/03/31' then 1 else 0 end ) as [50期退職], ";
            str += "sum(case when 退職年月日 between '2022/04/01' and '2023/03/31' then 1 else 0 end ) as [51期退職], ";

            str += "sum(case when 入社年月日 between '2005/04/01' and '2006/03/31' then 1 else 0 end ) as [34期入社], ";
            str += "sum(case when 入社年月日 between '2006/04/01' and '2007/03/31' then 1 else 0 end ) as [35期入社], ";
            str += "sum(case when 入社年月日 between '2007/04/01' and '2008/03/31' then 1 else 0 end ) as [36期入社], ";
            str += "sum(case when 入社年月日 between '2008/04/01' and '2009/03/31' then 1 else 0 end ) as [37期入社], ";
            str += "sum(case when 入社年月日 between '2009/04/01' and '2010/03/31' then 1 else 0 end ) as [38期入社], ";
            str += "sum(case when 入社年月日 between '2010/04/01' and '2011/03/31' then 1 else 0 end ) as [39期入社], ";
            str += "sum(case when 入社年月日 between '2011/04/01' and '2012/03/31' then 1 else 0 end ) as [40期入社], ";
            str += "sum(case when 入社年月日 between '2012/04/01' and '2013/03/31' then 1 else 0 end ) as [41期入社], ";
            str += "sum(case when 入社年月日 between '2013/04/01' and '2014/03/31' then 1 else 0 end ) as [42期入社], ";
            str += "sum(case when 入社年月日 between '2014/04/01' and '2015/03/31' then 1 else 0 end ) as [43期入社], ";
            str += "sum(case when 入社年月日 between '2015/04/01' and '2016/03/31' then 1 else 0 end ) as [44期入社], ";
            str += "sum(case when 入社年月日 between '2016/04/01' and '2017/03/31' then 1 else 0 end ) as [45期入社], ";
            str += "sum(case when 入社年月日 between '2017/04/01' and '2018/03/31' then 1 else 0 end ) as [46期入社], ";
            str += "sum(case when 入社年月日 between '2018/04/01' and '2019/03/31' then 1 else 0 end ) as [47期入社], ";
            str += "sum(case when 入社年月日 between '2019/04/01' and '2020/03/31' then 1 else 0 end ) as [48期入社], ";
            str += "sum(case when 入社年月日 between '2020/04/01' and '2021/03/31' then 1 else 0 end ) as [49期入社], ";
            str += "sum(case when 入社年月日 between '2021/04/01' and '2022/03/31' then 1 else 0 end ) as [50期入社], ";
            str += "sum(case when 入社年月日 between '2022/04/01' and '2023/03/31' then 1 else 0 end ) as [51期入社], ";

            str += "sum(case when 入社年月日 <= '2005/04/01' and (退職年月日 >= '2005/04/01' or 退職年月日 is null) then 1 else 0 end ) as [34期期首], ";
            str += "sum(case when 入社年月日 <= '2006/04/01' and (退職年月日 >= '2006/04/01' or 退職年月日 is null) then 1 else 0 end ) as [35期期首], ";
            str += "sum(case when 入社年月日 <= '2007/04/01' and (退職年月日 >= '2007/04/01' or 退職年月日 is null) then 1 else 0 end ) as [36期期首], ";
            str += "sum(case when 入社年月日 <= '2008/04/01' and (退職年月日 >= '2008/04/01' or 退職年月日 is null) then 1 else 0 end ) as [37期期首], ";
            str += "sum(case when 入社年月日 <= '2009/04/01' and (退職年月日 >= '2009/04/01' or 退職年月日 is null) then 1 else 0 end ) as [38期期首], ";
            str += "sum(case when 入社年月日 <= '2010/04/01' and (退職年月日 >= '2010/04/01' or 退職年月日 is null) then 1 else 0 end ) as [39期期首], ";
            str += "sum(case when 入社年月日 <= '2011/04/01' and (退職年月日 >= '2011/04/01' or 退職年月日 is null) then 1 else 0 end ) as [40期期首], ";
            str += "sum(case when 入社年月日 <= '2012/04/01' and (退職年月日 >= '2012/04/01' or 退職年月日 is null) then 1 else 0 end ) as [41期期首], ";
            str += "sum(case when 入社年月日 <= '2013/04/01' and (退職年月日 >= '2013/04/01' or 退職年月日 is null) then 1 else 0 end ) as [42期期首], ";
            str += "sum(case when 入社年月日 <= '2014/04/01' and (退職年月日 >= '2014/04/01' or 退職年月日 is null) then 1 else 0 end ) as [43期期首], ";
            str += "sum(case when 入社年月日 <= '2015/04/01' and (退職年月日 >= '2015/04/01' or 退職年月日 is null) then 1 else 0 end ) as [44期期首], ";
            str += "sum(case when 入社年月日 <= '2016/04/01' and (退職年月日 >= '2016/04/01' or 退職年月日 is null) then 1 else 0 end ) as [45期期首], ";
            str += "sum(case when 入社年月日 <= '2017/04/01' and (退職年月日 >= '2017/04/01' or 退職年月日 is null) then 1 else 0 end ) as [46期期首], ";
            str += "sum(case when 入社年月日 <= '2018/04/01' and (退職年月日 >= '2018/04/01' or 退職年月日 is null) then 1 else 0 end ) as [47期期首], ";
            str += "sum(case when 入社年月日 <= '2019/04/01' and (退職年月日 >= '2019/04/01' or 退職年月日 is null) then 1 else 0 end ) as [48期期首], ";
            str += "sum(case when 入社年月日 <= '2020/04/01' and (退職年月日 >= '2020/04/01' or 退職年月日 is null) then 1 else 0 end ) as [49期期首], ";
            str += "sum(case when 入社年月日 <= '2021/04/01' and (退職年月日 >= '2021/04/01' or 退職年月日 is null) then 1 else 0 end ) as [50期期首],  ";
            str += "sum(case when 入社年月日 <= '2022/04/01' and (退職年月日 >= '2022/04/01' or 退職年月日 is null) then 1 else 0 end ) as [51期期首]  ";

            //str += "from dbo.accessNew where reskey like '%" + textBox1.Text + "%' and 地区名 like '%" + comboBox1.SelectedItem + "%'";
            str += "from dbo.accessNew " + hama;
            Com.InHistory("退職率", "", "");

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = str;
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

            DataTable Disp = new DataTable();
            Disp.Columns.Add("期", typeof(string));
            Disp.Columns.Add("期首従業員数", typeof(int));
            Disp.Columns.Add("入職者数", typeof(int));
            Disp.Columns.Add("退職者数", typeof(int));
            Disp.Columns.Add("入職率", typeof(decimal));
            Disp.Columns.Add("退職率", typeof(decimal));

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr51 = Disp.NewRow();
                nr51["期"] = "51期(2022年4月～)";
                nr51["期首従業員数"] = row["51期期首"];
                nr51["入職者数"] = row["51期入社"];
                nr51["退職者数"] = row["51期退職"];
                nr51["入職率"] = row["51期期首"].ToString() == "" ? 0 : row["51期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["51期入社"]) / Convert.ToDecimal(row["51期期首"]) * 100, 1);
                nr51["退職率"] = row["51期期首"].ToString() == "" ? 0 : row["51期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["51期退職"]) / Convert.ToDecimal(row["51期期首"]) * 100, 1);
                Disp.Rows.Add(nr51);

                DataRow nr50 = Disp.NewRow();
                nr50["期"] = "50期(2021年4月～)";
                nr50["期首従業員数"] = row["50期期首"];
                nr50["入職者数"] = row["50期入社"];
                nr50["退職者数"] = row["50期退職"];
                nr50["入職率"] = row["50期期首"].ToString() == "" ? 0 : row["50期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["50期入社"]) / Convert.ToDecimal(row["50期期首"]) * 100, 1);
                nr50["退職率"] = row["50期期首"].ToString() == "" ? 0 : row["50期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["50期退職"]) / Convert.ToDecimal(row["50期期首"]) * 100, 1);
                Disp.Rows.Add(nr50);

                DataRow nr49 = Disp.NewRow();
                nr49["期"] = "49期(2020年4月～)";
                nr49["期首従業員数"] = row["49期期首"];
                nr49["入職者数"] = row["49期入社"];
                nr49["退職者数"] = row["49期退職"];
                nr49["入職率"] = row["49期期首"].ToString() == "" ? 0 : row["49期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["49期入社"]) / Convert.ToDecimal(row["49期期首"]) * 100, 1);
                nr49["退職率"] = row["49期期首"].ToString() == "" ? 0 : row["49期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["49期退職"]) / Convert.ToDecimal(row["49期期首"]) * 100, 1);
                Disp.Rows.Add(nr49);

                DataRow nr48 = Disp.NewRow();
                nr48["期"] = "48期(2019年4月～)";
                nr48["期首従業員数"] = row["48期期首"];
                nr48["入職者数"] = row["48期入社"];
                nr48["退職者数"] = row["48期退職"];
                nr48["入職率"] = row["48期期首"].ToString() == "" ? 0 : row["48期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["48期入社"]) / Convert.ToDecimal(row["48期期首"]) * 100, 1);
                nr48["退職率"] = row["48期期首"].ToString() == "" ? 0 : row["48期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["48期退職"]) / Convert.ToDecimal(row["48期期首"]) * 100, 1);
                Disp.Rows.Add(nr48);

                DataRow nr47 = Disp.NewRow();
                nr47["期"] = "47期(2018年4月～)";
                nr47["期首従業員数"] = row["47期期首"];
                nr47["入職者数"] = row["47期入社"];
                nr47["退職者数"] = row["47期退職"];
                nr47["入職率"] = row["47期期首"].ToString() == "" ? 0 : row["47期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["47期入社"]) / Convert.ToDecimal(row["47期期首"]) * 100, 1);
                nr47["退職率"] = row["47期期首"].ToString() == "" ? 0 : row["47期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["47期退職"]) / Convert.ToDecimal(row["47期期首"]) * 100, 1);
                Disp.Rows.Add(nr47);

                DataRow nr46 = Disp.NewRow();
                nr46["期"] = "46期(2017年4月～)";
                nr46["期首従業員数"] = row["46期期首"];
                nr46["入職者数"] = row["46期入社"];
                nr46["退職者数"] = row["46期退職"];
                nr46["入職率"] = row["46期期首"].ToString() == "" ? 0 : row["46期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["46期入社"]) / Convert.ToDecimal(row["46期期首"]) * 100, 1);
                nr46["退職率"] = row["46期期首"].ToString() == "" ? 0 : row["46期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["46期退職"]) / Convert.ToDecimal(row["46期期首"]) * 100, 1);
                Disp.Rows.Add(nr46);

                DataRow nr45 = Disp.NewRow();
                nr45["期"] = "45期(2016年4月～)";
                nr45["期首従業員数"] = row["45期期首"];
                nr45["入職者数"] = row["45期入社"];
                nr45["退職者数"] = row["45期退職"];
                nr45["入職率"] = row["45期期首"].ToString() == "" ? 0 : row["45期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["45期入社"]) / Convert.ToDecimal(row["45期期首"]) * 100, 1);
                nr45["退職率"] = row["45期期首"].ToString() == "" ? 0 : row["45期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["45期退職"]) / Convert.ToDecimal(row["45期期首"]) * 100, 1);
                Disp.Rows.Add(nr45);

                DataRow nr44 = Disp.NewRow();
                nr44["期"] = "44期(2015年4月～)";
                nr44["期首従業員数"] = row["44期期首"];
                nr44["入職者数"] = row["44期入社"];
                nr44["退職者数"] = row["44期退職"];
                nr44["入職率"] = row["44期期首"].ToString() == "" ? 0 : row["44期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["44期入社"]) / Convert.ToDecimal(row["44期期首"]) * 100, 1);
                nr44["退職率"] = row["44期期首"].ToString() == "" ? 0 : row["44期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["44期退職"]) / Convert.ToDecimal(row["44期期首"]) * 100, 1);
                Disp.Rows.Add(nr44);

                DataRow nr43 = Disp.NewRow();
                nr43["期"] = "43期(2014年4月～)";
                nr43["期首従業員数"] = row["43期期首"];
                nr43["入職者数"] = row["43期入社"];
                nr43["退職者数"] = row["43期退職"];
                nr43["入職率"] = row["43期期首"].ToString() == "" ? 0 : row["43期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["43期入社"]) / Convert.ToDecimal(row["43期期首"]) * 100, 1);
                nr43["退職率"] = row["43期期首"].ToString() == "" ? 0 : row["43期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["43期退職"]) / Convert.ToDecimal(row["43期期首"]) * 100, 1);
                Disp.Rows.Add(nr43);

                DataRow nr42 = Disp.NewRow();
                nr42["期"] = "42期(2013年4月～)";
                nr42["期首従業員数"] = row["42期期首"];
                nr42["入職者数"] = row["42期入社"];
                nr42["退職者数"] = row["42期退職"];
                nr42["入職率"] = row["42期期首"].ToString() == "" ? 0 : row["42期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["42期入社"]) / Convert.ToDecimal(row["42期期首"]) * 100, 1);
                nr42["退職率"] = row["42期期首"].ToString() == "" ? 0 : row["42期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["42期退職"]) / Convert.ToDecimal(row["42期期首"]) * 100, 1);
                Disp.Rows.Add(nr42);

                DataRow nr41 = Disp.NewRow();
                nr41["期"] = "41期(2012年4月～)";
                nr41["期首従業員数"] = row["41期期首"];
                nr41["入職者数"] = row["41期入社"];
                nr41["退職者数"] = row["41期退職"];
                nr41["入職率"] = row["41期期首"].ToString() == "" ? 0 : row["41期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["41期入社"]) / Convert.ToDecimal(row["41期期首"]) * 100, 1);
                nr41["退職率"] = row["41期期首"].ToString() == "" ? 0 : row["41期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["41期退職"]) / Convert.ToDecimal(row["41期期首"]) * 100, 1);
                Disp.Rows.Add(nr41);

                DataRow nr40 = Disp.NewRow();
                nr40["期"] = "40期(2011年4月～)";
                nr40["期首従業員数"] = row["40期期首"];
                nr40["入職者数"] = row["40期入社"];
                nr40["退職者数"] = row["40期退職"];
                nr40["入職率"] = row["40期期首"].ToString() == "" ? 0 : row["40期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["40期入社"]) / Convert.ToDecimal(row["40期期首"]) * 100, 1);
                nr40["退職率"] = row["40期期首"].ToString() == "" ? 0 : row["40期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["40期退職"]) / Convert.ToDecimal(row["40期期首"]) * 100, 1);
                Disp.Rows.Add(nr40);

                DataRow nr39 = Disp.NewRow();
                nr39["期"] = "39期(2010年4月～)";
                nr39["期首従業員数"] = row["39期期首"];
                nr39["入職者数"] = row["39期入社"];
                nr39["退職者数"] = row["39期退職"];
                nr39["入職率"] = row["39期期首"].ToString() == "" ? 0 : row["39期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["39期入社"]) / Convert.ToDecimal(row["39期期首"]) * 100, 1);
                nr39["退職率"] = row["39期期首"].ToString() == "" ? 0 : row["39期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["39期退職"]) / Convert.ToDecimal(row["39期期首"]) * 100, 1);
                Disp.Rows.Add(nr39);

                DataRow nr38 = Disp.NewRow();
                nr38["期"] = "38期(2009年4月～)";
                nr38["期首従業員数"] = row["38期期首"];
                nr38["入職者数"] = row["38期入社"];
                nr38["退職者数"] = row["38期退職"];
                nr38["入職率"] = row["38期期首"].ToString() == "" ? 0 : row["38期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["38期入社"]) / Convert.ToDecimal(row["38期期首"]) * 100, 1);
                nr38["退職率"] = row["38期期首"].ToString() == "" ? 0 : row["38期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["38期退職"]) / Convert.ToDecimal(row["38期期首"]) * 100, 1);
                Disp.Rows.Add(nr38);

                DataRow nr37 = Disp.NewRow();
                nr37["期"] = "37期(2008年4月～)";
                nr37["期首従業員数"] = row["37期期首"];
                nr37["入職者数"] = row["37期入社"];
                nr37["退職者数"] = row["37期退職"];
                nr37["入職率"] = row["37期期首"].ToString() == "" ? 0 : row["37期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["37期入社"]) / Convert.ToDecimal(row["37期期首"]) * 100, 1);
                nr37["退職率"] = row["37期期首"].ToString() == "" ? 0 : row["37期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["37期退職"]) / Convert.ToDecimal(row["37期期首"]) * 100, 1);
                Disp.Rows.Add(nr37);

                DataRow nr36 = Disp.NewRow();
                nr36["期"] = "36期(2007年4月～)";
                nr36["期首従業員数"] = row["36期期首"];
                nr36["入職者数"] = row["36期入社"];
                nr36["退職者数"] = row["36期退職"];
                nr36["入職率"] = row["36期期首"].ToString() == "" ? 0 : row["36期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["36期入社"]) / Convert.ToDecimal(row["36期期首"]) * 100, 1);
                nr36["退職率"] = row["36期期首"].ToString() == "" ? 0 : row["36期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["36期退職"]) / Convert.ToDecimal(row["36期期首"]) * 100, 1);
                Disp.Rows.Add(nr36);

                DataRow nr35 = Disp.NewRow();
                nr35["期"] = "35期(2006年4月～)";
                nr35["期首従業員数"] = row["35期期首"];
                nr35["入職者数"] = row["35期入社"];
                nr35["退職者数"] = row["35期退職"];
                nr35["入職率"] = row["35期期首"].ToString() == "" ? 0 : row["35期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["35期入社"]) / Convert.ToDecimal(row["35期期首"]) * 100, 1);
                nr35["退職率"] = row["35期期首"].ToString() == "" ? 0 : row["35期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["35期退職"]) / Convert.ToDecimal(row["35期期首"]) * 100, 1);
                Disp.Rows.Add(nr35);

                DataRow nr34 = Disp.NewRow();
                nr34["期"] = "34期(2005年4月～)";
                nr34["期首従業員数"] = row["34期期首"];
                nr34["入職者数"] = row["34期入社"];
                nr34["退職者数"] = row["34期退職"];
                nr34["入職率"] = row["34期期首"].ToString() == "" ? 0 : row["34期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["34期入社"]) / Convert.ToDecimal(row["34期期首"]) * 100, 1);
                nr34["退職率"] = row["34期期首"].ToString() == "" ? 0 : row["34期期首"].ToString() == "0" ? 0 : Math.Round(Convert.ToDecimal(row["34期退職"]) / Convert.ToDecimal(row["34期期首"]) * 100, 1);
                Disp.Rows.Add(nr34);

                //decimal kisyu = 0;
                //decimal nyuu = 0;
                //decimal tai = 0;

                //kisyu = (Convert.ToDecimal(row["40期期首"]) + Convert.ToDecimal(row["41期期首"]) + Convert.ToDecimal(row["42期期首"]) + Convert.ToDecimal(row["43期期首"]) + Convert.ToDecimal(row["44期期首"]) + Convert.ToDecimal(row["45期期首"])) / 6;
                //nyuu = (Convert.ToDecimal(row["40期入社"]) + Convert.ToDecimal(row["41期入社"]) + Convert.ToDecimal(row["42期入社"]) + Convert.ToDecimal(row["43期入社"]) + Convert.ToDecimal(row["44期入社"]) + Convert.ToDecimal(row["45期入社"])) / 6;
                //tai = (Convert.ToDecimal(row["40期退職"]) + Convert.ToDecimal(row["41期退職"]) + Convert.ToDecimal(row["42期退職"]) + Convert.ToDecimal(row["43期退職"]) + Convert.ToDecimal(row["44期退職"]) + Convert.ToDecimal(row["45期退職"])) / 6;

                //DataRow nrav = Disp.NewRow();
                //nrav["期"] = "40期～45期平均";
                //nrav["期首従業員数"] = Math.Round(kisyu);
                //nrav["入職者数"] = Math.Round(nyuu);
                //nrav["退職者数"] = Math.Round(tai);
                //nrav["入職率"] = Math.Round(nyuu / kisyu * 100, 1);
                //nrav["退職率"] = Math.Round(tai / kisyu * 100, 1);
                //Disp.Rows.Add(nrav);
            }

            dataGridView5.DataSource = Disp;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //データ表示 iskw名誉会長対応
            DataView();
            GetDataBetsu();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //TODO 1907 絞り込めるはず
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            SyousaiData(drv[0].ToString());
        }

        private void dataGridView9_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0) return;　//最初行とばし
            if (e.ColumnIndex == 0) return; //日付とばし
            if (e.ColumnIndex == 1) return; //年齢とばし

            if (Convert.ToString(dataGridView9.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value) != Convert.ToString(e.Value))
            {
                //前がnullで後が0はとばし
                if (dataGridView9.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.Equals(DBNull.Value)) return;

                e.CellStyle.BackColor = Color.SpringGreen;
            }

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            checkBox6.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            checkedListBox2.Items.Clear();
            checkedListBox3.Items.Clear();
            checkedListBox4.Items.Clear();

            if (checkBox6.Checked)
            { 

                GetData("退職含");
            }
            else
            {
                GetData("");
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                checkedListBox4.SetItemChecked(i, true);
            }

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            checkBox6.Enabled = true;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox8.SelectedIndex == -1)
            {
                MessageBox.Show("対象年月選んでください");
                return;
            }
            if (syainNo.Text == "")
            {
                MessageBox.Show("対象者選んでください");
                return;
            }
            //pdf
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button6.Enabled = false;

            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s出勤簿データ取得_個別('" + comboBox8.SelectedItem.ToString().Replace("/", "") + "') where 社員番号 = '" + syainNo.Text + "'order by 組織名, 現場名, カナ名");

            bool flg = true;
            if (kubun.Text == "施設") flg = false;
            Com.GetSyukkinbo(dt, Convert.ToDateTime(comboBox8.SelectedItem.ToString() + "/01"), flg, false);

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button6.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox8.SelectedIndex == -1)
            {
                MessageBox.Show("対象年月選んでください");
                return;
            }
            if (syainNo.Text == "")
            {
                MessageBox.Show("対象者選んでください");
                return;
            }
            //pdf
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button7.Enabled = false;

            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s出勤簿データ取得_個別('" + comboBox8.SelectedItem.ToString().Replace("/", "") + "') where 社員番号 = '" + syainNo.Text + "'order by 組織名, 現場名, カナ名");

            bool flg = true;
            if (kubun.Text == "施設") flg = false;
            Com.GetSyukkinbo(dt, Convert.ToDateTime(comboBox8.SelectedItem.ToString() + "/01"), flg, true);

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button7.Enabled = true;
        }

        private void shikakudgv_SelectionChanged(object sender, EventArgs e)
        {
            ShikakuClear();

            DataGridViewRow dgr = shikakudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            GetShikakukoumokuData(drv[0].ToString());
        }

        private void ShikakuClear()
        {
            shikakucode.Text = "";
            shikakuname.Text = "";

            shikakusyutokubi.Value = null;
            shikakuno.Text = "";
            shikakukigenday.Value = null;
        }

        private void GetShikakukoumokuData(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainNo.Text + "' and 資格コード = '" + str + "' and 適用終了日 = '9999/12/31'");

            if (dt.Rows.Count == 0) return;

            shikakucode.Text = dt.Rows[0][0].ToString();
            shikakuname.Text = dt.Rows[0][1].ToString();

            //shikakusyutokubi.Text = dt.Rows[0][2].ToString();
            if (dt.Rows[0][2].ToString() == "")
            {
                shikakusyutokubi.Value = null;
            }
            else
            {
                shikakusyutokubi.Value = Convert.ToDateTime(dt.Rows[0][2].ToString());
            }



            shikakuno.Text = dt.Rows[0][3].ToString();

            //shikakukigenday.Text = dt.Rows[0][4].ToString();

            if (dt.Rows[0][4].ToString().Trim() == "" || dt.Rows[0][4] == DBNull.Value)
            {
                shikakukigenday.Value = null;
            }
            else
            {
                shikakukigenday.Value = Convert.ToDateTime(dt.Rows[0][4].ToString());
            }

            //if (dt.Rows[0][4].Equals(DBNull.Value) || dt.Rows[0][4].ToString() == "          " || dt.Rows[0][4].ToString() == "")
            //{
            //    shikakukigenday.Text = ""; //資格有効期限
            //    radioButton1.Checked = false;
            //}
            //else
            //{
            //    shikakukigenday.Text = dt.Rows[0][4].ToString(); //資格有効期限
            //    radioButton1.Checked = true;
            //}
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (shikakusyutokubi.Text == "")
            {
                MessageBox.Show("取得日は必須です。");
                return;
            }

            if (shikakuno.Text == "")
            {
                MessageBox.Show("資格取得番号は必須です。");
                return;
            }

            int ri = shikakudgv.CurrentCell.RowIndex;

            UpdateShikaku();
            GetShikaku();

            shikakudgv.CurrentCell = shikakudgv[1, ri];

            ShikakuClear();

            DataGridViewRow dgr = shikakudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            GetShikakukoumokuData(drv[0].ToString());


            MessageBox.Show("更新しましたー");
        }

        private void UpdateShikaku()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[s資格情報更新]";

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格コード", SqlDbType.VarChar)); Cmd.Parameters["資格コード"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格取得日", SqlDbType.VarChar)); Cmd.Parameters["資格取得日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格取得番号", SqlDbType.VarChar)); Cmd.Parameters["資格取得番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格有効期限", SqlDbType.Char)); Cmd.Parameters["資格有効期限"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar)); Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainNo.Text;
                    Cmd.Parameters["資格コード"].Value = shikakucode.Text;
                    Cmd.Parameters["資格取得日"].Value = Convert.ToDateTime(shikakusyutokubi.Value).ToString("yyyy/MM/dd");
                    Cmd.Parameters["資格取得番号"].Value = shikakuno.Text;

                    if (shikakukigenday.Text == "")
                    {
                        Cmd.Parameters["資格有効期限"].Value = DBNull.Value;
                    }
                    else
                    {
                        //Cmd.Parameters["免許証"].Value = menkyonew.Text;
                        Cmd.Parameters["資格有効期限"].Value = Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy/MM/dd");
                    }

                    Cmd.Parameters["更新者"].Value = Program.loginname;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void GetShikaku()
        {
            shikakudgv.DataSource = "";

            //資格手当情報
            DataTable dtshikaku = new DataTable();
            dtshikaku = Com.GetDB("select 資格コード, 資格名, 取得番号, 有効期限, 期限, 規程額, 手当額 as 手当対象額 from dbo.s資格一覧 where 社員番号 = '" + syainNo.Text + "'");

            shikakudgv.DataSource = dtshikaku;
            shikakudgv.Columns[5].DefaultCellStyle.Format = "#,0";
            shikakudgv.Columns[6].DefaultCellStyle.Format = "#,0";
            shikakudgv.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            shikakudgv.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void wareki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki.SelectedItem == null) return;

            //今日の西暦になってしまう！
            if (shikakukigenday.Text == "")
            {
                if (wareki.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    shikakukigenday.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                }

                if (wareki.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    shikakukigenday.Value = new DateTime(2019, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    return;
                }

                if (wareki.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        shikakukigenday.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        shikakukigenday.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                }
            }
            else
            {
                if (wareki.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    shikakukigenday.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("dd")));
                }

                if (wareki.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    shikakukigenday.Value = new DateTime(2019, Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("dd")));
                    return;
                }

                if (wareki.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        shikakukigenday.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        shikakukigenday.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("dd")));
                    }
                }
            }
        }

        private void shikakukigenday_ValueChanged(object sender, EventArgs e)
        {

            if (shikakukigenday.Text == "")
            {
                wareki.SelectedIndex = -1;
                return;
            }

            if (Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy")) > 1989)
            {
                wareki.SelectedIndex = wareki.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy")) == 2019)
            {
                wareki.SelectedIndex = wareki.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy")) > 2019)
            {
                wareki.SelectedIndex = wareki.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(shikakukigenday.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        private bool BeforeCheck()
        {
            //TODO 入社入力に同じ処理があります。。
            //必須項目が入ってない場合は出力できない処理
            string msg = "";
            if (keiyakunengetsu.Value.Equals(DBNull.Value)) msg += "・作成日" + Com.nl;
            if (koyoukubun.SelectedItem?.ToString() == "") msg += "・雇用区分" + Com.nl;
            //if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり" && koyoukaishibi.Value.Equals(DBNull.Value)) msg += "・雇用開始日" + Com.nl;
            if (koyoukaishibi.Value.Equals(DBNull.Value)) msg += "・雇用開始日" + Com.nl;
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり" && koyousyuuryoubi.Value.Equals(DBNull.Value)) msg += "・雇用終了日" + Com.nl;
            if (koushinkubun.SelectedItem?.ToString() == "") msg += "・更新区分" + Com.nl;
            if (gyoumunaiyou.Text == "") msg += "・業務内容" + Com.nl;
            if ((dtpe0.Value - dtps0.Value).Hours == 0 && (dtpe1.Value - dtps1.Value).Hours == 0) msg += "・勤務時間" + Com.nl;
            if (zikangairoudou.SelectedItem?.ToString() == "") msg += "・時間外労働" + Com.nl;
            if (yakankinmu.SelectedItem?.ToString() == "") msg += "・夜間勤務" + Com.nl;
            if (kyuujitsukaisuu.SelectedItem?.ToString() == "") msg += "・休日回数" + Com.nl;
            if (kyuusyustu.SelectedItem?.ToString() == "") msg += "・休出有無" + Com.nl;
            if (teinen.SelectedItem?.ToString() == "") msg += "・定年区分" + Com.nl;
            if (syouyo.SelectedItem?.ToString() == "") msg += "・賞与区分" + Com.nl;
            if (taisyokukin.SelectedItem?.ToString() == "") msg += "・退職金区分" + Com.nl;

            if (!koyoukaishibi.Value.Equals(DBNull.Value) && !koyousyuuryoubi.Value.Equals(DBNull.Value))
            {
                if (Convert.ToDateTime(koyoukaishibi.Value) > Convert.ToDateTime(koyousyuuryoubi.Value)) msg += "契約期間がおかしー" + Com.nl;
            }

            if (msg != "")
            {
                msg = "下記項目は必須です。" + Com.nl + msg;
                MessageBox.Show(msg);
                return true;
            }
            else
            {
                return false;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (syainNo.Text == "")
            {
                MessageBox.Show("誰も選択されてません。。"); 
                return;
            }

            if (BeforeCheck()) return;

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button7.Enabled = false;

            //更新して出力
            UpdateRoudou();

            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from r労働条件取得 where 社員番号 = '" + syainNo.Text + "'");

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            if (kyuuyokubuncode.Text == "E1" || kyuuyokubuncode.Text == "F1")
            {
                //パート・アルバイト
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\11_労働条件ファミリー.xlsx");
            }
            else
            {
                //TODO 日給者・役員・契約社員の対応が未
                //月給者
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\11_労働条件一般.xlsx");
            }



            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[j,i + 1].Value = dt.Rows[i][j].ToString();
                }
            }

            string localPass = @"C:\ODIS\ROUDOU\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");
            System.Diagnostics.Process.Start(exlName + ".xlsx");
           


            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button7.Enabled = true;
        }

        private void label88_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            //TODO 入社入力に同じ処理があります。。

            if (syainNo.Text == "")
            {
                MessageBox.Show("誰も選択されてません。。");
                return;
            }

            if (BeforeCheck()) return;

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            rowjkoushin.Enabled = false;

            UpdateRoudou();

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            rowjkoushin.Enabled = true;

            MessageBox.Show("更新しました。");
        }


        private void UpdateRoudou()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[r労働条件更新]";
                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("契約年月", SqlDbType.VarChar)); Cmd.Parameters["契約年月"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用区分", SqlDbType.Char)); Cmd.Parameters["雇用区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用開始日", SqlDbType.VarChar)); Cmd.Parameters["雇用開始日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用終了日", SqlDbType.VarChar)); Cmd.Parameters["雇用終了日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新区分", SqlDbType.Char)); Cmd.Parameters["更新区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("就業場所", SqlDbType.VarChar)); Cmd.Parameters["就業場所"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("業務内容", SqlDbType.VarChar)); Cmd.Parameters["業務内容"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定始業", SqlDbType.Time)); Cmd.Parameters["定始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定終業", SqlDbType.Time)); Cmd.Parameters["定終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定休憩", SqlDbType.VarChar)); Cmd.Parameters["定休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1始業", SqlDbType.Time)); Cmd.Parameters["S1始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1終業", SqlDbType.Time)); Cmd.Parameters["S1終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1休憩", SqlDbType.VarChar)); Cmd.Parameters["S1休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2始業", SqlDbType.Time)); Cmd.Parameters["S2始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2終業", SqlDbType.Time)); Cmd.Parameters["S2終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2休憩", SqlDbType.VarChar)); Cmd.Parameters["S2休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3始業", SqlDbType.Time)); Cmd.Parameters["S3始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3終業", SqlDbType.Time)); Cmd.Parameters["S3終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3休憩", SqlDbType.VarChar)); Cmd.Parameters["S3休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4始業", SqlDbType.Time)); Cmd.Parameters["S4始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4終業", SqlDbType.Time)); Cmd.Parameters["S4終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4休憩", SqlDbType.VarChar)); Cmd.Parameters["S4休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5始業", SqlDbType.Time)); Cmd.Parameters["S5始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5終業", SqlDbType.Time)); Cmd.Parameters["S5終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5休憩", SqlDbType.VarChar)); Cmd.Parameters["S5休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("時間外労働区分", SqlDbType.Char)); Cmd.Parameters["時間外労働区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("夜間勤務区分", SqlDbType.Char)); Cmd.Parameters["夜間勤務区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("休日回数", SqlDbType.VarChar)); Cmd.Parameters["休日回数"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("休出有無", SqlDbType.Char)); Cmd.Parameters["休出有無"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定年区分", SqlDbType.Char)); Cmd.Parameters["定年区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("賞与区分", SqlDbType.Char)); Cmd.Parameters["賞与区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("退職金区分", SqlDbType.Char)); Cmd.Parameters["退職金区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainNo.Text;
                    Cmd.Parameters["契約年月"].Value = keiyakunengetsu.Value;
                    Cmd.Parameters["雇用区分"].Value = koyoukubun.Text;
                    Cmd.Parameters["雇用開始日"].Value = koyoukaishibi.Value;
                    Cmd.Parameters["雇用終了日"].Value = koyousyuuryoubi.Value;
                    Cmd.Parameters["更新区分"].Value = koushinkubun.Text;
                    Cmd.Parameters["就業場所"].Value = syuugyoubasyo.Text;
                    Cmd.Parameters["業務内容"].Value = gyoumunaiyou.Text;
                    Cmd.Parameters["定始業"].Value = dtps0.Text;
                    Cmd.Parameters["定終業"].Value = dtpe0.Text;
                    Cmd.Parameters["定休憩"].Value = kyuukei0.Text;
                    Cmd.Parameters["S1始業"].Value = dtps1.Text;
                    Cmd.Parameters["S1終業"].Value = dtpe1.Text;
                    Cmd.Parameters["S1休憩"].Value = kyuukei1.Text;
                    Cmd.Parameters["S2始業"].Value = dtps2.Text;
                    Cmd.Parameters["S2終業"].Value = dtpe2.Text;
                    Cmd.Parameters["S2休憩"].Value = kyuukei2.Text;
                    Cmd.Parameters["S3始業"].Value = dtps3.Text;
                    Cmd.Parameters["S3終業"].Value = dtpe3.Text;
                    Cmd.Parameters["S3休憩"].Value = kyuukei3.Text;
                    Cmd.Parameters["S4始業"].Value = dtps4.Text;
                    Cmd.Parameters["S4終業"].Value = dtpe4.Text;
                    Cmd.Parameters["S4休憩"].Value = kyuukei4.Text;
                    Cmd.Parameters["S5始業"].Value = dtps5.Text;
                    Cmd.Parameters["S5終業"].Value = dtpe5.Text;
                    Cmd.Parameters["S5休憩"].Value = kyuukei5.Text;
                    Cmd.Parameters["時間外労働区分"].Value = zikangairoudou.Text;
                    Cmd.Parameters["夜間勤務区分"].Value = yakankinmu.Text;
                    Cmd.Parameters["休日回数"].Value = kyuujitsukaisuu.Text;
                    Cmd.Parameters["休出有無"].Value = kyuusyustu.Text;
                    Cmd.Parameters["定年区分"].Value = teinen.Text;
                    Cmd.Parameters["賞与区分"].Value = syouyo.Text;
                    Cmd.Parameters["退職金区分"].Value = taisyokukin.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        //MessageBox.Show("更新しました。");
                    }
                }
            }
        }

        private void dtp0_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe0.Value - dtps0.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps0.Value.Year, dtps0.Value.Month, dtps0.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps0.Value;

                TimeSpan ts2 = dtpe0.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei0.Text = "120分";
                    }
                    else
                    {
                        kyuukei0.Text = "60分";
                    }
                }
                else
                {
                    kyuukei0.Text = "45分";
                }
            }
            else
            {
                kyuukei0.Text = "";
            }

            //勤務時間追加
            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh0.Text = h.ToString();
            }
            else
            {
                kinmuh0.Text = "";
            }
        }

        private void dtp1_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe1.Value - dtps1.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps1.Value.Year, dtps1.Value.Month, dtps1.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps1.Value;

                TimeSpan ts2 = dtpe1.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei1.Text = "120分";
                    }
                    else
                    {
                        kyuukei1.Text = "60分";
                    }
                }
                else
                {
                    kyuukei1.Text = "45分";
                }
            }
            else
            {
                kyuukei1.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh1.Text = h.ToString();
            }
            else
            {
                kinmuh1.Text = "";
            }
        }

        private void dtp2_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe2.Value - dtps2.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;


            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps2.Value.Year, dtps2.Value.Month, dtps2.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps2.Value;

                TimeSpan ts2 = dtpe2.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei2.Text = "120分";
                    }
                    else
                    {
                        kyuukei2.Text = "60分";
                    }
                }
                else
                {
                    kyuukei2.Text = "45分";
                }
            }
            else
            {
                kyuukei2.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh2.Text = h.ToString();
            }
            else
            {
                kinmuh2.Text = "";
            }
        }

        private void dtp3_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe3.Value - dtps3.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps3.Value.Year, dtps3.Value.Month, dtps3.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps3.Value;

                TimeSpan ts2 = dtpe3.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei3.Text = "120分";
                    }
                    else
                    {
                        kyuukei3.Text = "60分";
                    }
                }
                else
                {
                    kyuukei3.Text = "45分";
                }
            }
            else
            {
                kyuukei3.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh3.Text = h.ToString();
            }
            else
            {
                kinmuh3.Text = "";
            }
        }

        private void dtp4_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe4.Value - dtps4.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps4.Value.Year, dtps4.Value.Month, dtps4.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps4.Value;

                TimeSpan ts2 = dtpe4.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei4.Text = "120分";
                    }
                    else
                    {
                        kyuukei4.Text = "60分";
                    }
                }
                else
                {
                    kyuukei4.Text = "45分";
                }
            }
            else
            {
                kyuukei4.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh4.Text = h.ToString();
            }
            else
            {
                kinmuh4.Text = "";
            }
        }

        private void dtp5_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe5.Value - dtps5.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps5.Value.Year, dtps5.Value.Month, dtps5.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps5.Value;

                TimeSpan ts2 = dtpe5.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei5.Text = "120分";
                    }
                    else
                    {
                        kyuukei5.Text = "60分";
                    }
                }
                else
                {
                    kyuukei5.Text = "45分";
                }
            }
            else
            {
                kyuukei5.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh5.Text = h.ToString();
            }
            else
            {
                kinmuh5.Text = "";
            }
        }

        private void koyoukubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり")
            {
                //koyoukaishibi.Visible = true;
                koyousyuuryoubi.Visible = true;
            }
            else
            {
                //koyoukaishibi.Value = null;
                //koyoukaishibi.Visible = false;

                koyousyuuryoubi.Value = null;
                koyousyuuryoubi.Visible = false;
            }
        }

        private void label122_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click_1(object sender, EventArgs e)
        {

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[k緊急連絡先更新]";

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("携帯番号", SqlDbType.VarChar)); Cmd.Parameters["携帯番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("固定電話", SqlDbType.VarChar)); Cmd.Parameters["固定電話"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名1", SqlDbType.VarChar)); Cmd.Parameters["氏名1"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("カナ名1", SqlDbType.VarChar)); Cmd.Parameters["カナ名1"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("続柄1", SqlDbType.VarChar)); Cmd.Parameters["続柄1"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("電話番号1", SqlDbType.VarChar)); Cmd.Parameters["電話番号1"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名2", SqlDbType.VarChar)); Cmd.Parameters["氏名2"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("カナ名2", SqlDbType.VarChar)); Cmd.Parameters["カナ名2"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("続柄2", SqlDbType.VarChar)); Cmd.Parameters["続柄2"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("電話番号2", SqlDbType.VarChar)); Cmd.Parameters["電話番号2"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("最終更新日", SqlDbType.DateTime)); Cmd.Parameters["最終更新日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("最終更新者", SqlDbType.VarChar)); Cmd.Parameters["最終更新者"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainNo.Text;

                    Cmd.Parameters["携帯番号"].Value = honkeitai.Text;
                    Cmd.Parameters["固定電話"].Value = honkotei.Text;

                    Cmd.Parameters["氏名1"].Value = kaz1name.Text;
                    Cmd.Parameters["カナ名1"].Value = kaz1kana.Text;
                    Cmd.Parameters["続柄1"].Value = kaz1gara.Text;
                    Cmd.Parameters["電話番号1"].Value = kaz1no.Text;

                    Cmd.Parameters["氏名2"].Value = kaz2name.Text;
                    Cmd.Parameters["カナ名2"].Value = kaz2kana.Text;
                    Cmd.Parameters["続柄2"].Value = kaz2gara.Text;
                    Cmd.Parameters["電話番号2"].Value = kaz2no.Text;

                    Cmd.Parameters["最終更新日"].Value = DateTime.Now;
                    Cmd.Parameters["最終更新者"].Value = Program.loginname;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);

                        //MessageBox.Show("更新しましたー");

                        GetKinkyuu(syainNo.Text);
                    }
                }
            }
        }


        private void ValidCom(Control cont, CancelEventArgs e, string str, string reg)
        {
            //初回エラークリア
            this.errorProvider1.SetError(cont, "");

            if (cont.Text == "") return;

            //全角文字を半角文字に変換
            string result = Microsoft.VisualBasic.Strings.StrConv(cont.Text, VbStrConv.Narrow, 0).Trim();

            cont.Text = result;

            //電話番号以外除去
            Regex regex = new Regex(reg);
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError(cont, str + "が正しくありません。２つの-(ハイフン)は必須です。");
                e.Cancel = true;
                return;
            }

            //ハイフン除去
            //cont.Text = result.Replace("-", "");
        }

        private void honkeitai_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "本人携帯電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void honkotei_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "本人固定電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kaz1no_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "ご家族優先1電話番号", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kaz2no_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "ご家族優先2電話番号", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kaz1kana_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角→半角　ひらがな→カタカナへ
            string result = Microsoft.VisualBasic.Strings.StrConv(Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Katakana), VbStrConv.Narrow).Trim();

            //カナと空白以外はエラー
            Regex regex = new Regex(@"^[ｦ-ﾝﾞﾟ ]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "カタカナ以外は入力しないでください");
                e.Cancel = true;
            }

            ((TextBox)sender).Text = result;
        }

        private void kaz2kana_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角→半角　ひらがな→カタカナへ
            string result = Microsoft.VisualBasic.Strings.StrConv(Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Katakana), VbStrConv.Narrow).Trim();

            //カナと空白以外はエラー
            Regex regex = new Regex(@"^[ｦ-ﾝﾞﾟ ]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "カタカナ以外は入力しないでください");
                e.Cancel = true;
            }

            ((TextBox)sender).Text = result;
        }

        private void label130_Click(object sender, EventArgs e)
        {

        }

        private void taisyokubi_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
