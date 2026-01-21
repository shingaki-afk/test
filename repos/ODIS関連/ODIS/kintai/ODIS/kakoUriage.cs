using System;
using System.Data;
using System.Windows.Forms;
using Npgsql;
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using System.Drawing;
using System.Runtime.InteropServices;
using ODIS.ODIS;

namespace ODIS
{
    public partial class kakouriage : Form
    {
        public kakouriage()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //コンボボックス初期設定
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt2 = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select top 1 * from dbo.過去売上";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            comboBox6.Items.Add("全対象");
            comboBox7.Items.Add("全対象");
            foreach (DataColumn col in dt2.Columns)
            {
                if (col.ColumnName == "reskey") break;
                comboBox6.Items.Add(col.ColumnName);
                comboBox7.Items.Add(col.ColumnName);
            }


            #region チェックリストボックスの初期設定
            checkedListBox1.Items.Add("那覇");
            checkedListBox1.Items.Add("八重山");
            checkedListBox1.Items.Add("北部");

            checkedListBox2.Items.Add("現業");
            checkedListBox2.Items.Add("技術企画");
            checkedListBox2.Items.Add("マンション");

            checkedListBox2.Items.Add("客室");
            checkedListBox2.Items.Add("サービス");
            checkedListBox2.Items.Add("総括");
            checkedListBox2.Items.Add("食堂");
            checkedListBox2.Items.Add("行雲");

            checkedListBox2.Items.Add("指定管理");

            checkedListBox2.Items.Add("エンジ１課");
            checkedListBox2.Items.Add("エンジ２課");
            checkedListBox2.Items.Add("エンジ３課");

            checkedListBox2.Items.Add("施設");
            checkedListBox2.Items.Add("施運*");

            checkedListBox2.Items.Add("米軍施設");
            checkedListBox2.Items.Add("米軍プロジェクト");
            checkedListBox2.Items.Add("警備");
            checkedListBox2.Items.Add("機械警備");

            checkedListBox3.Items.Add("契約固定");
            checkedListBox3.Items.Add("臨時");
            checkedListBox3.Items.Add("契約変動");
            checkedListBox3.Items.Add("物品");
            checkedListBox3.Items.Add("自販機");
            checkedListBox3.Items.Add("行雲");
            checkedListBox3.Items.Add("食堂");
            checkedListBox3.Items.Add("朝食");

            checkedListBox4.Items.Add("自社");
            checkedListBox4.Items.Add("外注");

            checkedListBox5.Items.Add("売上");
            checkedListBox5.Items.Add("取消");
            checkedListBox5.Items.Add("値引");
            checkedListBox5.Items.Add("返品");

            checkedListBox6.Items.Add("2005年");
            checkedListBox6.Items.Add("2006年");
            checkedListBox6.Items.Add("2007年");
            checkedListBox6.Items.Add("2008年");
            checkedListBox6.Items.Add("2009年");
            checkedListBox6.Items.Add("2010年");
            checkedListBox6.Items.Add("2011年");
            checkedListBox6.Items.Add("2012年");

            checkedListBox7.Items.Add("01月");
            checkedListBox7.Items.Add("02月");
            checkedListBox7.Items.Add("03月");
            checkedListBox7.Items.Add("04月");
            checkedListBox7.Items.Add("05月");
            checkedListBox7.Items.Add("06月");
            checkedListBox7.Items.Add("07月");
            checkedListBox7.Items.Add("08月");
            checkedListBox7.Items.Add("09月");
            checkedListBox7.Items.Add("10月");
            checkedListBox7.Items.Add("11月");
            checkedListBox7.Items.Add("12月");

            #endregion

            //初期値設定
            Clear();
        }

        private void Clear()
        {
            //文字列で絞込
            textBox1.Text = "";
            textBox2.Text = "";

            //検索列指定
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;

            //一列表示
            checkBox1.Checked = false;
            
            //コード表示
            checkBox2.Checked = false;
            checkBox2.Visible = false;

            //期間表示
            checkBox3.Checked = false;

            dateTimePicker1.Value = new DateTime(2005, 08, 01);
            dateTimePicker2.Value = new DateTime(2012, 12, 31);
            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;
            label12.Visible = false;

            //金額
            checkBox4.Checked = false;

            //
            textBox3.Text = "";
            textBox4.Text = "";

            textBox3.Visible = false;
            textBox4.Visible = false;
            label9.Visible = false;
            label11.Visible = false;

            //チェックリスト
            #region チェックリストボックスのチェックを入れる
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

            for (int i = 0; i < checkedListBox5.Items.Count; i++)
            {
                checkedListBox5.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox6.Items.Count; i++)
            {
                checkedListBox6.SetItemChecked(i, true);
            }

            for (int i = 0; i < checkedListBox7.Items.Count; i++)
            {
                checkedListBox7.SetItemChecked(i, true);
            }
            #endregion
        }



        private void button1_Click(object sender, EventArgs e)
        {         
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ処理
            ViewKakoUriage();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        //過去の売上データを表示
        private void ViewKakoUriage()
        {

            #region チェックボックスが１つ以上入っているかチェック
            string errStr = "";
            if (checkedListBox1.CheckedItems.Count == 0) errStr += "【地区名】\n";
            if (checkedListBox2.CheckedItems.Count == 0) errStr += "【職種名】\n";
            if (checkedListBox3.CheckedItems.Count == 0) errStr += "【売上名】\n";
            if (checkedListBox4.CheckedItems.Count == 0) errStr += "【作業名】\n";
            if (checkedListBox5.CheckedItems.Count == 0) errStr += "【取引名】\n";
            if (checkedListBox6.CheckedItems.Count == 0) errStr += "【年度】\n";
            if (checkedListBox7.CheckedItems.Count == 0) errStr += "【月度】\n";

            if (errStr != "")
            {
                MessageBox.Show("必ず一個以上はチェックを入れてください。\n入ってないグループは下記です。\n" + errStr);
                return;
            }
            #endregion

            #region 文字列検索
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    if (comboBox6.SelectedIndex == 0)
                    {
                        result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                    else
                    {
                        result += " and (" + comboBox6.SelectedItem.ToString() + " like '%" + s + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Com.isOneByteChar(s) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or " + comboBox6.SelectedItem.ToString() + " like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }
            }

            //test 除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                foreach (string s in ar2)
                {
                    if (comboBox7.SelectedIndex == 0)
                    {
                        result += " and (reskey not like '%" + s + "%' and reskey not like '%" + Com.isOneByteChar(s) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                    else
                    {
                        result += " and (" + comboBox7.SelectedItem.ToString() + " not like '%" + s + "%' and " + comboBox7.SelectedItem.ToString() + " not like '%" + Com.isOneByteChar(s) + "%' and " + comboBox7.SelectedItem.ToString() + " not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and " + comboBox7.SelectedItem.ToString() + " not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and " + comboBox7.SelectedItem.ToString() + " not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and " + comboBox7.SelectedItem.ToString() + " not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }
            }
            #endregion

            #region チェックボックスの絞込文字列を設定
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    result += " and 地区名 <> '" + checkedListBox1.Items[i].ToString() + "'";
                }
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    if (checkedListBox2.Items[i].ToString() == "施運*")
                    {
                        result += " and 職種名 not like '" + checkedListBox2.Items[i].ToString().Replace("*","") + "%'";
                    }
                    else
                    {
                        result += " and 職種名 <> '" + checkedListBox2.Items[i].ToString() + "'";
                    }
                }
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    result += " and 売上名 <> '" + checkedListBox3.Items[i].ToString() + "'";
                }
            }

            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                if (!checkedListBox4.GetItemChecked(i))
                {
                    result += " and 作業名 <> '" + checkedListBox4.Items[i].ToString() + "'";
                }
            }

            for (int i = 0; i < checkedListBox5.Items.Count; i++)
            {
                if (!checkedListBox5.GetItemChecked(i))
                {
                    result += " and 取引名 <> '" + checkedListBox5.Items[i].ToString() + "'";
                }
            }

            if (!checkBox3.Checked)
            {
                for (int i = 0; i < checkedListBox6.Items.Count; i++)
                {
                    if (!checkedListBox6.GetItemChecked(i))
                    {
                        result += " and 年月度 not like '" + checkedListBox6.Items[i].ToString().Replace("年", "") + "%'";
                    }
                }

                for (int i = 0; i < checkedListBox7.Items.Count; i++)
                {
                    if (!checkedListBox7.GetItemChecked(i))
                    {
                        result += " and 年月度 not like '%" + checkedListBox7.Items[i].ToString().Replace("月", "") + "'";
                    }
                }
            }
            else
            {
                string sDate = dateTimePicker1.Value.ToString("yyyyMM");
                string eDate = dateTimePicker2.Value.ToString("yyyyMM");
                result += " and 年月度 >= " + sDate + " and 年月度 <= " + eDate;
            }
            #endregion

            //金額
            if (checkBox4.Checked)
            {
                result += " and 金額 >= '" + textBox3.Text + "' and 金額 < '" + textBox4.Text + "'";
            }

            //先頭が「and」の場合、削除
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            if (result.Length > 0)
            {
                result = " where " + result;
            }

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.過去売上" + result;
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

            comboBox6.Items.Add("全対象");
            comboBox7.Items.Add("全対象");
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "reskey") break;
                comboBox6.Items.Add(col.ColumnName);
                comboBox7.Items.Add(col.ColumnName);
            }

            int ct = dt.Rows.Count;

            //検索履歴登録
            Com.InHistory("12_旧売上検索", result, ct.ToString());

            DataTable Disp = new DataTable();

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            decimal uriageAll = 0;

            #region データ加工

            if (checkBox2.Checked)
            {
                if (checkBox1.Checked)
                {
                    Disp.Columns.Add("年月度", typeof(string));
                    Disp.Columns.Add("地区CD", typeof(string));
                    Disp.Columns.Add("地区名", typeof(string));
                    Disp.Columns.Add("職種CD", typeof(string));
                    Disp.Columns.Add("職種名", typeof(string));
                    Disp.Columns.Add("売上CD", typeof(string));
                    Disp.Columns.Add("売上名", typeof(string));
                    Disp.Columns.Add("作業CD", typeof(string));
                    Disp.Columns.Add("作業名", typeof(string));
                    Disp.Columns.Add("取引CD", typeof(string));
                    Disp.Columns.Add("取引名", typeof(string));
                    Disp.Columns.Add("受託物件CD", typeof(string));
                    Disp.Columns.Add("受託物件名", typeof(string));
                    Disp.Columns.Add("業務CD", typeof(string));
                    Disp.Columns.Add("業務名", typeof(string));
                    Disp.Columns.Add("摘要", typeof(string));
                    Disp.Columns.Add("数量", typeof(decimal));
                    Disp.Columns.Add("単位名", typeof(string));
                    Disp.Columns.Add("単価", typeof(decimal));
                    Disp.Columns.Add("金額", typeof(decimal));
                    Disp.Columns.Add("消費税額", typeof(decimal));
                    Disp.Columns.Add("税区分名", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["年月度"] = row["年月度"];
                        nr["地区CD"] = row["地区CD"];
                        nr["地区名"] = row["地区名"];
                        nr["職種CD"] = row["職種CD"];
                        nr["職種名"] = row["職種名"];
                        nr["売上CD"] = row["売上CD"];
                        nr["売上名"] = row["売上名"];
                        nr["作業CD"] = row["作業CD"];
                        nr["作業名"] = row["作業名"];
                        nr["取引CD"] = row["取引CD"];
                        nr["取引名"] = row["取引名"];
                        nr["受託物件CD"] = row["受託物件CD"];
                        nr["受託物件名"] = row["受託物件名"];
                        nr["業務CD"] = row["業務CD"];
                        nr["業務名"] = row["業務名"];
                        nr["摘要"] = row["摘要"];
                        nr["数量"] = Convert.ToDecimal(row["数量"]);
                        nr["単位名"] = row["単位名"];
                        nr["単価"] = Convert.ToDecimal(row["単価"]);
                        nr["金額"] = Convert.ToDecimal(row["金額"]);
                        nr["消費税額"] = Convert.ToDecimal(row["消費税額"]);
                        nr["税区分名"] = row["税区分名"];
                        Disp.Rows.Add(nr);

                        if (row["取引名"].ToString() == "売上")
                        {
                            uriageAll += Convert.ToDecimal(row["金額"]);
                        }
                        else
                        {
                            uriageAll -= Convert.ToDecimal(row["金額"]);
                        }
                    }
                }
                else
                {
                    Disp.Columns.Add("年月度\n地区名\n職種名", typeof(string));
                    Disp.Columns.Add("売上名\n作業名\n取引名", typeof(string));
                    Disp.Columns.Add("受託物件名\n業務名\n摘要", typeof(string));
                    Disp.Columns.Add("数量\n単位名\n単価", typeof(string));
                    Disp.Columns.Add("金額\n消費税額\n税区分名", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["年月度\n地区名\n職種名"] = row["年月度"] + "\n" + row["地区名"] + "\n" + row["職種名"];
                        nr["売上名\n作業名\n取引名"] = row["売上名"] + "\n" + row["作業名"] + "\n" + row["取引名"];
                        nr["受託物件名\n業務名\n摘要"] = row["受託物件名"] + "\n" + row["業務名"] + "\n" + row["摘要"];
                        nr["数量\n単位名\n単価"] = Convert.ToDecimal(row["数量"]).ToString("#,0") + "\n" + row["単位名"] + "\n" + Convert.ToDecimal(row["単価"]).ToString("#,0");
                        nr["金額\n消費税額\n税区分名"] = Convert.ToDecimal(row["金額"]).ToString("#,0") + "\n" + Convert.ToDecimal(row["消費税額"]).ToString("#,0") + "\n" + row["税区分名"];
                        Disp.Rows.Add(nr);

                        if (row["取引名"].ToString() == "売上")
                        {
                            uriageAll += Convert.ToDecimal(row["金額"]);
                        }
                        else
                        {
                            uriageAll -= Convert.ToDecimal(row["金額"]);
                        }
                    }
                }
            }
            else
            {
                if (checkBox1.Checked)
                {
                    Disp.Columns.Add("年月度", typeof(string));
                    Disp.Columns.Add("地区名", typeof(string));
                    Disp.Columns.Add("職種名", typeof(string));
                    Disp.Columns.Add("売上名", typeof(string));
                    Disp.Columns.Add("作業名", typeof(string));
                    Disp.Columns.Add("取引名", typeof(string));
                    Disp.Columns.Add("受託物件名", typeof(string));
                    Disp.Columns.Add("業務名", typeof(string));
                    Disp.Columns.Add("摘要", typeof(string));
                    Disp.Columns.Add("数量", typeof(decimal));
                    Disp.Columns.Add("単位名", typeof(string));
                    Disp.Columns.Add("単価", typeof(decimal));
                    Disp.Columns.Add("金額", typeof(decimal));
                    Disp.Columns.Add("消費税額", typeof(decimal));
                    Disp.Columns.Add("税区分名", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["年月度"] = row["年月度"];
                        nr["地区名"] = row["地区名"];
                        nr["職種名"] = row["職種名"];
                        nr["売上名"] = row["売上名"];
                        nr["作業名"] = row["作業名"];
                        nr["取引名"] = row["取引名"];
                        nr["受託物件名"] = row["受託物件名"];
                        nr["業務名"] = row["業務名"];
                        nr["摘要"] = row["摘要"];
                        nr["数量"] = Convert.ToDecimal(row["数量"]);
                        nr["単位名"] = row["単位名"];
                        nr["単価"] = Convert.ToDecimal(row["単価"]);
                        nr["金額"] = Convert.ToDecimal(row["金額"]);
                        nr["消費税額"] = Convert.ToDecimal(row["消費税額"]);
                        nr["税区分名"] = row["税区分名"];
                        Disp.Rows.Add(nr);

                        if (row["取引名"].ToString() == "売上")
                        {
                            uriageAll += Convert.ToDecimal(row["金額"]);
                        }
                        else
                        {
                            uriageAll -= Convert.ToDecimal(row["金額"]);
                        }
                    }
                }
                else
                {
                    Disp.Columns.Add("年月度\n地区名\n職種名", typeof(string));
                    Disp.Columns.Add("売上名\n作業名\n取引名", typeof(string));
                    Disp.Columns.Add("受託物件名\n業務名\n摘要", typeof(string));
                    Disp.Columns.Add("数量\n単位名\n単価", typeof(string));
                    Disp.Columns.Add("金額\n消費税額\n税区分名", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["年月度\n地区名\n職種名"] = row["年月度"] + "\n" + row["地区名"] + "\n" + row["職種名"];
                        nr["売上名\n作業名\n取引名"] = row["売上名"] + "\n" + row["作業名"] + "\n" + row["取引名"];
                        nr["受託物件名\n業務名\n摘要"] = row["受託物件名"] + "\n" + row["業務名"] + "\n" + row["摘要"];
                        nr["数量\n単位名\n単価"] = Convert.ToDecimal(row["数量"]).ToString("#,0") + "\n" + row["単位名"] + "\n" + Convert.ToDecimal(row["単価"]).ToString("#,0");
                        nr["金額\n消費税額\n税区分名"] = Convert.ToDecimal(row["金額"]).ToString("#,0") + "\n" + Convert.ToDecimal(row["消費税額"]).ToString("#,0") + "\n" + row["税区分名"];
                        Disp.Rows.Add(nr);

                        if (row["取引名"].ToString() == "売上")
                        {
                            uriageAll += Convert.ToDecimal(row["金額"]);
                        }
                        else
                        {
                            uriageAll -= Convert.ToDecimal(row["金額"]);
                        }
                    }
                }
            }

#endregion

            //データグリッドビューの高さ指定　※セット前にすること！
            if (checkBox1.Checked)
            {
                dataGridView1.RowTemplate.Height = 20;
            }
            else
            {
                dataGridView1.RowTemplate.Height = 50;
            }


            label1.Text = ct.ToString() + " 件";
            label10.Text = "\\" + uriageAll.ToString("#,0") + "円";
            dataGridView1.DataSource = Disp;

                if (checkBox1.Checked)
                {
                    //一列表示

                    // セル内で文字列を折り返す
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

                    // 文字列全体が表示されるように行の幅を自動調節する
                    //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                    ////金額右寄せ
                    dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView1.Columns[0].Width = 100;
                    dataGridView1.Columns[1].Width = 100;
                    dataGridView1.Columns[2].Width = 100;
                    dataGridView1.Columns[3].Width = 100;
                    dataGridView1.Columns[4].Width = 100;
                    dataGridView1.Columns[5].Width = 100;
                    dataGridView1.Columns[6].Width = 250;
                    dataGridView1.Columns[7].Width = 250;
                    dataGridView1.Columns[8].Width = 100;
                    dataGridView1.Columns[9].Width = 100;
                    dataGridView1.Columns[10].Width =100;
                    dataGridView1.Columns[11].Width =100;
                    dataGridView1.Columns[12].Width =100;
                    dataGridView1.Columns[13].Width =100;
                    dataGridView1.Columns[14].Width =100;
                }
                else
                {
                    //複数列表示

                    // セル内で文字列を折り返す
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                    //// 文字列全体が表示されるように行の高さを自動調節する
                    //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

                    //// 文字列全体が表示されるように行の幅を自動調節する
                    //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    //////金額右寄せ
                    dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView1.Columns[0].Width = 100;
                    dataGridView1.Columns[1].Width = 100;
                    dataGridView1.Columns[2].Width = 300;
                    dataGridView1.Columns[3].Width = 100;
                    dataGridView1.Columns[4].Width = 100;
                }
            //}
            
            System.GC.Collect();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //ボタン無効化・カーソル変更
                button1.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                //データ処理
                ViewKakoUriage();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }


        private void label7_Click(object sender, EventArgs e)
        {
            if (checkedListBox6.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox6.Items.Count; i++)
                {
                    checkedListBox6.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox6.Items.Count; i++)
                {
                    checkedListBox6.SetItemChecked(i, false);
                }
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            if (checkedListBox7.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox7.Items.Count; i++)
                {
                    checkedListBox7.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox7.Items.Count; i++)
                {
                    checkedListBox7.SetItemChecked(i, false);
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
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

        private void label3_Click(object sender, EventArgs e)
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

        private void label4_Click(object sender, EventArgs e)
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

        private void label5_Click(object sender, EventArgs e)
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

        private void label6_Click(object sender, EventArgs e)
        {
            if (checkedListBox5.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox5.Items.Count; i++)
                {
                    checkedListBox5.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox5.Items.Count; i++)
                {
                    checkedListBox5.SetItemChecked(i, false);
                }
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void label11_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Value = new DateTime(2005, 08, 01);
            dateTimePicker2.Value = new DateTime(2012, 12, 31);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked)
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                label12.Visible = false;

                checkedListBox6.Enabled = true;
                checkedListBox7.Enabled = true;
                checkedListBox6.Visible = true;
                checkedListBox7.Visible = true;
                label7.Visible = true;
                label8.Visible = true;

            }
            else
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                label12.Visible = true;


                checkedListBox6.Enabled = false;
                checkedListBox7.Enabled = false;
                checkedListBox6.Visible = false;
                checkedListBox7.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
            }
        }

        private void kakouriage_FormClosed(object sender, FormClosedEventArgs e)
        {
            //System.GC.Collect();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Checked = false;
                checkBox2.Visible = true;
            }
            else
            {
                checkBox2.Checked = false;
                checkBox2.Visible = false;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                //押されたキーが 0～9でない場合は、イベントをキャンセルする
                e.Handled = true;

                MessageBox.Show("半角数字で入力ください。");
            }
            else
            {

            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                textBox3.Visible = true;
                textBox4.Visible = true;
                label9.Visible = true;
                label11.Visible = true;
            }
            else
            {
                textBox3.Visible = false;
                textBox4.Visible = false;
                label9.Visible = false;
                label11.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
        }
    }
}
