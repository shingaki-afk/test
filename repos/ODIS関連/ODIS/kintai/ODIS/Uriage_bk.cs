using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Npgsql;
using Microsoft.VisualBasic;
using System.Data.SqlClient;

namespace ODIS.ODIS
{
    public partial class Uriage_bk : Form
    {
        //private DataTable dt = new DataTable();

        public Uriage_bk()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //データ取得
            //GetUriageData();

            //初期値設定
            Clear();
        }

        private void Clear()
        {
            //文字列で絞込
            textBox1.Text = "";
            textBox2.Text = "";

            //一列表示
            checkBox1.Checked = false;

            //コード表示
            checkBox2.Checked = false;

            //日付設定
            DateTime today = DateTime.Today;
            dateTimePicker1.Value = today.AddMonths(-1).AddDays(-today.Day + 1);
            dateTimePicker2.Value = today.AddDays(-today.Day);

            //金額
            checkBox4.Checked = false;

            textBox3.Text = "";
            textBox4.Text = "";

            textBox3.Visible = false;
            textBox4.Visible = false;
            label9.Visible = false;
            label11.Visible = false;
        }

        private void GetUriageData()
        {
            //処理速度計算
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (キー like '%" + s + "%' or キー like '%" + Com.isOneByteChar(s) + "%' or キー like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or キー like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or キー like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or キー like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            //除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                foreach (string s in ar2)
                {
                    result += " and (キー not like '%" + s + "%' and キー not like '%" + Com.isOneByteChar(s) + "%' and キー not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and キー not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and キー not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and キー not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //金額
            if (checkBox4.Checked)
            {
                result += " and 売上額 >= '" + textBox3.Text + "' and 売上額 < '" + textBox4.Text + "'";
            }

            string sDate = dateTimePicker1.Value.ToString("yyyyMM");
            string eDate = dateTimePicker2.Value.ToString("yyyyMM");


            int nRet;

            NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr);
            conn.Open();
            DataTable dt = new DataTable();

            string sql = "select * from kpcp01.\"GetUriageDataSearchAll\" where 売上年月 between '" + sDate + "' and '" + eDate + "'" + result;

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
            nRet = adapter.Fill(dt);

            conn.Close();

            //解放
            adapter.Dispose();
            conn.Dispose();


            DataTable Disp = new DataTable();

            decimal uriageAll = 0;

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //列表示
            if (checkBox1.Checked)
            {
                //コード表示
                if (checkBox2.Checked)
                {
                    //一列表示
                    //コード表示

                    Disp.Columns.Add("売上年月", typeof(string));
                    Disp.Columns.Add("ﾚｺｰﾄﾞ番号", typeof(string));
                    Disp.Columns.Add("契約区分名", typeof(string));
                    Disp.Columns.Add("作業区分名", typeof(string));
                    Disp.Columns.Add("部門CD", typeof(string));
                    Disp.Columns.Add("部門名", typeof(string));

                    Disp.Columns.Add("取引先CD", typeof(string));
                    Disp.Columns.Add("取引先名", typeof(string));
                    Disp.Columns.Add("契約項目", typeof(string));
                    Disp.Columns.Add("工事CD", typeof(string));
                    Disp.Columns.Add("工事名", typeof(string));
                    Disp.Columns.Add("担当名", typeof(string));

                    //Disp.Columns.Add("締め日", typeof(string));
                    //Disp.Columns.Add("請求日", typeof(string));
                    Disp.Columns.Add("入力日", typeof(string));

                    Disp.Columns.Add("売上額", typeof(string));
                    Disp.Columns.Add("実施額", typeof(string));
                    Disp.Columns.Add("消費税額", typeof(string));

                    Disp.Columns.Add("数量", typeof(string));
                    Disp.Columns.Add("単価", typeof(string));
                    Disp.Columns.Add("税区分", typeof(string));

                    Disp.Columns.Add("支払業者CD", typeof(string));
                    Disp.Columns.Add("支払業者名", typeof(string));
                    Disp.Columns.Add("支払額", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月"] = row["売上年月"];
                        nr["ﾚｺｰﾄﾞ番号"] = row["レコード番号"];
                        nr["契約区分名"] = row["契約区分名称"];
                        nr["作業区分名"] = row["作業区分名称"];
                        nr["部門CD"] = row["部門コード"];
                        nr["部門名"] = row["部門名称"];

                        nr["取引先CD"] = row["取引先コード"];
                        nr["取引先名"] = row["取引先名称（正式名称）"];
                        nr["契約項目"] = row["契約項目"];
                        nr["工事CD"] = row["工事コード"] + "-" + row["工事枝コード"];
                        nr["工事名"] = row["工事名称（正式名称）"];
                        nr["担当名"] = row["担当者氏名"];

                        //nr["締め日"] = row["締め日"];
                        //nr["請求日"] = row["請求日"];
                        nr["入力日"] = row["売上入力日"];

                        nr["売上額"] = string.Format("{0:#,##0}", row["売上額"]);
                        nr["実施額"] = string.Format("{0:#,##0}", row["実施額"]);
                        nr["消費税額"] = string.Format("{0:#,##0}", row["消費税額"]);

                        nr["数量"] = string.Format("{0:#,##0}", row["数量"]) + row["単位"];
                        nr["単価"] = string.Format("{0:#,##0}", row["単価"]);
                        nr["税区分"] = row["消費税区分名称"] + " (" + row["課税区分名称"] + ") ";

                        nr["支払業者CD"] = row["支払業者"];
                        nr["支払業者名"] = row["支払業者名称"];
                        nr["支払額"] = string.Format("{0:#,##0}", row["支払額"]);
                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
                else
                {
                    //一列表示
                    //コード非表示
                    Disp.Columns.Add("売上年月", typeof(string));
                    Disp.Columns.Add("ﾚｺｰﾄﾞ番号", typeof(string));
                    Disp.Columns.Add("契約区分名", typeof(string));
                    Disp.Columns.Add("作業区分名", typeof(string));
                    Disp.Columns.Add("部門名", typeof(string));

                    Disp.Columns.Add("取引先名", typeof(string));
                    Disp.Columns.Add("契約項目", typeof(string));
                    Disp.Columns.Add("工事名", typeof(string));
                    Disp.Columns.Add("担当名", typeof(string));

                    //Disp.Columns.Add("締め日", typeof(string));
                    //Disp.Columns.Add("請求日", typeof(string));
                    Disp.Columns.Add("入力日", typeof(string));

                    Disp.Columns.Add("売上額", typeof(string));
                    Disp.Columns.Add("実施額", typeof(string));
                    Disp.Columns.Add("消費税額", typeof(string));

                    Disp.Columns.Add("数量", typeof(string));
                    Disp.Columns.Add("単価", typeof(string));
                    Disp.Columns.Add("税区分", typeof(string));

                    Disp.Columns.Add("支払業者名", typeof(string));
                    Disp.Columns.Add("支払額", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月"] = row["売上年月"];
                        nr["ﾚｺｰﾄﾞ番号"] = row["レコード番号"];
                        nr["契約区分名"] = row["契約区分名称"];
                        nr["作業区分名"] = row["作業区分名称"];
                        nr["部門名"] = row["部門名称"];

                        nr["取引先名"] = row["取引先名称（正式名称）"];
                        nr["契約項目"] = row["契約項目"];
                        nr["工事名"] = row["工事名称（正式名称）"];
                        nr["担当名"] = row["担当者氏名"];

                        //nr["締め日"] = row["締め日"];
                        //nr["請求日"] = row["請求日"];
                        nr["入力日"] = row["売上入力日"];

                        nr["売上額"] = string.Format("{0:#,##0}", row["売上額"]);
                        nr["実施額"] = string.Format("{0:#,##0}", row["実施額"]);
                        nr["消費税額"] = string.Format("{0:#,##0}", row["消費税額"]);

                        //nr["数量"] = string.Format("{0:#,##0}", row["数量"]) + " " + row["単位"];
                        nr["数量"] = string.Format("{0:#,##0}", row["数量"]) + row["単位"];
                        nr["単価"] = string.Format("{0:#,##0}", row["単価"]);
                        nr["税区分"] = row["消費税区分名称"] + " (" + row["課税区分名称"] + ") ";

                        nr["支払業者名"] = row["支払業者名称"];
                        nr["支払額"] = string.Format("{0:#,##0}", row["支払額"]);
                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
            }
            else
            {
                Disp.Columns.Add("売上年月　(ﾚｺｰﾄﾞ番号)\n契約(作業)区分名\n部門名", typeof(string));
                Disp.Columns.Add("取引先名\n契約項目\n工事名\n担当名", typeof(string));
                //Disp.Columns.Add("締め日\n請求日\n入力日", typeof(string));
                Disp.Columns.Add("入力日", typeof(string));
                Disp.Columns.Add("売上額\n実施額\n消費税額", typeof(string));
                Disp.Columns.Add("数量\n単価\n税区分", typeof(string));
                Disp.Columns.Add("支払業者\n支払業者名称\n支払額", typeof(string));

                if (checkBox2.Checked)
                {
                    //複数列表示
                    //コード表示
                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月　(ﾚｺｰﾄﾞ番号)\n契約(作業)区分名\n部門名"] = row["売上年月"] + " (" + row["レコード番号"] + ")\n" + row["契約区分名称"] + "(" + row["作業区分名称"] + ")\n" + row["部門コード"] + " " + row["部門名称"];
                        nr["取引先名\n契約項目\n工事名\n担当名"] = row["取引先コード"] + " " + row["取引先名称（正式名称）"] + "\n" + row["契約項目"] + "\n" + row["工事コード"] + "-" + row["工事枝コード"] + " " + row["工事名称（正式名称）"] + "\n" + row["担当者氏名"];
                        //nr["締め日\n請求日\n入力日"] = row["締め日"] + "\n" + row["請求日"] + "\n" + row["売上入力日"];
                        nr["入力日"] = row["売上入力日"];
                        nr["売上額\n実施額\n消費税額"] = string.Format("{0:#,##0}", row["売上額"]) + "\n" + string.Format("{0:#,##0}", row["実施額"]) + "\n" + string.Format("{0:#,##0}", row["消費税額"]);
                        nr["数量\n単価\n税区分"] = string.Format("{0:#,##0}", row["数量"]) + " " + row["単位"] + "\n" + string.Format("{0:#,##0}", row["単価"]) + "\n" + row["消費税区分名称"] + " (" + row["課税区分名称"] + ") ";
                        nr["支払業者\n支払業者名称\n支払額"] = row["支払業者"] + "\n" + row["支払業者名称"] + "\n" + string.Format("{0:#,##0}", row["支払額"]);
                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
                else
                {
                    //複数列表示
                    //コード表示

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月　(ﾚｺｰﾄﾞ番号)\n契約(作業)区分名\n部門名"] = row["売上年月"] + " (" + row["レコード番号"] + ")\n" + row["契約区分名称"] + "(" + row["作業区分名称"] + ")\n" + row["部門名称"];
                        nr["取引先名\n契約項目\n工事名\n担当名"] = row["取引先名称（正式名称）"] + "\n" + row["契約項目"] + "\n" + row["工事名称（正式名称）"] + "\n" + row["担当者氏名"];
                        //nr["締め日\n請求日\n入力日"] = row["締め日"] + "\n" + row["請求日"] + "\n" + row["売上入力日"];
                        nr["入力日"] = row["売上入力日"];
                        nr["売上額\n実施額\n消費税額"] = string.Format("{0:#,##0}", row["売上額"]) + "\n" + string.Format("{0:#,##0}", row["実施額"]) + "\n" + string.Format("{0:#,##0}", row["消費税額"]);
                        nr["数量\n単価\n税区分"] = string.Format("{0:#,##0}", row["数量"]) + " " + row["単位"] + "\n" + string.Format("{0:#,##0}", row["単価"]) + "\n" + row["消費税区分名称"] + " (" + row["課税区分名称"] + ") ";
                        nr["支払業者\n支払業者名称\n支払額"] = "" + "\n" + row["支払業者名称"] + "\n" + string.Format("{0:#,##0}", row["支払額"]);
                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
            }

            string ct = dt.Rows.Count.ToString();

            //TODO 共有クラスへ移動
            Com.InHistory("現売上", sDate + "～" + eDate + "【" + result + "】", ct);

            //データグリッドビューの高さ指定　※セット前にすること！
            if (checkBox1.Checked)
            {
                dataGridView1.RowTemplate.Height = 20;
            }
            else
            {
                dataGridView1.RowTemplate.Height = 75;
            }


            label1.Text = ct + " 件";
            label7.Text = "\\" + uriageAll.ToString("#,0") + "円";
            dataGridView1.DataSource = Disp;

            if (checkBox1.Checked)
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

                if (checkBox2.Checked)
                {
                    //金額右寄せ
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    //金額右寄せ
                    dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                //金額右寄せ
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //列幅変更
                dataGridView1.Columns[0].Width = 190;
                dataGridView1.Columns[1].Width = 530;
                dataGridView1.Columns[2].Width = 90;
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].Width = 160;
                dataGridView1.Columns[5].Width = 220;
            }

            //ストップ
            sw.Stop();

            //処理速度表示
            label2.Text = sw.Elapsed.TotalSeconds.ToString("F") + " 秒";

            System.GC.Collect();
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //ボタン無効化・カーソル変更
                button1.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                //データ処理
                GetUriageData();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ処理
            GetUriageData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
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
    }
}
