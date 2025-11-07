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

namespace ODIS.ODIS
{
    public partial class Kaikei : Form
    {
        
        private DataTable dt = new DataTable();


        public Kaikei()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //データ取得
            //GetFurikaeData();

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

            //DateTime datet = DateTime.Now.AddDays(-15);

            //DateTime firstDayOfMonth1 = new DateTime(datet.Year, datet.Month, 1);

            //dateTimePicker1.Value = firstDayOfMonth1;
            dateTimePicker1.Value = new DateTime(2023, 03, 01);

            //int days = DateTime.DaysInMonth(datet.Year, datet.Month); // その月の日数
            //var lastDayOfMonth2 = new DateTime(datet.Year, datet.Month, days);

            //dateTimePicker2.Value = lastDayOfMonth2;
            dateTimePicker2.Value = new DateTime(2023, 03, 31);

        }



        private void GetData()
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

            string sDate = dateTimePicker1.Value.ToString("yyyyMMdd");
            string eDate = dateTimePicker2.Value.ToString("yyyyMMdd");

            //期間
            //result += " and suitouymd >= " + sDate + " and suitouymd <= " + eDate;

            //先頭が「and」の場合、削除する
            //if (result.StartsWith(" and"))
            //{
            //    result = result.Remove(0, 4);
            //}

            //DataRow[] dtrow;
            //dtrow = dt.Select(result, "");

            int nRet;

            NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr);
            conn.Open();
            DataTable dt = new DataTable();

            string sql = "select * from kpcp01.\"CostomGetFurikaeDataNew\" where suitouymd between '" + sDate + "' and '" + eDate + "' " + result;

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
            nRet = adapter.Fill(dt);

            conn.Close();

            //解放
            adapter.Dispose();
            conn.Dispose();


            DataTable Disp = new DataTable();

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

                    Disp.Columns.Add("日付", typeof(string));
                    Disp.Columns.Add("伝票番号-行番号", typeof(string));

                    Disp.Columns.Add("【借方】\n科目CD", typeof(string));
                    Disp.Columns.Add("【借方】\n科目名", typeof(string));
                    Disp.Columns.Add("【借方】\n部門CD", typeof(string));
                    Disp.Columns.Add("【借方】\n部門名", typeof(string));
                    Disp.Columns.Add("【借方】\n取引先CD", typeof(string));
                    Disp.Columns.Add("【借方】\n取引先名", typeof(string));
                    Disp.Columns.Add("【借方】\n工事CD", typeof(string));
                    Disp.Columns.Add("【借方】\n工事名", typeof(string));
                    Disp.Columns.Add("【借方】\n費目CD", typeof(string));
                    Disp.Columns.Add("【借方】\n費目名", typeof(string));
                    Disp.Columns.Add("【借方】\n細目CD", typeof(string));
                    Disp.Columns.Add("【借方】\n細目名", typeof(string));
                    Disp.Columns.Add("【借方】\n金額", typeof(string));
                    Disp.Columns.Add("【借方】\n消費税", typeof(string));
                    Disp.Columns.Add("【借方】\n課税区分", typeof(string));

                    Disp.Columns.Add("【貸方】\n科目CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n科目名", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門名", typeof(string));
                    Disp.Columns.Add("【貸方】\n取引先CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n取引先名", typeof(string));
                    Disp.Columns.Add("【貸方】\n工事CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n工事名", typeof(string));

                    Disp.Columns.Add("【貸方】\n金額", typeof(string));
                    Disp.Columns.Add("【貸方】\n消費税", typeof(string));
                    Disp.Columns.Add("【貸方】\n課税区分", typeof(string));

                    Disp.Columns.Add("摘要", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付"] = row["suitouymd"];
                        nr["伝票番号-行番号"] = row["denpyounumber"] + "-" + row["gyounumber"];

                        nr["【借方】\n科目CD"] = row["l_kamokucode"] + "-" + row["l_uchiwakecode"];
                        nr["【借方】\n科目名"] = row["l_kamokuname"];
                        nr["【借方】\n部門CD"] = row["l_bumoncode"];
                        nr["【借方】\n部門名"] = row["l_bumonname"];
                        nr["【借方】\n取引先CD"] = row["l_torihikisakicode"];
                        nr["【借方】\n取引先名"] = row["l_torihikisakiname"];
                        nr["【借方】\n工事CD"] = row["l_koujicode"] + "" + row["l_koujiedacode"];
                        nr["【借方】\n工事名"] = row["l_koujiname"];

                        nr["【借方】\n費目CD"] = row["l_himokucode"];
                        nr["【借方】\n費目名"] = row["l_himokuname"];
                        nr["【借方】\n細目CD"] = row["l_saimokucode"] + "-" + row["l_saimokuedacode"];
                        nr["【借方】\n細目名"] = row["l_saimokuname"];

                        nr["【借方】\n金額"] = row["l_denpyoukingaku"];
                        nr["【借方】\n消費税"] = row["l_syouhizeikingaku"];
                        nr["【借方】\n課税区分"] = row["l_syouhizeikubun"];

                        nr["【貸方】\n科目CD"] = row["r_kamokucode"] + "-" + row["l_uchiwakecode"];
                        nr["【貸方】\n科目名"] = row["r_kamokuname"];
                        nr["【貸方】\n部門CD"] = row["r_bumoncode"];
                        nr["【貸方】\n部門名"] = row["r_bumonname"];
                        nr["【貸方】\n取引先CD"] = row["r_torihikisakicode"];
                        nr["【貸方】\n取引先名"] = row["r_torihikisakiname"];
                        nr["【貸方】\n工事CD"] = row["r_koujicode"] + "" + row["l_koujiedacode"];
                        nr["【貸方】\n工事名"] = row["r_koujiname"];

                        nr["【貸方】\n金額"] = row["r_denpyoukingaku"];
                        nr["【貸方】\n消費税"] = row["r_syouhizeikingaku"];
                        nr["【貸方】\n課税区分"] = row["r_syouhizeikubun"];

                        nr["摘要"] = row["tekiyou"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
                else
                {
                    //一列表示
                    //コード非表示
                    Disp.Columns.Add("日付", typeof(string));
                    Disp.Columns.Add("伝票番号-行番号", typeof(string));

                    Disp.Columns.Add("【借方】\n科目", typeof(string));
                    Disp.Columns.Add("【借方】\n部門", typeof(string));
                    Disp.Columns.Add("【借方】\n取引先", typeof(string));
                    Disp.Columns.Add("【借方】\n工事", typeof(string));
                    Disp.Columns.Add("【借方】\n費目", typeof(string));
                    Disp.Columns.Add("【借方】\n細目", typeof(string));
                    Disp.Columns.Add("【借方】\n金額", typeof(string));
                    Disp.Columns.Add("【借方】\n消費税", typeof(string));
                    Disp.Columns.Add("【借方】\n課税区分", typeof(string));

                    Disp.Columns.Add("【貸方】\n科目", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門", typeof(string));
                    Disp.Columns.Add("【貸方】\n取引先", typeof(string));
                    Disp.Columns.Add("【貸方】\n工事", typeof(string));

                    Disp.Columns.Add("【貸方】\n金額", typeof(string));
                    Disp.Columns.Add("【貸方】\n消費税", typeof(string));
                    Disp.Columns.Add("【貸方】\n課税区分", typeof(string));

                    Disp.Columns.Add("摘要", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付"] = row["suitouymd"];
                        nr["伝票番号-行番号"] = row["denpyounumber"] + "-" + row["gyounumber"];

                        nr["【借方】\n科目"] = row["l_kamokuname"];
                        nr["【借方】\n部門"] = row["l_bumonname"];
                        nr["【借方】\n取引先"] = row["l_torihikisakiname"];
                        nr["【借方】\n工事"] = row["l_koujiname"];

                        nr["【借方】\n費目"] = row["l_himokuname"];
                        nr["【借方】\n細目"] = row["l_saimokuname"];

                        nr["【借方】\n金額"] = row["l_denpyoukingaku"];
                        nr["【借方】\n消費税"] = row["l_syouhizeikingaku"];
                        nr["【借方】\n課税区分"] = row["l_syouhizeikubun"];

                        nr["【貸方】\n科目"] = row["r_kamokuname"];
                        nr["【貸方】\n部門"] = row["r_bumonname"];
                        nr["【貸方】\n取引先"] = row["r_torihikisakiname"];
                        nr["【貸方】\n工事"] = row["r_koujiname"];

                        nr["【貸方】\n金額"] = row["r_denpyoukingaku"];
                        nr["【貸方】\n消費税"] = row["r_syouhizeikingaku"];
                        nr["【貸方】\n課税区分"] = row["r_syouhizeikubun"];

                        nr["摘要"] = row["tekiyou"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }
            else
            {
                Disp.Columns.Add("日付\n伝票番号\n行番号", typeof(string));
                Disp.Columns.Add("【借方】\n科目\n部門\n取引先\n工事", typeof(string));
                Disp.Columns.Add("【借方】\n費目\n細目", typeof(string));
                Disp.Columns.Add("【借方】\n金額\n消費税\n課税区分", typeof(string));
                Disp.Columns.Add("【貸方】\n科目\n部門\n取引先\n工事", typeof(string));
                Disp.Columns.Add("【貸方】\n金額\n消費税\n課税区分", typeof(string));
                Disp.Columns.Add("摘要", typeof(string));

                if (checkBox2.Checked)
                {
                    //複数列表示
                    //コード表示
                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付\n伝票番号\n行番号"] = row["suitouymd"] + "\n" + row["denpyounumber"] + "\n" + row["gyounumber"];

                        nr["【借方】\n科目\n部門\n取引先\n工事"] = row["l_kamokucode"] + "-" + row["l_uchiwakecode"] + "　" + row["l_kamokuname"] + "\n" + row["l_bumoncode"] + "　　　　" + row["l_bumonname"] + "\n" + row["l_torihikisakicode"] + "　" + row["l_torihikisakiname"] + (row["l_koujicode"].ToString() == "" ? "" : "\n") + row["l_koujicode"] + (row["l_koujicode"].ToString() == "" ? "" : "-") + row["l_koujiedacode"] + "　" + row["l_koujiname"];
                        nr["【借方】\n費目\n細目"] = row["l_himokucode"] + "　" + row["l_himokuname"] + "\n" + row["l_saimokucode"] + "-" + row["l_saimokuedacode"] + "\n" + row["l_saimokuname"];
                        nr["【借方】\n金額\n消費税\n課税区分"] = row["l_denpyoukingaku"] + "\n" + row["l_syouhizeikingaku"] + "\n" + row["l_syouhizeikubun"];

                        nr["【貸方】\n科目\n部門\n取引先\n工事"] = row["r_kamokucode"] + "-" + row["r_uchiwakecode"] + "　" + row["r_kamokuname"] + "\n" + row["r_bumoncode"] + "　　　　" + row["r_bumonname"] + "\n" + row["r_torihikisakicode"] + "　" + row["r_torihikisakiname"] + (row["r_koujicode"].ToString() == "" ? "" : "\n") + row["r_koujicode"] + (row["r_koujicode"].ToString() == "" ? "" : "-") + row["r_koujiedacode"] + "　" + row["r_koujiname"];
                        nr["【貸方】\n金額\n消費税\n課税区分"] = row["r_denpyoukingaku"] + "\n" + row["r_syouhizeikingaku"] + "\n" + row["r_syouhizeikubun"];

                        nr["摘要"] = row["tekiyou"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
                else
                {
                    //複数列表示
                    //コード表示

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付\n伝票番号\n行番号"] = row["suitouymd"] + "\n" + row["denpyounumber"] + "\n" + row["gyounumber"];

                        nr["【借方】\n科目\n部門\n取引先\n工事"] = row["l_kamokuname"] + "\n" + row["l_bumonname"] + "\n" + row["l_torihikisakiname"] + "\n" + row["l_koujiname"];
                        nr["【借方】\n費目\n細目"] = row["l_himokuname"] + "\n" + row["l_saimokuname"];
                        nr["【借方】\n金額\n消費税\n課税区分"] = row["l_denpyoukingaku"] + "\n" + row["l_syouhizeikingaku"] + "\n" + row["l_syouhizeikubun"];

                        nr["【貸方】\n科目\n部門\n取引先\n工事"] = row["r_kamokuname"] + "\n" + row["r_bumonname"] + "\n" + row["r_torihikisakiname"] + row["r_koujiname"];
                        nr["【貸方】\n金額\n消費税\n課税区分"] = row["r_denpyoukingaku"] + "\n" + row["r_syouhizeikingaku"] + "\n" + row["r_syouhizeikubun"];

                        nr["摘要"] = row["tekiyou"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }

            string ct = dt.Rows.Count.ToString();

            //TODO 共有クラスへ移動
            Com.InHistory("01_会計検索(13/04～23/03)", sDate + "～" + eDate + "【" + res + "】", ct);

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
            dataGridView1.DataSource = Disp;

            if (checkBox1.Checked)
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

                if (checkBox2.Checked)
                {
                    //金額右寄せ
                    dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    //背景色変更
                    dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[12].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[13].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[15].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[16].DefaultCellStyle.BackColor = Color.AliceBlue;

                    dataGridView1.Columns[17].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[18].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[19].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[20].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[21].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[22].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[23].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[24].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[25].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[26].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[27].DefaultCellStyle.BackColor = Color.MistyRose;
                }
                else
                {
                    //金額右寄せ
                    dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    //背景色変更
                    dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.AliceBlue;

                    dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[12].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[13].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[15].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[16].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[17].DefaultCellStyle.BackColor = Color.MistyRose;
                }
            }
            else
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                //金額右寄せ
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //背景色変更
                dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.AliceBlue;
                dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.AliceBlue;
                dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.AliceBlue;

                dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.MistyRose;
                dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.MistyRose;

                if (checkBox2.Checked)
                {
                    //列幅変更
                    dataGridView1.Columns[0].Width = 100;
                    dataGridView1.Columns[1].Width = 360;
                    dataGridView1.Columns[2].Width = 130;
                    dataGridView1.Columns[3].Width = 140;
                    dataGridView1.Columns[4].Width = 360;
                    dataGridView1.Columns[5].Width = 140;
                    dataGridView1.Columns[6].Width = 310;
                }
                else
                {
                    //列幅変更
                    dataGridView1.Columns[0].Width = 100;
                    dataGridView1.Columns[1].Width = 360;
                    dataGridView1.Columns[2].Width = 130;
                    dataGridView1.Columns[3].Width = 140;
                    dataGridView1.Columns[4].Width = 360;
                    dataGridView1.Columns[5].Width = 140;
                    dataGridView1.Columns[6].Width = 310;
                }
            }

            //ストップ
            sw.Stop();

            //処理速度表示
            label2.Text = sw.Elapsed.TotalSeconds.ToString("F") + " 秒";

            System.GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ処理
            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
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
                GetData();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
        }
    }
}
