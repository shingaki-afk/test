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
    public partial class Kaikei_PCA : Form
    {

        private DataTable dt = new DataTable();


        public Kaikei_PCA()
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

            DateTime datet = DateTime.Now.AddDays(-15);

            DateTime firstDayOfMonth1 = new DateTime(datet.Year, datet.Month, 1);

            dateTimePicker1.Value = firstDayOfMonth1;
            //dateTimePicker1.Value = new DateTime(2023, 03, 01);

            int days = DateTime.DaysInMonth(datet.Year, datet.Month); // その月の日数
            var lastDayOfMonth2 = new DateTime(datet.Year, datet.Month, days);

            dateTimePicker2.Value = lastDayOfMonth2;
            //dateTimePicker2.Value = new DateTime(2023, 03, 31);

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

          

            string sql = "select * from dbo.PCA会計仕訳チェックデータ検索 where 伝票日付 between '" + sDate + "' and '" + eDate + "' " + result;
            sql += " order by 伝票日付, 伝票番号 ";
            
            dt = Com.GetDB(sql);


            DataTable Disp = new DataTable();

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            Disp.Columns.Add("日付\n伝票番号", typeof(string));
            Disp.Columns.Add("【借方】\n科目\n部門\n取引先\n工事", typeof(string));
            Disp.Columns.Add("【借方】\n工種", typeof(string));
            Disp.Columns.Add("【借方】\n金額\n消費税\n課税区分", typeof(string));
            Disp.Columns.Add("【貸方】\n科目\n部門\n取引先\n工事", typeof(string));
            Disp.Columns.Add("【貸方】\n金額\n消費税\n課税区分", typeof(string));
            Disp.Columns.Add("摘要", typeof(string));

            Int64 kingaku_kari = 0;
            Int64 kingaku_kashi = 0;

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = Disp.NewRow();
                nr["日付\n伝票番号"] = row["伝票日付"] + "\n" + row["伝票番号"];

                //nr["【借方】\n科目\n部門\n取引先\n工事"] = row["借方科目コード"] + "　" + row["借方科目名"] + "\n" + row["借方部門コード"] + "　" + row["借方部門名"] + "\n" + row["借方取引先コード"] + "　" + row["借方取引先名"] + (row["借方工事コード"].ToString() == "" ? "" : "\n") + row["借方工事コード"] + (row["借方工事コード"].ToString() == "" ? "" : "-") + row["借方工事コード"] + "　" + row["借方工事名"];
                nr["【借方】\n科目\n部門\n取引先\n工事"] = row["借方科目コード"] + "　" + row["借方科目名"] + "　　" + row["借方補助コード"] + "　" + row["借方補助名"] + "\n" + row["借方部門コード"] + "　" + row["借方部門名"] + "\n" + row["借方取引先コード"] + "　" + row["借方取引先名"] + "\n" + row["借方工事コード"] + "　" + row["借方工事名"];
                nr["【借方】\n工種"] = row["借方工種コード"] + "　" + row["借方工種名"];
                nr["【借方】\n金額\n消費税\n課税区分"] = row["借方金額"] + "\n" + row["借方消費税額"] + "\n" + row["借方税区分名"];

                //nr["【貸方】\n科目\n部門\n取引先\n工事"] = row["貸方科目コード"] + "　" + row["貸方科目名"] + "\n" + row["貸方部門コード"] + "　" + row["貸方部門名"] + "\n" + row["貸方取引先コード"] + "　" + row["貸方取引先名"] + (row["貸方工事コード"].ToString() == "" ? "" : "\n") + row["貸方工事コード"] + (row["貸方工事コード"].ToString() == "" ? "" : "-") + row["貸方工事コード"] + "　" + row["貸方工事名"];
                nr["【貸方】\n科目\n部門\n取引先\n工事"] = row["貸方科目コード"] + "　" + row["貸方科目名"] + "　　" + row["貸方補助コード"] + "　" + row["貸方補助名"] + "\n" + row["貸方部門コード"] + "　" + row["貸方部門名"] + "\n" + row["貸方取引先コード"] + "　" + row["貸方取引先名"] + "\n" + row["貸方工事コード"] + "　" + row["貸方工事名"];

                nr["【貸方】\n金額\n消費税\n課税区分"] = row["貸方金額"] + "\n" + row["貸方消費税額"] + "\n" + row["貸方税区分名"];

                nr["摘要"] = row["摘要文"].ToString();

                kingaku_kari += row["借方金額"] == DBNull.Value ? 0 :  Convert.ToInt64(row["借方金額"]);
                kingaku_kashi += row["貸方金額"] == DBNull.Value ? 0 : Convert.ToInt64(row["貸方金額"]); 

                Disp.Rows.Add(nr);
            }


            string ct = dt.Rows.Count.ToString();

            lbl_kari.Text = kingaku_kari.ToString("#,##0.##;-#,##0.##;#");
            lbl_kashi.Text = kingaku_kashi.ToString("#,##0.##;-#,##0.##;#");

            //TODO 共有クラスへ移動
            Com.InHistory("01_会計検索(24/04～)", sDate + "～" + eDate + "【" + res + "】", ct);


            dataGridView1.RowTemplate.Height = 75;


            label1.Text = ct + " 件";
            dataGridView1.DataSource = Disp;

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


            //列幅変更
            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].Width = 360;
            dataGridView1.Columns[2].Width = 130;
            dataGridView1.Columns[3].Width = 140;
            dataGridView1.Columns[4].Width = 360;
            dataGridView1.Columns[5].Width = 140;
            dataGridView1.Columns[6].Width = 310;



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
