using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;

namespace ODIS.ODIS
{
    public partial class KakoKaikei : Form
    {
        public KakoKaikei()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //都度取得に変更！

            //データ取得
            //GetKakoKaikeiData();

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
            dateTimePicker1.Value = new DateTime(2005, 04, 01);
            dateTimePicker2.Value = new DateTime(2013, 03, 31);
        }

        private void DataView()
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
                    //result += " and (reskey not like '%" + s + "%' or reskey not like '%" + Com.isOneByteChar(s) + "%' or reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    result += " and (reskey not like '%" + s + "%' and reskey not like '%" + Com.isOneByteChar(s) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            string sDate = dateTimePicker1.Value.ToString("yyyyMMdd");
            string eDate = dateTimePicker2.Value.ToString("yyyyMMdd");

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
                        Cmd.CommandText = "select * from dbo.旧伝票検索 where 伝票日付 between " + sDate + " and " + eDate + result;
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

                    Disp.Columns.Add("伝票日付", typeof(string));
                    Disp.Columns.Add("伝票番号-行番号", typeof(string));
                    Disp.Columns.Add("地区", typeof(string));

                    Disp.Columns.Add("【借方】\n科目CD", typeof(string));
                    Disp.Columns.Add("【借方】\n科目名", typeof(string));
                    Disp.Columns.Add("【借方】\n内訳CD", typeof(string));
                    Disp.Columns.Add("【借方】\n内訳名", typeof(string));
                    Disp.Columns.Add("【借方】\n部門CD", typeof(string));
                    Disp.Columns.Add("【借方】\n部門名", typeof(string));
                    Disp.Columns.Add("【借方】\n金額", typeof(int));

                    Disp.Columns.Add("【貸方】\n科目CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n科目名", typeof(string));
                    Disp.Columns.Add("【貸方】\n内訳CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n内訳名", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門CD", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門名", typeof(string));
                    Disp.Columns.Add("【貸方】\n金額", typeof(int));

                    Disp.Columns.Add("摘要", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["伝票日付"] = row["伝票日付"];
                        nr["伝票番号-行番号"] = row["伝票番号"] + "-" + row["伝票行"];
                        nr["地区"] = row["地区"];

                        nr["【借方】\n科目CD"] = row["借方科目コード"];
                        nr["【借方】\n科目名"] = row["借方科目名"];
                        nr["【借方】\n内訳CD"] = row["借方内訳コード"];
                        nr["【借方】\n内訳名"] = row["借方内訳名"];
                        nr["【借方】\n部門CD"] = row["借方部門コード"];
                        nr["【借方】\n部門名"] = row["借方部門名"];
                        nr["【借方】\n金額"] = row["借方金額"];

                        nr["【貸方】\n科目CD"] = row["貸方科目コード"];
                        nr["【貸方】\n科目名"] = row["貸方科目名"];
                        nr["【貸方】\n内訳CD"] = row["貸方内訳コード"];
                        nr["【貸方】\n内訳名"] = row["貸方内訳名"];
                        nr["【貸方】\n部門CD"] = row["貸方部門コード"];
                        nr["【貸方】\n部門名"] = row["貸方部門名"];
                        nr["【貸方】\n金額"] = row["貸方金額"];

                        nr["摘要"] = row["摘要"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
                else
                {
                    //一列表示
                    //コード非表示
                    Disp.Columns.Add("伝票日付", typeof(string));
                    Disp.Columns.Add("伝票番号-行番号", typeof(string));
                    Disp.Columns.Add("地区", typeof(string));

                    Disp.Columns.Add("【借方】\n科目名", typeof(string));
                    Disp.Columns.Add("【借方】\n内訳名", typeof(string));
                    Disp.Columns.Add("【借方】\n部門名", typeof(string));
                    Disp.Columns.Add("【借方】\n金額", typeof(int));

                    Disp.Columns.Add("【貸方】\n科目名", typeof(string));
                    Disp.Columns.Add("【貸方】\n内訳名", typeof(string));
                    Disp.Columns.Add("【貸方】\n部門名", typeof(string));
                    Disp.Columns.Add("【貸方】\n金額", typeof(int));

                    Disp.Columns.Add("摘要", typeof(string));

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["伝票日付"] = row["伝票日付"];
                        nr["伝票番号-行番号"] = row["伝票番号"] + "-" + row["伝票行"];
                        nr["地区"] = row["地区"];

                        nr["【借方】\n科目名"] = row["借方科目名"];
                        nr["【借方】\n内訳名"] = row["借方内訳名"];
                        nr["【借方】\n部門名"] = row["借方部門名"];
                        nr["【借方】\n金額"] = row["借方金額"];

                        nr["【貸方】\n科目名"] = row["貸方科目名"];
                        nr["【貸方】\n内訳名"] = row["貸方内訳名"];
                        nr["【貸方】\n部門名"] = row["貸方部門名"];
                        nr["【貸方】\n金額"] = row["貸方金額"];

                        nr["摘要"] = row["摘要"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }
            else
            {
                Disp.Columns.Add("日付\n伝票番号\n地区", typeof(string));
                Disp.Columns.Add("【借方】\n科目\n内訳\n部門", typeof(string));
                Disp.Columns.Add("【借方】\n金額", typeof(int));
                Disp.Columns.Add("【貸方】\n科目\n内訳\n部門", typeof(string));
                Disp.Columns.Add("【貸方】\n金額", typeof(int));
                Disp.Columns.Add("摘要", typeof(string));

                if (checkBox2.Checked)
                {
                    //複数列表示
                    //コード表示
                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付\n伝票番号\n地区"] = row["伝票日付"] + "\n" + row["伝票番号"] + "-" + row["伝票行"] + "\n" + row["地区"];

                        nr["【借方】\n科目\n内訳\n部門"] = row["借方科目コード"] + "　" + row["借方科目名"] + "\n" + row["借方内訳コード"] + "　" + row["借方内訳名"] + "\n" + row["借方部門コード"] + "　" + row["借方部門名"];
                        nr["【借方】\n金額"] = row["借方金額"];

                        nr["【貸方】\n科目\n内訳\n部門"] = row["貸方科目コード"] + "　" + row["貸方科目名"] + "\n" + row["貸方内訳コード"] + "　" + row["貸方内訳名"] + "\n" + row["貸方部門コード"] + "　" + row["貸方部門名"];
                        nr["【貸方】\n金額"] = row["貸方金額"];

                        nr["摘要"] = row["摘要"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
                else
                {
                    //複数列表示
                    //コード非表示

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["日付\n伝票番号\n地区"] = row["伝票日付"] + "\n" + row["伝票番号"] + "-" + row["伝票行"] + "\n" + row["地区"];

                        nr["【借方】\n科目\n内訳\n部門"] = row["借方科目名"] + "\n" + row["借方内訳名"] + "\n" + row["借方部門名"];
                        nr["【借方】\n金額"] = row["借方金額"];

                        nr["【貸方】\n科目\n内訳\n部門"] = row["貸方科目名"] + "\n" + row["貸方内訳名"] + "\n" + row["貸方部門名"];
                        nr["【貸方】\n金額"] = row["貸方金額"];

                        nr["摘要"] = row["摘要"].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }

            string ct = dt.Rows.Count.ToString("#,##0.##;-#,##0.##;#");

            //TODO 共有クラスへ移動
            Com.InHistory("旧会計", sDate + "～" + eDate + "【" + res + "】", ct);

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

            #region 表示形式
            if (checkBox1.Checked)
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

                if (checkBox2.Checked)
                {
                    //金額右寄せ
                    dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[16].DefaultCellStyle.Format = "#,0";

                    //背景色変更
                    dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.AliceBlue;

                    dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[12].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[13].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[15].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[16].DefaultCellStyle.BackColor = Color.MistyRose;
                }
                else
                {
                    //金額右寄せ
                    dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";
                    //背景色変更
                    dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.AliceBlue;

                    dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.MistyRose;
                }
            }
            else
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                //金額右寄せ
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";

                //背景色変更
                dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.AliceBlue;
                dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.AliceBlue;

                dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.MistyRose;
                dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.MistyRose;

                //列幅変更
                dataGridView1.Columns[0].Width = 120;
                dataGridView1.Columns[1].Width = 230;
                dataGridView1.Columns[2].Width = 100;
                dataGridView1.Columns[3].Width = 230;
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[5].Width = 440;
            }
            #endregion

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
            DataView();

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
                DataView();

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
