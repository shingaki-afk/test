using Microsoft.VisualBasic;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KeisuuKanri : Form
    {
        public KeisuuKanri()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //日付設定
            DateTime today = DateTime.Today;
            if (today.Day >= 15)
            {
                dateTimePicker1.Value = today;
                dateTimePicker2.Value = today;
            }
            else
            {
                dateTimePicker1.Value = today.AddMonths(-1);
                dateTimePicker2.Value = today.AddMonths(-1);
            }

            //checkedListBox1.Items.Add("契約固定");
            checkedListBox1.Items.Add("契約臨時");
            checkedListBox1.Items.Add("臨時");
            checkedListBox1.Items.Add("物品");
            checkedListBox1.SetItemChecked(0, true);
            checkedListBox1.SetItemChecked(1, true);
            checkedListBox1.SetItemChecked(2, true);
            //checkedListBox1.SetItemChecked(3, true);

            checkedListBox2.Items.Add("自社");
            checkedListBox2.Items.Add("外注");
            checkedListBox2.SetItemChecked(0, true);
            checkedListBox2.SetItemChecked(1, true);

            checkedListBox3.Items.Add("10_現業");
            checkedListBox3.Items.Add("11_技術企画");
            checkedListBox3.Items.Add("20_客室");
            checkedListBox3.Items.Add("30_施設");
            checkedListBox3.Items.Add("31_警備");
            checkedListBox3.Items.Add("32_遠方監視");
            checkedListBox3.Items.Add("40_エンジ");
            checkedListBox3.Items.Add("41_米軍");
            checkedListBox3.Items.Add("50_PPP/PFI");
            checkedListBox3.Items.Add("60_サービス");
            checkedListBox3.SetItemChecked(0, true);
            checkedListBox3.SetItemChecked(1, true);
            checkedListBox3.SetItemChecked(2, true);
            checkedListBox3.SetItemChecked(3, true);
            checkedListBox3.SetItemChecked(4, true);
            checkedListBox3.SetItemChecked(5, true);
            checkedListBox3.SetItemChecked(6, true);
            checkedListBox3.SetItemChecked(7, true);
            checkedListBox3.SetItemChecked(8, true);
            checkedListBox3.SetItemChecked(9, true);

            checkedListBox4.Items.Add("那覇");
            checkedListBox4.Items.Add("八重山");
            checkedListBox4.Items.Add("北部");
            checkedListBox4.SetItemChecked(0, true);
            checkedListBox4.SetItemChecked(1, true);
            checkedListBox4.SetItemChecked(2, true);

            //SetBumon();
        }

        //private void SetBumon()
        //{
        //    //リストボックスの項目(Item)を消去
        //    checkedListBox3.Items.Clear();

        //    DataTable dt = new DataTable();

        //    string sql = "select distinct 部門コード, 部門区分 from kpcp01.\"CostomKeisuukanridaityou_serch\" order by 部門コード";

        //    dt = Com.GetPosDB(sql);

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        //checkedListBox3.Items.Add(row["部門コード"].ToString() + ' ' + row["部門"].ToString());
        //        checkedListBox3.Items.Add(row["部門"].ToString());
        //    }

        //    for (int i = 0; i < checkedListBox3.Items.Count; i++)
        //    {
        //        checkedListBox3.SetItemChecked(i, true);
        //    }
        //}

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


            Com.InHistory("15_計数管理台帳", "", "");
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

            //契約区分
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        result += " and 契約区分 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            //臨時・外注
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 作業区分 <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            //部門
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    result += " and 部門区分 <> '" + checkedListBox3.Items[i].ToString() + "'";
                }
            }

            //地区
            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                if (!checkedListBox4.GetItemChecked(i))
                {
                    result += " and 地区区分 <> '" + checkedListBox4.Items[i].ToString() + "'";
                }
            }


            string sDate = dateTimePicker1.Value.ToString("yyyyMM");
            string eDate = dateTimePicker2.Value.ToString("yyyyMM");


            int nRet;

            NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr);
            conn.Open();
            DataTable dt = new DataTable();

            //string sql = "select * from kpcp01.\"GetUriageDataSearchAll\" where 売上年月 between '" + sDate + "' and '" + eDate + "'" + result;
            string sql = "select * from kpcp01.\"CostomKeisuukanridaityou_serch\" where 売上年月 between '" + sDate + "' and '" + eDate + "'" + result;

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
            nRet = adapter.Fill(dt);

            conn.Close();

            //解放
            adapter.Dispose();
            conn.Dispose();

            dataGridView1.DataSource = dt;

            decimal uriageAll = 0;
            decimal keihiAll = 0;

            foreach (DataRow row in dt.Rows)
            {
                uriageAll += Convert.ToDecimal(row["売上額"]);

                if (row["経費合計"].ToString() != "")
                { 
                    keihiAll += Convert.ToDecimal(row["経費合計"]);
                }
            }

            //金額右寄せ
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[7].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[11].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[12].DefaultCellStyle.Format = "#,##0.#";
            dataGridView1.Columns[13].DefaultCellStyle.Format = "#,##0.#";
            
            //列幅変更
            dataGridView1.Columns[0].Width = 70;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[2].Width = 70;
            dataGridView1.Columns[3].Width = 110;
            dataGridView1.Columns[4].Width = 260;
            dataGridView1.Columns[5].Width = 260;
            dataGridView1.Columns[6].Width = 120;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].Width = 80;
            dataGridView1.Columns[14].Width = 70;
            dataGridView1.Columns[15].Width = 150;

            dataGridView1.Columns[16].Visible = false;
            dataGridView1.Columns[17].Visible = false;
            dataGridView1.Columns[18].Visible = false;
            dataGridView1.Columns[19].Visible = false;
            dataGridView1.Columns[20].Visible = false;

            //ストップ
            sw.Stop();

            //処理速度表示
            label2.Text = sw.Elapsed.TotalSeconds.ToString("F") + " 秒";

            label1.Text = dt.Rows.Count + " 件";
            label7.Text = uriageAll.ToString("#,0") + "円";
            label11.Text = keihiAll.ToString("#,0") + "円";
            label17.Text = (uriageAll - keihiAll).ToString("#,0") + "円";

            if (uriageAll > 0)
            { 
                label10.Text = Math.Round(keihiAll / uriageAll * 100, 1).ToString() + "%";
            }
            else
            {
                label10.Text = "";
            }

            System.GC.Collect();
        }

        private void label13_Click(object sender, EventArgs e)
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

        private void label12_Click(object sender, EventArgs e)
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

        private void label14_Click(object sender, EventArgs e)
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

        private void label15_Click(object sender, EventArgs e)
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

        private void label9_Click(object sender, EventArgs e)
        {

        }
    }
}
