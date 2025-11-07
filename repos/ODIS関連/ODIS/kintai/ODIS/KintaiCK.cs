using Microsoft.VisualBasic;
using ODIS.ODIS;
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
    public partial class KintaiCK : Form
    {
        private DataTable dt = new DataTable();

        public KintaiCK()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            checkedListBox1.Items.Add("本社");
            checkedListBox1.Items.Add("那覇");
            checkedListBox1.Items.Add("八重山");
            checkedListBox1.Items.Add("北部");
            checkedListBox1.Items.Add("広域");
            checkedListBox1.Items.Add("宮古島");
            checkedListBox1.Items.Add("久米島");

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }

            DataTable ymdt = new DataTable();
            int ct = 0;

            ymdt = Com.GetDB("select distinct 処理年 + 処理月 as 処理年月 from dbo.KM_給与明細 order by 処理年 +処理月");
            
            foreach (DataRow row in ymdt.Rows)
            {
                comboBox1.Items.Add(row["処理年月"]);
                comboBox2.Items.Add(row["処理年月"]);
                ct++;
            }

            comboBox1.SelectedIndex = ct-1;
            comboBox2.SelectedIndex = ct-1;

            GetData();
            //Com.InHistory("41_勤怠検索", "", "");
        }

        private void GetData()
        {
            //ボタン無効化・カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            string sql = "select 処理年月, 社員番号, 氏名, 地区名,組織名,現場名,支給区分, 勤務時間 ";
            sql += " ,総労働時間, 延長時間,法休時間,所休時間,残業時間,[60超残Ｈ],深夜時間,遅刻回数,遅刻時間 ";
            sql += " ,所定,法休,所休,有給,特休,無特,振休,公休,調休,届欠,無届,回数１,回数２ from k勤怠検索 ";
            sql += " where 処理年月 between '" + comboBox1.SelectedItem + "' and '" + comboBox2.SelectedItem + "' ";
            //地区
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        sql += " and 地区名 <> '" + checkedListBox1.Items[i] + "'";
                    }
                }
            }

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    sql += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            sql += " order by 社員番号, 地区名, 組織名, 現場名, 処理年月";
            dt = Com.GetDB(sql);

            dataGridView1.DataSource = dt;

            if (dt.Rows.Count == 0) return;


            for (int i = 0; i < dt.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i < 7)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.Columns[i].Width = 80;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[i].Width = 40;
                }

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                //dataGridView1.Columns[i].DefaultCellStyle.Format = "N1";

            }

            Com.InHistory("41_勤怠検索", dt.Rows.Count.ToString(), textBox1.Text);

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val == 0)
                {
                    //e.CellStyle.ForeColor = Color.Gray;
                    e.Value = null;
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //テーブルクリア
            dt.Clear();
            //グリッド表示クリア
            dataGridView1.DataSource = "";


            GetData();
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            GetData();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void label3_Click(object sender, EventArgs e)
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

            GetData();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                GetData();
            }
        }
    }
}
