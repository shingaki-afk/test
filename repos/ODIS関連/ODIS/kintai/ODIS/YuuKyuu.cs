using Microsoft.VisualBasic;
using Npgsql;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class YuuKyuu : Form
    {
        public YuuKyuu()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            SetTiku();
            SetBumon();
            SetGenba();
            GetData();

            checkBox1.Checked = true;

            Com.InHistory("42_有給年間取得状況一覧", "", "");
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            checkedListBox1.Items.Add("1_本社");
            checkedListBox1.Items.Add("2_那覇");
            checkedListBox1.Items.Add("3_八重山");
            checkedListBox1.Items.Add("4_北部");
            checkedListBox1.Items.Add("5_広域");
            checkedListBox1.Items.Add("6_宮古島");
            checkedListBox1.Items.Add("7_久米島");

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void SetBumon()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox2.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当事務 from dbo.有給年間5日以上取得状況一覧 where 組織CD <> '0000' ";
            if (!checkedListBox1.GetItemChecked(0)) sql += " and 組織CD not like '1%' "; //本社
            if (!checkedListBox1.GetItemChecked(1)) sql += " and 組織CD not like '2%' "; //那覇
            if (!checkedListBox1.GetItemChecked(2)) sql += " and 組織CD not like '3%' "; //八重山
            if (!checkedListBox1.GetItemChecked(3)) sql += " and 組織CD not like '4%' "; //北部
            if (!checkedListBox1.GetItemChecked(4)) sql += " and 組織CD not like '5%' "; //多面
            if (!checkedListBox1.GetItemChecked(5)) sql += " and 組織CD not like '6%' "; //宮古島
            if (!checkedListBox1.GetItemChecked(6)) sql += " and 組織CD not like '7%' "; //久米島

            dt = Com.GetDB(sql);
            
            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["担当事務"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void SetGenba()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct 現場名 from dbo.有給年間5日以上取得状況一覧 where 現場名 <> '0000' ";
            if (!checkedListBox1.GetItemChecked(0)) sql += " and 組織CD not like '1%' "; //本社
            if (!checkedListBox1.GetItemChecked(1)) sql += " and 組織CD not like '2%' "; //那覇
            if (!checkedListBox1.GetItemChecked(2)) sql += " and 組織CD not like '3%' "; //八重山
            if (!checkedListBox1.GetItemChecked(3)) sql += " and 組織CD not like '4%' "; //北部
            if (!checkedListBox1.GetItemChecked(4)) sql += " and 組織CD not like '5%' "; //広域
            if (!checkedListBox1.GetItemChecked(5)) sql += " and 組織CD not like '6%' "; //宮古島
            if (!checkedListBox1.GetItemChecked(6)) sql += " and 組織CD not like '7%' "; //久米島

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 担当事務 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                //checkedListBox3.Items.Add(row["現場名"].ToString() + ' ' + row["現場名"].ToString());
                checkedListBox3.Items.Add(row["現場名"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        private void GetData()
        {
            //ボタン無効化・カーソル変更
            Cursor.Current = Cursors.WaitCursor;

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


            DataTable dt = new DataTable();
            string sql = "select * from dbo.有給年間表示 where 組織CD <> '0000' ";

            //達成
            if (checkBox1.Checked) sql += " and 状況 <> '達成' and 警告 <> '-' ";

            //文字
            sql += result;

            //地区
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        sql += " and 組織CD not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
                    }
                }
            }

            //部門
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    sql += " and 担当事務 <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            //現場
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    sql += " and 現場名 <> '" + checkedListBox3.Items[i].ToString() + "'";
                }
            }

            sql += " order by 直近付与日, 直近付与後使用数";

            dt = Com.GetDB(sql);

            dataGridView1.DataSource = dt;

            //基本情報
            dataGridView1.Columns[0].Width = 60; //社員番号
            dataGridView1.Columns[1].Width = 100; //氏名
            dataGridView1.Columns[2].Width = 50; //地区
            dataGridView1.Columns[3].Width = 90; //組織
            dataGridView1.Columns[4].Width = 180; //現場

            //カウントダウン
            for (int i = 5; i <= 16; i++)
            {
                dataGridView1.Columns[i].Width = 28; //0
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            dataGridView1.Columns[17].Width = 90;
            dataGridView1.Columns[18].Width = 30;
            dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[19].Width = 250;
            dataGridView1.Columns[20].Width = 50;

            dataGridView1.Columns[21].Visible = false;
            dataGridView1.Columns[22].Visible = false;
            dataGridView1.Columns[23].Visible = false;
            dataGridView1.Columns[24].Visible = false;

            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.Beige;

            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[2].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[3].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[4].HeaderCell.Style.BackColor = Color.Beige;

            //dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            //dataGridView1.Columns[14].HeaderCell.Style.BackColor = Color.AntiqueWhite;
            


            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

        }

        //表示期間
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetTiku();
            SetBumon();
            SetGenba();
            GetData();
        }

        //地区の全選択、全解除
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

            SetBumon();
            SetGenba();
            GetData();
        }

 
        //地区のチェック変更イベント
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetBumon();
            SetGenba();
            GetData();
        }

        //部門の全選択、全解除
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
            SetGenba();
            GetData();
        }

        //部門のチェック変更イベント
        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetGenba();
            GetData();
        }

        //現場の全選択、全解除
        private void label1_Click(object sender, EventArgs e)
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

            GetData();
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
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
