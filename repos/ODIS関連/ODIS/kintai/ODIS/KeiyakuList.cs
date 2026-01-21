using Microsoft.VisualBasic;
using Npgsql;
using System;
using System.Data;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KeiyakuList : Form
    {
        private string result;
        private DataTable dt = new DataTable();

        public KeiyakuList()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            //dataGridView3.Font = new Font(dataGridView3.Font.Name, 10);

            // 選択モードを行単位での選択のみにする
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            checkedListBox1.Items.Add("契約固定");
            checkedListBox1.Items.Add("契約臨時");
            checkedListBox1.SetItemChecked(0, true);
            checkedListBox1.SetItemChecked(1, true);

            checkedListBox2.Items.Add("自社");
            checkedListBox2.Items.Add("外注");
            checkedListBox2.SetItemChecked(0, true);
            checkedListBox2.SetItemChecked(1, true);

            checkedListBox4.Items.Add("1_売上");
            checkedListBox4.Items.Add("2_実施");
            checkedListBox4.SetItemChecked(0, true);
            checkedListBox4.SetItemChecked(1, true);

            SetBumon();
            GetUriageData();

            Com.InHistory("13_契約一覧", "", "");
        }

        private void SetBumon()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct 部門コード, 部門 from kpcp01.\"CostomKeiyakuList\" order by 部門コード";

            dt = Com.GetPosDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["部門コード"].ToString() + ' ' + row["部門"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        private void GetDisp()
        {
            //検索文字列処理
            ResultStr();

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            DataTable Disp = new DataTable();
            Disp.Columns.Add("連番", typeof(string));
            Disp.Columns.Add("工事コード", typeof(string));
            Disp.Columns.Add("工事名", typeof(string));
            Disp.Columns.Add("契約名", typeof(string));
            Disp.Columns.Add("契約区分", typeof(string));
            Disp.Columns.Add("区分", typeof(string));
            Disp.Columns.Add("部門コード", typeof(string));
            Disp.Columns.Add("部門", typeof(string));
            Disp.Columns.Add("額名", typeof(string));

            Disp.Columns.Add("4月", typeof(decimal));
            Disp.Columns.Add("5月", typeof(decimal));
            Disp.Columns.Add("6月", typeof(decimal));
            Disp.Columns.Add("7月", typeof(decimal));
            Disp.Columns.Add("8月", typeof(decimal));
            Disp.Columns.Add("9月", typeof(decimal));
            Disp.Columns.Add("10月", typeof(decimal));
            Disp.Columns.Add("11月", typeof(decimal));
            Disp.Columns.Add("12月", typeof(decimal));
            Disp.Columns.Add("1月", typeof(decimal));
            Disp.Columns.Add("2月", typeof(decimal));
            Disp.Columns.Add("3月", typeof(decimal));
            Disp.Columns.Add("備考", typeof(string));

            foreach (DataRow row in dtrow)
            {
                DataRow nr = Disp.NewRow();
                nr["連番"] = row["連番"];
                nr["工事コード"] = row["工事コード"];
                nr["工事名"] = row["工事名"];
                nr["契約名"] = row["契約名"];
                nr["契約区分"] = row["契約区分"];
                nr["区分"] = row["区分"];
                nr["部門コード"] = row["部門コード"];
                nr["部門"] = row["部門"];
                nr["額名"] = row["額名"];

                nr["4月"] = row["4月"];
                nr["5月"] = row["5月"];
                nr["6月"] = row["6月"];
                nr["7月"] = row["7月"];
                nr["8月"] = row["8月"];
                nr["9月"] = row["9月"];
                nr["10月"] = row["10月"];
                nr["11月"] = row["11月"];
                nr["12月"] = row["12月"];
                nr["1月"] = row["1月"];
                nr["2月"] = row["2月"];
                nr["3月"] = row["3月"];
                nr["備考"] = row["備考"];

                Disp.Rows.Add(nr);
            }

            dataGridView3.DataSource = Disp;

            int ct = Disp.Columns.Count;

            dataGridView3.Columns[0].Width = 40;
            dataGridView3.Columns[1].Width = 40;
            dataGridView3.Columns[2].Width = 200;
            dataGridView3.Columns[3].Width = 200;
            dataGridView3.Columns[4].Width = 60;
            dataGridView3.Columns[5].Width = 50;
            dataGridView3.Columns[6].Width = 50;
            dataGridView3.Columns[7].Width = 100;

            for (int i = 7; i < ct - 1; i++)
            {
                dataGridView3.Columns[i].Width = 60;
            }



            //ヘッダーの中央表示
            for (int i = 0; i < ct; i++)
            {
                dataGridView3.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //三桁区切り表示
            for (int i = 0; i < ct; i++)
            {
                dataGridView3.Columns[i].DefaultCellStyle.Format = "#,0";
            }

            //表示位置
            for (int i = 0; i < ct; i++)
            {
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            dataGridView3.Columns[ct - 1].Width = 350;
            dataGridView3.Columns[ct - 1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

        }

        private void ResultStr()
        {
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            //TODO
            result = "";

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
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
                    result += " and 区分 <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            //部門コード
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    result += " and 部門コード <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                }
            }

            //額名
            for (int i = 0; i < checkedListBox4.Items.Count; i++)
            {
                if (!checkedListBox4.GetItemChecked(i))
                {
                    result += " and 額名 <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                }
            }

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }


        }

        private void GetUriageData()
        {
            dataGridView3.DataSource = null;

            string sql = "";
            sql = "select * from kpcp01.\"CostomKeiyakuList_serch\" ";
            //sql += " where 契約区分 = '契約固定' ";
            sql += " order by 工事コード, 連番, 部門コード, 額名 ";

            dt = Com.GetPosDB(sql);
            

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetDisp();
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
        }

        private void label1_Click(object sender, EventArgs e)
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

        private void label2_Click(object sender, EventArgs e)
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
            textBox1.Text = "";
        }

        private void label4_Click(object sender, EventArgs e)
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
                //カーソル変更
                Cursor.Current = Cursors.WaitCursor;

                GetDisp();

                //カーソル変更・メッセージキュー処理
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }
        }
    }
}
