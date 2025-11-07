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
    public partial class Tekisei : Form
    {
        private SqlConnection Cn;

        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public Tekisei()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            SetBuomn();
            SetSyokusyu();

            //メインの更新
            GetMainData();

            //dataGridView1.ColumnHeadersHeight = 100;
            //dataGridView5.ColumnHeadersHeight = 100;

            Com.InHistory("25_適正人員入力", "", "");

        }

        private void GetMainData()
        {
            checkedListBox1.Enabled = false;
            checkedListBox2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //リセット!
            dataGridView5.DataSource = null;
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            //編集
            string sql = "select a.組織名, a.現場名,  ";
            sql += "a.パートアルバイト, a.サブチーフ, a.チーフ, a.日給, a.係員, a.副主任, a.主任, a.電三, a.管理職, ";
            sql += "(isnull(a.[パートアルバイト], 0) + isnull(a.[サブチーフ], 0) + isnull(a.[チーフ], 0) + isnull(a.日給, 0) + isnull(a.係員, 0) + isnull(a.副主任, 0) + isnull(a.主任, 0) + isnull(a.電三, 0) + isnull(a.管理職, 0)) as [①合計] ";
            sql += " , a.組織CD, a.現場CD from dbo.担当テーブル a where a.定員数 > 0 ";

            //一覧
            string sqlgen = "select * from dbo.k欠員一覧 where 担当区分 like '%%' ";
            //一覧合計
            string sqlgensum = "";

            sqlgensum += "select '合計' as 組織名, '' as 現場名,";
            sqlgensum += "sum([①パート]) as [①パート], sum([①サブチーフ]) as [①サブチーフ], sum([①チーフ]) as [①チーフ], sum([①日給]) as [①日給], sum([①係員]) as [①係員], sum([①副主任]) as [①副主任], sum([①主任]) as [①主任], sum([①電三]) as [①電三], sum([①管理職]) as [①管理職], sum([①合計]) as [①合計], ";
            sqlgensum += "sum([②パート]) as [②パート], sum([②サブチーフ]) as [②サブチーフ], sum([②チーフ]) as [②チーフ], sum([②日給]) as [②日給], sum([②係員]) as [②係員], sum([②副主任]) as [②副主任], sum([②主任]) as [②主任], sum([②電三]) as [②電三], sum([②管理職]) as [②管理職], sum([②合計]) as [②合計], ";
            sqlgensum += "sum([③パート]) as [③パート], sum([③サブチーフ]) as [③サブチーフ], sum([③チーフ]) as [③チーフ], sum([③日給]) as [③日給], sum([③係員]) as [③係員], sum([③副主任]) as [③副主任], sum([③主任]) as [③主任], sum([③電三]) as [③電三], sum([③管理職]) as [③管理職], sum([③合計]) as [③合計]";
            sqlgensum += "from dbo.k欠員一覧 where 担当区分 like '%%' ";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    sql += " and isnull(担当区分,'') <> '" + checkedListBox1.Items[i].ToString() + "'";
                    sqlgen += " and isnull(担当区分,'') <> '" + checkedListBox1.Items[i].ToString() + "'";
                    sqlgensum += " and isnull(担当区分,'') <> '" + checkedListBox1.Items[i].ToString() + "'";
                }
            }

            //職種
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    sql += " and isnull(担当事務,'') <> '" + checkedListBox2.Items[i].ToString() + "'";
                    sqlgen += " and isnull(担当事務,'') <> '" + checkedListBox2.Items[i].ToString() + "'";
                    sqlgensum += " and isnull(担当事務,'') <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            sql += " order by a.組織CD, a.現場CD ";
            sqlgen += " order by 組織CD, 現場CD ";

            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView5.DataSource = dt;

            dataGridView5.Columns[0].ReadOnly = true;
            dataGridView5.Columns[1].ReadOnly = true;
            dataGridView5.Columns[11].ReadOnly = true;

            dataGridView5.Columns[12].Visible = false;
            dataGridView5.Columns[13].Visible = false;

            dataGridView5.Columns[0].Width = 90;
            dataGridView5.Columns[1].Width = 150;

            for (int i = 2; i < dataGridView5.Columns.Count; i++)
            {
                dataGridView5.Columns[i].Width = 30;
                dataGridView5.Columns[i].DefaultCellStyle.BackColor = Color.AntiqueWhite;
                dataGridView5.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //ヘッダーの色変更
                dataGridView5.Columns[i].HeaderCell.Style.BackColor = Color.AliceBlue;
            }

            dataGridView5.Columns[11].DefaultCellStyle.BackColor = Color.AliceBlue;




            dataGridView1.DataSource = Com.GetDB(sqlgen);
            dataGridView8.DataSource = Com.GetDB(sqlgensum);

            dataGridView1.Columns[0].Width = 90;
            dataGridView1.Columns[1].Width = 150;
            dataGridView8.Columns[0].Width = 90;
            dataGridView8.Columns[1].Width = 150;

            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].Width = 30;
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                if (i < 32)
                { 
                dataGridView8.Columns[i].Width = 30;
                dataGridView8.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }

                if (i < 12)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.AliceBlue;
                    dataGridView8.Columns[i].DefaultCellStyle.BackColor = Color.AliceBlue;
                    dataGridView8.Columns[i].HeaderCell.Style.BackColor = Color.AliceBlue;
                }
                else if (i < 22)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.MistyRose;
                    dataGridView8.Columns[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    dataGridView8.Columns[i].HeaderCell.Style.BackColor = Color.MistyRose;
                }
                else if (i < 32)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Beige;
                    dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Beige;
                    dataGridView8.Columns[i].DefaultCellStyle.BackColor = Color.Beige;
                    dataGridView8.Columns[i].HeaderCell.Style.BackColor = Color.Beige;
                }
            }

            dataGridView1.Columns[32].Visible = false;
            dataGridView1.Columns[33].Visible = false;
            dataGridView1.Columns[34].Visible = false;
            dataGridView1.Columns[35].Visible = false;


            //カーソル変更・メッセージキュー処理・コンボボックス有効化

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            checkedListBox1.Enabled = true;
            checkedListBox2.Enabled = true;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;


            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                GetMainData();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }



            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;
        }

        private void SetBuomn()
        {
            checkedListBox1.Items.Clear();

            DataTable bumondt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.担当テーブル where 定員数 > 0 order by 担当区分 ";
            bumondt = Com.GetDB(sql);

            foreach (DataRow row in bumondt.Rows)
            {
                checkedListBox1.Items.Add(row["担当区分"]);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void SetSyokusyu()
        {
            checkedListBox2.Items.Clear();

            DataTable syokusyudt = new DataTable();
            string sql = "select distinct 担当事務 from dbo.担当テーブル where 定員数 > 0 ";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            sql += " order by 担当事務 ";

            syokusyudt = Com.GetDB(sql);

            foreach (DataRow row in syokusyudt.Rows)
            {
                checkedListBox2.Items.Add(row["担当事務"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
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

            //SetBuomn();
            SetSyokusyu();
            GetMainData();
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

            GetMainData();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSyokusyu();
            GetMainData();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetMainData();
        }
    }
}
