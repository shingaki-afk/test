using Npgsql;
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


    public partial class Kamoku : Form
    {
        //期が変わるタイミングで変更する必要がある
        private string maey = "2019";
        private string atoy = "2020";

        public Kamoku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            checkedListBox1.Items.Add("1_本社");
            checkedListBox1.Items.Add("2_那覇");
            checkedListBox1.Items.Add("3_八重山");
            checkedListBox1.Items.Add("4_北部");

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }

            SetBumon();
            


            //フォントサイズの変更
            //dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            GetData();

            //dataGridView1でセル、行、列が複数選択されないようにする
            //dataGridView1.MultiSelect = false;

        }

        private void SetBumon()
        {
            DataTable dt = new DataTable();
            int nRet;

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    string sql = "select distinct bumonkubun from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8200' and '8600' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' order by bumonkubun";
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["bumonkubun"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void GetData()
        {
            //ボタン無効化・カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            DataTable dt = new DataTable();
            int nRet;

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    string sql = "";
                    sql += "select * from (";
                    //sql += "select kamokucode as \"CD\", max(case when uchiwakecode = '0000' then kamokuname else '' end) as 科目名";
                    sql += "select kamokucode as \"CD\", max(kamokuname) as 科目名";
                    sql += " , sum(case when suitouymd like '" + maey + "04%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "04%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"04月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "05%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "05%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"05月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "06%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "06%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"06月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "07%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "07%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"07月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "08%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "08%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"08月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "09%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "09%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"09月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "10%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "10%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"10月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "11%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "11%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"11月\"";
                    sql += " , sum(case when suitouymd like '" + maey + "12%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "12%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"12月\"";
                    sql += " , sum(case when suitouymd like '" + atoy + "01%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "01%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"01月\"";
                    sql += " , sum(case when suitouymd like '" + atoy + "02%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "02%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"02月\"";
                    sql += " , sum(case when suitouymd like '" + atoy + "03%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "03%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"03月\"";
                    sql += " , sum(case when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"年間\"";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8230' and '8600' ";
                    //地区
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (!checkedListBox1.GetItemChecked(i))
                        {
                            if (!checkedListBox1.GetItemChecked(i))
                            {
                                sql += " and bumoncode not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
                            }
                        }
                    }

                    //部門
                    for (int i = 0; i < checkedListBox2.Items.Count; i++)
                    {
                        if (!checkedListBox2.GetItemChecked(i))
                        {
                            sql += " and bumonkubun <> '" + checkedListBox2.Items[i].ToString() + "'";
                         }
                    }

                    sql += " group by kamokucode";
                    sql += " having sum(case when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) > 0";

                    sql += " union all select '9999' as \"CD\", '【合計】' as 科目名 ";
                    sql += " , sum(case when suitouymd like '" + maey + "04%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "04%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"04月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "05%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "05%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"05月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "06%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "06%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"06月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "07%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "07%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"07月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "08%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "08%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"08月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "09%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "09%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"09月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "10%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "10%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"10月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "11%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "11%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"11月\" ";
                    sql += " , sum(case when suitouymd like '" + maey + "12%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + maey + "12%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"12月\" ";
                    sql += " , sum(case when suitouymd like '" + atoy + "01%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "01%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"01月\" ";
                    sql += " , sum(case when suitouymd like '" + atoy + "02%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "02%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"02月\" ";
                    sql += " , sum(case when suitouymd like '" + atoy + "03%' and taisyakukubunb = '1' then denpyoukingaku when suitouymd like '" + atoy + "03%' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"03月\" ";
                    sql += " , sum(case when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) as \"年間\" ";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8230' and '8600' ";

                    //地区
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (!checkedListBox1.GetItemChecked(i))
                        {
                            sql += " and bumoncode not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
                        }
                    }

                    //部門
                    for (int i = 0; i < checkedListBox2.Items.Count; i++)
                    {
                        if (!checkedListBox2.GetItemChecked(i))
                        {
                            sql += " and bumonkubun <> '" + checkedListBox2.Items[i].ToString() + "'";
                        }
                    }


                    sql += " ) temp order by \"CD\"";
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            dataGridView1.DataSource = dt;


            ///comboBox1.

            //表示処理
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 120;

            for (int i = 2; i < 15; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }

                dataGridView1.Columns[i].Width = 60;

                //三桁区切り表示
                dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.Beige;

            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.Beige;

            dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[14].HeaderCell.Style.BackColor = Color.AntiqueWhite;


            Com.InHistory("科目別損益", "", "");

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //ソートエラー対応
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;

            //科目コード、科目名はスルー
            if (dataGridView1.CurrentCell.ColumnIndex < 2) return;

            string row = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            string col = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderCell.Value.ToString();
            //MessageBox.Show(row + " " + col);


            DataTable dt = new DataTable();
            int nRet;

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    string sql = ""; 
                    sql += "select kamokuname as 科目名, bumonname as 部門名, koujiname as 現場名,";
                    sql += " case when taisyakukubunb = '1' then denpyoukingaku else denpyoukingaku * -1 end as 金額, ";
                    sql += " tekiyou as 摘要,torihikisakiname as 取引先名";
                    sql += " , suitouymd as 日付, denpyounumber as 伝票番号, gyounumber as 行番";
                    sql += " , inputtantousyaname as 入力者, registrationtantousyaname as 更新者";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8230' and '8600'";
                    if (col == "年間")
                    {
                        sql += " and suitouymd between '" + maey + "0401' and '" + atoy + "0331'";
                    }
                    else
                    { 
                        if (col == "01" || col == "02" || col == "03")
                        {
                            sql += " and suitouymd like '" + atoy + col.Replace("月", "") + "%'";
                        }
                        else
                        { 
                            sql += " and suitouymd like '" + maey + col.Replace("月","") + "%'";
                        }
                    }
                    if (row == "9999")
                    {

                    }
                    else
                    { 
                        sql += " and kamokucode = '" + row + "'";
                    }

                    //地区
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (!checkedListBox1.GetItemChecked(i))
                        {
                            sql += " and bumoncode not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
                        }
                    }

                    //部門
                    for (int i = 0; i < checkedListBox2.Items.Count; i++)
                    {
                        if (!checkedListBox2.GetItemChecked(i))
                        {
                            sql += " and bumonkubun <> '" + checkedListBox2.Items[i].ToString() + "'";
                        }
                    }

                    sql += " order by kamokucode, bumoncode, koujicode";

                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            dataGridView2.DataSource = dt;

            dataGridView2.Columns[0].Width = 100;//科目名
            dataGridView2.Columns[1].Width = 100;//部門名
            dataGridView2.Columns[2].Width = 250;//現場名
            dataGridView2.Columns[3].Width =  70;//金額
            dataGridView2.Columns[4].Width = 400;//摘要
            dataGridView2.Columns[5].Width = 250;//取引先

            //金額右寄
            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //三桁区切り表示
            dataGridView2.Columns[3].DefaultCellStyle.Format = "#,0";
            dataGridView2.Columns[4].DefaultCellStyle.Format = "#,0";
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

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            //部門と連動
            //地区
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    //チェックが外れている地区は非表示にする

                }
            }

            GetData();
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

            GetData();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void Kamoku_Load(object sender, EventArgs e)
        {

        }
    }
}
