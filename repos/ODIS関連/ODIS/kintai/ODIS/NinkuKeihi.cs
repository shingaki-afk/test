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
    public partial class NinkuKeihi : Form
    {
        public NinkuKeihi()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            kizyunym.Value = DateTime.Today;
            syouyo.Value = Convert.ToDecimal("1.5");
            yuukyuu.Value = 10;
            hukuri.Value = 9000;
            kyouiku.Value = 18000;
            hihuku.Value = 20000;
            etc.Value = 2000;

            SetTiku();
            SetBumon();
            SetYakusyoku();

            GetData();

            Com.InHistory("人工単価", "", "");
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.s社員基本情報_期間指定('" + Convert.ToDateTime(kizyunym.Value).ToString("yyyy/MM/dd") + "') where 在籍区分 <> '9' and 現場名 not like '事務所%' and 給与支給区分 in ('C1') ";
            
            if (keiyakuck.Checked)
            { 
                sql += "and 契約社員 is null ";
            }

            sql += "order by 担当区分";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox1.Items.Add(row["担当区分"]);
            }

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
            string sql = "select distinct 職種 from dbo.s社員基本情報_期間指定('" + Convert.ToDateTime(kizyunym.Value).ToString("yyyy/MM/dd") + "') where 在籍区分 <> '9' and 現場名 not like '事務所%' and 給与支給区分 in ('C1') ";

            if (keiyakuck.Checked)
            {
                sql += "and 契約社員 is null ";
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            sql += " order by 職種";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["職種"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void SetYakusyoku()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct 役職CD, 役職名 from dbo.s社員基本情報_期間指定('" + Convert.ToDateTime(kizyunym.Value).ToString("yyyy/MM/dd") + "') where 在籍区分 <> '9' and 現場名 not like '事務所%' and 現場名 not like '事務所%' and 給与支給区分 in ('C1') ";

            if (keiyakuck.Checked)
            {
                sql += "and 契約社員 is null ";
            }
            
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += " order by 役職CD";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["役職名"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        private void GetData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "n人工取得";
                        //Cmd.CommandText = "n人工取得_平均"; 
                        Cmd.CommandTimeout = 600;

                        Cmd.Parameters.Add(new SqlParameter("基準年月", SqlDbType.VarChar));
                        Cmd.Parameters["基準年月"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("賞与引当", SqlDbType.VarChar));
                        Cmd.Parameters["賞与引当"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("有給", SqlDbType.VarChar));
                        Cmd.Parameters["有給"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("福利厚生", SqlDbType.VarChar));
                        Cmd.Parameters["福利厚生"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("教育訓練研修", SqlDbType.VarChar));
                        Cmd.Parameters["教育訓練研修"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("被服", SqlDbType.VarChar));
                        Cmd.Parameters["被服"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("その他", SqlDbType.VarChar));
                        Cmd.Parameters["その他"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("基本給加算", SqlDbType.VarChar));
                        Cmd.Parameters["基本給加算"].Direction = ParameterDirection.Input;
                        
                        Cmd.Parameters.Add(new SqlParameter("諸手当加算", SqlDbType.VarChar));
                        Cmd.Parameters["諸手当加算"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("深夜割増", SqlDbType.VarChar));
                        Cmd.Parameters["深夜割増"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("search", SqlDbType.VarChar));
                        Cmd.Parameters["search"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["基準年月"].Value = Convert.ToDateTime(kizyunym.Value).ToString("yyyy/MM/dd");
                        Cmd.Parameters["賞与引当"].Value = syouyo.Value.ToString();
                        Cmd.Parameters["有給"].Value = yuukyuu.Value.ToString();
                        Cmd.Parameters["福利厚生"].Value = hukuri.Value.ToString();
                        Cmd.Parameters["教育訓練研修"].Value = kyouiku.Value.ToString();
                        Cmd.Parameters["被服"].Value = hihuku.Value.ToString();
                        Cmd.Parameters["その他"].Value = etc.Value.ToString();

                        Cmd.Parameters["基本給加算"].Value = addkihon.Value.ToString();
                        Cmd.Parameters["諸手当加算"].Value = addsyoteate.Value.ToString();
                        Cmd.Parameters["深夜割増"].Value = addshinya.Value.ToString();

                        string search = "";

                        if (keiyakuck.Checked)
                        {
                            search += "and 契約社員 is null ";
                        }

                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            if (!checkedListBox1.GetItemChecked(i)) search += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
                        }

                        for (int i = 0; i < checkedListBox2.Items.Count; i++)
                        {
                            if (!checkedListBox2.GetItemChecked(i)) search += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
                        }

                        for (int i = 0; i < checkedListBox3.Items.Count; i++)
                        {
                            if (!checkedListBox3.GetItemChecked(i)) search += " and 役職名 <> '" + checkedListBox3.Items[i].ToString() + "' ";
                        }

                        Cmd.Parameters["search"].Value = search;
                        //Cmd.Parameters["search"].Value = " and 役職名 = '主任'";


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



            //役職コード削除
            dt.Columns.RemoveAt(0);


            DataTable repdt = new DataTable();
            repdt = replaceDataTable(dt);

            dataGridView2.DataSource = repdt;

            dataGridView2.Columns[0].Width = 180;
            dataGridView2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 1; i < repdt.Columns.Count; i++)
            {
                dataGridView2.Columns[i].Width = 60;
                dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[i].DefaultCellStyle.Format = "#,0";
            }

            //現場経費 科目別合計額
            string sql = "";
            DataTable gkeihi = new DataTable();

            sql = "select 科目コード, 科目名, sum(金額) as 合計額 from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql += "where 伝票日付 between '20230401' and '20240331' ";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += "and 科目コード > '8010' and 摘要文 not like '給与__月支給分%' and 摘要文 not like '賞与___0%' and 科目コード not in ('8215','8346') and 摘要文 not like '%労災保険概算納付%' and 摘要文 not like '%月分社会保険料差額分（子ども子育て拠出金%' ";
            sql += "and 摘要文 not like '%全友協沖縄支部へ活動資金助成金を期末剰余金返金%' and 摘要文 not like '全友協沖縄支部より%年度資金剰余金を戻入支払' and 摘要文 not like '%全友協沖縄支部へ%月分会費として%' ";
            sql += "and 科目コード not in ('8251','8281') and 工種名 not like '臨時%' group by 科目コード, 科目名 order by 科目コード ";

            gkeihi = Com.GetDB(sql);
            dataGridView1.DataSource = gkeihi;

        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //dataGridView2.Rows[0].DefaultCellStyle.BackColor = Color.LightGray;

            for (int i = 1; i <= 4; i++)
            {
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.MediumAquamarine;
            }

            for (int i = 5; i <= 16; i++)
            {
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Aquamarine;
            }

            dataGridView2.Rows[17].DefaultCellStyle.BackColor = Color.LightSeaGreen; //直接人件費合計

            for (int i = 18; i <= 25; i++)
            {
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue; //間接人件費①
            }

            dataGridView2.Rows[26].DefaultCellStyle.BackColor = Color.SteelBlue; //間接人件費合計
            //dataGridView2.Rows[27].DefaultCellStyle.BackColor = Color.RoyalBlue; //人件費合計

            //for (int i = 28; i <= 29; i++)
            //{
            //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;　//間接人件費②
            //}






        }



        private DataTable replaceDataTable(DataTable dt)
        {
            DataTable retDt = new DataTable();
            DataRow row = null;
            try
            {
                // 戻り値のDataTable作成

                //列名
                retDt.Columns.Add((string)dt.Columns[0].ColumnName, typeof(string));


                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    retDt.Columns.Add((string)dt.Rows[j].ItemArray[0], typeof(decimal));
                }


                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    row = retDt.NewRow();
                    row[(string)dt.Columns[0].ColumnName] = dt.Columns[i].ColumnName;
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        row[(string)dt.Rows[j].ItemArray[0]] = dt.Rows[j].ItemArray[i];
                    }

                    retDt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return retDt;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            SetBumon();
            SetYakusyoku();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            SetYakusyoku();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();

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

            SetBumon();
            SetYakusyoku();
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

            SetYakusyoku();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

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
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void keiyakuck_CheckedChanged(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            keiyakuck.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            SetTiku();
            SetBumon();
            SetYakusyoku();
            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            keiyakuck.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            SetTiku();
            SetBumon();
            SetYakusyoku();
            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;
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

            DataTable dt = new DataTable();
            //string sql = "";



            //現場経費 詳細
            string sql2 = "";
            DataTable gksyousai = new DataTable();

            sql2 = "select * from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql2 += "where 伝票日付 between '20230401' and '20240331' ";


            sql2 += "and 科目コード = '" + row + "' ";
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql2 += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql2 += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql2 += "and 科目コード > '8010' and 摘要文 not like '給与__月支給分%' and 摘要文 not like '賞与___0%' and 科目コード not in ('8215','8346') and 摘要文 not like '%労災保険概算納付%' and 摘要文 not like '%月分社会保険料差額分（子ども子育て拠出金%' ";
            sql2 += "and 摘要文 not like '%全友協沖縄支部へ活動資金助成金を期末剰余金返金%' and 摘要文 not like '全友協沖縄支部より%年度資金剰余金を戻入支払' and 摘要文 not like '%全友協沖縄支部へ%月分会費として%' ";
            sql2 += "and 科目コード not in ('8251','8281') and 工種名 not like '臨時%' order by 科目コード ";

            gksyousai = Com.GetDB(sql2);
            dataGridView3.DataSource = gksyousai;






            //sql += "select 科目名, 部門名, 現場名, 金額, 摘要文, 取引先名, 伝票日付, 伝票番号, 科目コード, 部門コード, 現場コード, 担当事務, 担当区分, 消費税額, 税区分コード, 税区分名,工種名　from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            //sql += " where 科目コード between '8000' and '9900' ";
            //if (col == "年間")
            //{
            //    sql += " and 伝票日付 between '" + maey + "0401' and '" + atoy + "0331'";
            //}
            //else
            //{
            //    if (col == "01月" || col == "02月" || col == "03月")
            //    {
            //        sql += " and 伝票日付 like '" + atoy + col.Replace("月", "") + "%'";
            //    }
            //    else
            //    {
            //        sql += " and 伝票日付 like '" + maey + col.Replace("月", "") + "%'";
            //    }
            //}

            //if (row == "8299") //現場経費
            //{
            //    sql += " and 科目コード between '8200' and '8298'";
            //}
            //else if (row == "9980") //管理経費
            //{
            //    sql += " and 科目コード between '8300' and '8999'";
            //}
            //else if (row == "9990") //全体経費
            //{
            //    sql += " and 科目コード between '8200' and '8999'";
            //}
            //else
            //{
            //    sql += " and 科目コード = '" + row + "'";
            //}

            //sql += GetTSG();


            ////sql += " order by 科目コード, 部門コード, 現場コード,金額";
            //sql += " order by 金額 desc";

            //dt = Com.GetDB(sql);



            //dataGridView2.DataSource = dt;

            ////売上
            //if (row.Substring(0, 1) == "0")
            //{
            //    //MessageBox.Show("売上！");
            //    //TODO 
            //}
            //else
            //{
            //    dataGridView2.Columns[0].Width = 120;//科目名
            //    dataGridView2.Columns[1].Width = 120;//部門名
            //    dataGridView2.Columns[2].Width = 250;//現場名
            //    dataGridView2.Columns[3].Width = 70;//金額
            //    dataGridView2.Columns[4].Width = 400;//摘要
            //    dataGridView2.Columns[5].Width = 250;//取引先
            //    dataGridView2.Columns[6].Width = 60;//
            //    dataGridView2.Columns[7].Width = 60;//
            //    dataGridView2.Columns[8].Width = 60;//
            //    dataGridView2.Columns[9].Width = 60;//
            //    dataGridView2.Columns[10].Width = 60;//
            //    dataGridView2.Columns[11].Width = 60;//

            //    //金額右寄
            //    dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    //dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //    //三桁区切り表示
            //    dataGridView2.Columns[3].DefaultCellStyle.Format = "#,0";
            //    //dataGridView2.Columns[4].DefaultCellStyle.Format = "#,0";
            //}
        }


    }
}
