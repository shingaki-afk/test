using Npgsql;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Kamoku : Form
    {
        //YM
        private string maey = "";
        private string atoy = "";

        private string maeyex = "";
        private string atoyex = "";

        //科目コードの幅
        //private string kamos = "8200";
        //private string kamoe = "9900";

        public Kamoku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            dataGridView1.CellFormatting += dataGridView1_CellFormatting;
            dataGridView1.DataError += dataGridView1_DataError;

            comboBox1.Items.Add("54期(2025～2026)");
            comboBox1.Items.Add("53期(2024～2025)");
            comboBox1.Items.Add("52期(2023～2024)");
            //comboBox1.Items.Add("51期(2022～2023)");

            comboBox1.SelectedIndex = 0;

            SetTiku();
            SetBumon();
            SetGenba();

            //フォントサイズの変更
            //dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            GetData();

            //dataGridView1でセル、行、列が複数選択されないようにする
            //dataGridView1.MultiSelect = false;
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' and 科目コード between '8000' and '9900' order by 担当区分 ";
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
            checkedListBox2.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 職種 from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' and 科目コード between '8000' and '9900'";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            sql += " order by 職種 ";

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

        private void SetGenba()
        {
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            //売上
            string sql = "select distinct 現場コード,現場名 from dbo.PCA会計仕訳データ_貸借区分_科目別損益用　where 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' and 科目コード between '8000' and '9900'";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += " order by 現場コード,現場名 ";


            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["現場コード"].ToString() + ' ' + row["現場名"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        private string GetTSG()
        {
            string sql = "";
            //部門
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    sql += " and isnull(担当区分,'') <> '" + checkedListBox1.Items[i].ToString() + "'";
                }
            }

            //職種
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    sql += " and isnull(職種,'') <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            //現場
            int itemcount = checkedListBox3.Items.Count; //項目数合計
            int ckcount = checkedListBox3.CheckedItems.Count; //チェック項目数
            //アイテム数合計(200)の半分(100)より、チェック数が少ない場合はチェック無を条件にsql作成
            if (itemcount / 2 > ckcount)
            {
                string sql3 = "";
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (checkedListBox3.GetItemChecked(i))
                    {
                        if (checkedListBox3.Items[i].ToString().Length < 5)
                        {
                            sql3 += " or isnull(現場コード,'') = '' ";
                        }
                        else
                        {
                            sql3 += " or isnull(現場コード,'') = '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                        }
                    }
                }
                if (sql3.Length > 0) sql += "and ( " + sql3.Substring(4) + " ) ";

            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (!checkedListBox3.GetItemChecked(i))
                    {
                        if (checkedListBox3.Items[i].ToString().Length < 5)
                        {
                            sql += " and isnull(現場コード,'') <> '' ";
                        }
                        else
                        {
                            sql += " and isnull(現場コード,'') <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                        }
                    }
                }
            }

            return sql;
        }

        private void GetData()
        {
            //ボタン無効化・カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            DataTable dt = new DataTable();

            //期間指定
            maey = comboBox1.SelectedItem.ToString().Substring(4, 4);
            atoy = comboBox1.SelectedItem.ToString().Substring(9, 4);

            maeyex = (Convert.ToInt16(maey) - 1).ToString();
            atoyex = (Convert.ToInt16(atoy) - 1).ToString();

            DataTable dt2 = new DataTable();

            //月平均算出時の分母
            dt2 = Com.GetDB("select count(distinct left(伝票日付, 6)) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' and 科目コード = '8010'");
            int i2 = Convert.ToInt32(dt2.Rows[0][0].ToString());
            //エラー防止
            if (i2 == 0) i2 = 1;

            DataTable dt3 = new DataTable();

            //売上比算出時の分母(売上)
            dt3 = Com.GetDB("select sum(金額) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' and 科目コード = '8010'");

            //エラー対応
            double uriage = 0;

            if (dt3.Rows[0][0].ToString() == "")
            {
                uriage = 1;
            }
            else
            {
                uriage = Convert.ToDouble(dt3.Rows[0][0].ToString());
            }

            //地区と部門と現場で絞り込みする
            string sql2 = GetTSG();
            string sql = "";

            sql += " select 科目コード, max(科目名) as 科目名,  ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";

            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";

            sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            //sql += " '' as '月平均(目論見)' ";

            sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";

            if (checkBox1.Checked)
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8000' and '9900'";
            }
            else
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8000' and '9900'";
            }

            if (checkBox3.Checked)
            {
                sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            }

            sql += sql2;
            sql += " group by 科目コード ";
            sql += " union all ";
            sql += " select '8300' as 科目コード,'【原価合計】' as 科目名, ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";

            if (checkBox1.Checked)
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8212' and '8298' ";
            }
            else
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8212' and '8298' ";
            }

            if (checkBox3.Checked)
            {
                sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            }

            sql += sql2;

            //sql += " union all ";
            //sql += " select '9980' as 科目コード,'【管理経費合計】' as 科目名, ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            //sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            //sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            //sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            //sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            //sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            //if (checkBox1.Checked)
            //{
            //    sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8300' and '8999' ";
            //}
            //else
            //{
            //    sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8300' and '8999' ";
            //}
            //if (checkBox3.Checked)
            //{
            //    sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            //}
            //sql += sql2;

            sql += " union all ";
            sql += " select '9990' as 科目コード,'【経費総計】' as 科目名, ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";

            if (checkBox1.Checked)
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8200' and '9900'";
            }
            else
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8200' and '9900'";
            }

            if (checkBox3.Checked)
            {
                sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            }

            sql += sql2;


            //8600販管費合計(8310～8599)
            sql += " union all ";
            sql += " select '8600' as 科目コード,'【販管費合計】' as 科目名, ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";


            if (checkBox1.Checked)
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8310' and '8599'";
            }
            else
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8310' and '8599'";
            }

            if (checkBox3.Checked)
            {
                sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            }

            sql += sql2;

            //9000営業外損益合計(8600～8999)
            sql += " union all ";
            sql += " select '9000' as 科目コード,'【営業外損益合計】' as 科目名, ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";


            if (checkBox1.Checked)
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8600' and '8999'";
            }
            else
            {
                sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '8600' and '8999'";
            }

            if (checkBox3.Checked)
            {
                sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            }

            sql += sql2;



            ////9500特別損益合計(9001～9499)
            //sql += " union all ";
            //sql += " select '9500' as 科目コード,'【特別損益合計】' as 科目名, ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "04%' then 金額 else 0 end ),0) as '04月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "05%' then 金額 else 0 end ),0) as '05月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "06%' then 金額 else 0 end ),0) as '06月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "07%' then 金額 else 0 end ),0) as '07月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "08%' then 金額 else 0 end ),0) as '08月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "09%' then 金額 else 0 end ),0) as '09月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "10%' then 金額 else 0 end ),0) as '10月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "11%' then 金額 else 0 end ),0) as '11月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + maey + "12%' then 金額 else 0 end ),0) as '12月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "01%' then 金額 else 0 end ),0) as '01月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "02%' then 金額 else 0 end ),0) as '02月', ";
            //sql += " isnull(sum(case when 伝票日付 like '" + atoy + "03%' then 金額 else 0 end ),0) as '03月', ";
            //sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) as '年間', ";
            //sql += " isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) * 100 / " + uriage + " as '売上比(%)', ";
            //sql += " floor(isnull(sum(case when 伝票日付 between '" + maey + "0401' and '" + atoy + "0331' then 金額 else 0 end), 0) / " + i2 + ") as '月平均', ";
            //sql += " floor(isnull(sum(case when 伝票日付 between '" + maeyex + "0401' and '" + atoyex + "0331' then 金額 else 0 end), 0) / 12 ) as '月平均(前期)' ";
            //sql += " from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";


            //if (checkBox1.Checked)
            //{
            //    sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '9001' and '9499'";
            //}
            //else
            //{
            //    sql += " where 伝票日付 between '" + maeyex + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') and 科目コード between '9001' and '9499'";
            //}

            //if (checkBox3.Checked)
            //{
            //    sql += " and 科目コード not in ('8214','8215','8345','8346') ";
            //}

            //sql += sql2;


            sql += " order by 科目コード ";

            dt = Com.GetDB(sql);

            //売上レコードがある場合は処理する
            if (dt.Select("科目コード = '8010'", "").Length == 0)
            {
                if (checkBox2.Checked)
                {
                    DataTable dummy = new DataTable();
                    dummy = dt.Copy();

                    dt.Clear();
                    dt.ImportRow(dummy.Select("科目コード = '8300'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '8600'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '9000'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '9990'", "")[0]);
                }
                //return;
            }
            else
            {
                //売上
                DataRow[] uriall;
                uriall = dt.Select("科目コード = '8010'", "");//売上
                decimal u02 = uriall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][2].ToString());
                decimal u03 = uriall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][3].ToString());
                decimal u04 = uriall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][4].ToString());
                decimal u05 = uriall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][5].ToString());
                decimal u06 = uriall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][6].ToString());
                decimal u07 = uriall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][7].ToString());
                decimal u08 = uriall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][8].ToString());
                decimal u09 = uriall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][9].ToString());
                decimal u10 = uriall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][10].ToString());
                decimal u11 = uriall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][11].ToString());
                decimal u12 = uriall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][12].ToString());
                decimal u13 = uriall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][13].ToString());
                decimal u14 = uriall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][14].ToString());
                decimal u15 = uriall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][15].ToString());
                decimal u16 = uriall[0][16].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][16].ToString());
                decimal u17 = uriall[0][17].ToString() == "" ? 0 : Convert.ToDecimal(uriall[0][17].ToString());

                //8300
                DataRow[] gkeihiall;
                gkeihiall = dt.Select("科目コード = '8300'", "");
                decimal gk02 = gkeihiall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][2].ToString());
                decimal gk03 = gkeihiall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][3].ToString());
                decimal gk04 = gkeihiall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][4].ToString());
                decimal gk05 = gkeihiall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][5].ToString());
                decimal gk06 = gkeihiall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][6].ToString());
                decimal gk07 = gkeihiall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][7].ToString());
                decimal gk08 = gkeihiall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][8].ToString());
                decimal gk09 = gkeihiall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][9].ToString());
                decimal gk10 = gkeihiall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][10].ToString());
                decimal gk11 = gkeihiall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][11].ToString());
                decimal gk12 = gkeihiall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][12].ToString());
                decimal gk13 = gkeihiall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][13].ToString());
                decimal gk14 = gkeihiall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][14].ToString());
                decimal gk15 = gkeihiall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][15].ToString());
                decimal gk16 = gkeihiall[0][16].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][16].ToString());
                decimal gk17 = gkeihiall[0][17].ToString() == "" ? 0 : Convert.ToDecimal(gkeihiall[0][17].ToString());

                //8600販管費合計
                DataRow[] hankanall;
                hankanall = dt.Select("科目コード = '8600'", "");
                decimal hk02 = hankanall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][2].ToString());
                decimal hk03 = hankanall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][3].ToString());
                decimal hk04 = hankanall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][4].ToString());
                decimal hk05 = hankanall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][5].ToString());
                decimal hk06 = hankanall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][6].ToString());
                decimal hk07 = hankanall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][7].ToString());
                decimal hk08 = hankanall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][8].ToString());
                decimal hk09 = hankanall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][9].ToString());
                decimal hk10 = hankanall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][10].ToString());
                decimal hk11 = hankanall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][11].ToString());
                decimal hk12 = hankanall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][12].ToString());
                decimal hk13 = hankanall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][13].ToString());
                decimal hk14 = hankanall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][14].ToString());
                decimal hk15 = hankanall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][15].ToString());
                decimal hk16 = hankanall[0][16].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][16].ToString());
                decimal hk17 = hankanall[0][17].ToString() == "" ? 0 : Convert.ToDecimal(hankanall[0][17].ToString());


                //9000営業外損益合計
                DataRow[] eigaiall;
                eigaiall = dt.Select("科目コード = '9000'", "");
                decimal ek02 = eigaiall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][2].ToString());
                decimal ek03 = eigaiall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][3].ToString());
                decimal ek04 = eigaiall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][4].ToString());
                decimal ek05 = eigaiall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][5].ToString());
                decimal ek06 = eigaiall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][6].ToString());
                decimal ek07 = eigaiall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][7].ToString());
                decimal ek08 = eigaiall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][8].ToString());
                decimal ek09 = eigaiall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][9].ToString());
                decimal ek10 = eigaiall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][10].ToString());
                decimal ek11 = eigaiall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][11].ToString());
                decimal ek12 = eigaiall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][12].ToString());
                decimal ek13 = eigaiall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][13].ToString());
                decimal ek14 = eigaiall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][14].ToString());
                decimal ek15 = eigaiall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][15].ToString());
                decimal ek16 = eigaiall[0][16].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][16].ToString());
                decimal ek17 = eigaiall[0][17].ToString() == "" ? 0 : Convert.ToDecimal(eigaiall[0][17].ToString());

                ////管理経費
                //DataRow[] kankeihiall;
                //kankeihiall = dt.Select("科目コード = '9980'", "");
                //decimal kk02 = kankeihiall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][2].ToString());
                //decimal kk03 = kankeihiall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][3].ToString());
                //decimal kk04 = kankeihiall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][4].ToString());
                //decimal kk05 = kankeihiall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][5].ToString());
                //decimal kk06 = kankeihiall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][6].ToString());
                //decimal kk07 = kankeihiall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][7].ToString());
                //decimal kk08 = kankeihiall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][8].ToString());
                //decimal kk09 = kankeihiall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][9].ToString());
                //decimal kk10 = kankeihiall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][10].ToString());
                //decimal kk11 = kankeihiall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][11].ToString());
                //decimal kk12 = kankeihiall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][12].ToString());
                //decimal kk13 = kankeihiall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][13].ToString());
                //decimal kk14 = kankeihiall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][14].ToString());
                //decimal kk15 = kankeihiall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(kankeihiall[0][15].ToString());

                //全体経費
                DataRow[] keihiall;
                keihiall = dt.Select("科目コード = '9990'", "");//全体経費
                decimal k02 = keihiall[0][2].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][2].ToString());
                decimal k03 = keihiall[0][3].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][3].ToString());
                decimal k04 = keihiall[0][4].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][4].ToString());
                decimal k05 = keihiall[0][5].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][5].ToString());
                decimal k06 = keihiall[0][6].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][6].ToString());
                decimal k07 = keihiall[0][7].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][7].ToString());
                decimal k08 = keihiall[0][8].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][8].ToString());
                decimal k09 = keihiall[0][9].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][9].ToString());
                decimal k10 = keihiall[0][10].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][10].ToString());
                decimal k11 = keihiall[0][11].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][11].ToString());
                decimal k12 = keihiall[0][12].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][12].ToString());
                decimal k13 = keihiall[0][13].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][13].ToString());
                decimal k14 = keihiall[0][14].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][14].ToString());
                decimal k15 = keihiall[0][15].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][15].ToString());
                decimal k16 = keihiall[0][16].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][16].ToString());
                decimal k17 = keihiall[0][17].ToString() == "" ? 0 : Convert.ToDecimal(keihiall[0][17].ToString());


                //合計のみ表示
                if (checkBox2.Checked)
                {
                    DataTable dummy = new DataTable();
                    dummy = dt.Copy();

                    dt.Clear();

                    dt.ImportRow(dummy.Select("科目コード = '8010'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '8300'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '8600'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '9000'", "")[0]);
                    dt.ImportRow(dummy.Select("科目コード = '9990'", "")[0]);
                }

                //売上総利益
                DataRow gdr2 = dt.NewRow();
                gdr2[0] = "8301";
                gdr2[1] = "売上総利益";
                gdr2[2] = u02 - gk02;
                gdr2[3] = u03 - gk03;
                gdr2[4] = u04 - gk04;
                gdr2[5] = u05 - gk05;
                gdr2[6] = u06 - gk06;
                gdr2[7] = u07 - gk07;
                gdr2[8] = u08 - gk08;
                gdr2[9] = u09 - gk09;
                gdr2[10] = u10 - gk10;
                gdr2[11] = u11 - gk11;
                gdr2[12] = u12 - gk12;
                gdr2[13] = u13 - gk13;
                gdr2[14] = u14 - gk14;
                gdr2[15] = u15 - gk15;
                gdr2[16] = u16 - gk16;
                gdr2[17] = u17 - gk17;
                dt.Rows.Add(gdr2);

                //原価率
                DataRow gdr = dt.NewRow();
                gdr[0] = "8302";
                gdr[1] = "原価率";
                gdr[2] = u02 == 0 ? 0 : gk02 / u02 * 100;
                gdr[3] = u03 == 0 ? 0 : gk03 / u03 * 100;
                gdr[4] = u04 == 0 ? 0 : gk04 / u04 * 100;
                gdr[5] = u05 == 0 ? 0 : gk05 / u05 * 100;
                gdr[6] = u06 == 0 ? 0 : gk06 / u06 * 100;
                gdr[7] = u07 == 0 ? 0 : gk07 / u07 * 100;
                gdr[8] = u08 == 0 ? 0 : gk08 / u08 * 100;
                gdr[9] = u09 == 0 ? 0 : gk09 / u09 * 100;
                gdr[10] = u10 == 0 ? 0 : gk10 / u10 * 100;
                gdr[11] = u11 == 0 ? 0 : gk11 / u11 * 100;
                gdr[12] = u12 == 0 ? 0 : gk12 / u12 * 100;
                gdr[13] = u13 == 0 ? 0 : gk13 / u13 * 100;
                gdr[14] = u14 == 0 ? 0 : gk14 / u14 * 100;
                gdr[15] = u15 == 0 ? 0 : gk15 / u15 * 100;
                gdr[16] = u16 == 0 ? 0 : gk16 / u16 * 100;
                gdr[17] = u17 == 0 ? 0 : gk17 / u17 * 100;
                dt.Rows.Add(gdr);

                //営業利益
                DataRow edr = dt.NewRow();
                edr[0] = "8601";
                edr[1] = "営業利益";
                edr[2] = u02 - gk02 - hk02;
                edr[3] = u03 - gk03 - hk03;
                edr[4] = u04 - gk04 - hk04;
                edr[5] = u05 - gk05 - hk05;
                edr[6] = u06 - gk06 - hk06;
                edr[7] = u07 - gk07 - hk07;
                edr[8] = u08 - gk08 - hk08;
                edr[9] = u09 - gk09 - hk09;
                edr[10] = u10 - gk10 - hk10;
                edr[11] = u11 - gk11 - hk11;
                edr[12] = u12 - gk12 - hk12;
                edr[13] = u13 - gk13 - hk13;
                edr[14] = u14 - gk14 - hk14;
                edr[15] = u15 - gk15 - hk15;
                edr[16] = u16 - gk16 - hk16;
                edr[17] = u17 - gk17 - hk17;
                dt.Rows.Add(edr);

                //経常利益
                DataRow kdr = dt.NewRow();
                kdr[0] = "9001";
                kdr[1] = "経常利益";
                kdr[2] = u02 - gk02 - hk02 - ek02;
                kdr[3] = u03 - gk03 - hk03 - ek03;
                kdr[4] = u04 - gk04 - hk04 - ek04;
                kdr[5] = u05 - gk05 - hk05 - ek05;
                kdr[6] = u06 - gk06 - hk06 - ek06;
                kdr[7] = u07 - gk07 - hk07 - ek07;
                kdr[8] = u08 - gk08 - hk08 - ek08;
                kdr[9] = u09 - gk09 - hk09 - ek09;
                kdr[10] = u10 - gk10 - hk10 - ek10;
                kdr[11] = u11 - gk11 - hk11 - ek11;
                kdr[12] = u12 - gk12 - hk12 - ek12;
                kdr[13] = u13 - gk13 - hk13 - ek13;
                kdr[14] = u14 - gk14 - hk14 - ek14;
                kdr[15] = u15 - gk15 - hk15 - ek15;
                kdr[16] = u16 - gk16 - hk16 - ek16;
                kdr[17] = u17 - gk17 - hk17 - ek17;
                dt.Rows.Add(kdr);


                //総利益
                DataRow dr2 = dt.NewRow();
                dr2[0] = "9998";
                dr2[1] = "当期利益";
                dr2[2] = u02 - k02;
                dr2[3] = u03 - k03;
                dr2[4] = u04 - k04;
                dr2[5] = u05 - k05;
                dr2[6] = u06 - k06;
                dr2[7] = u07 - k07;
                dr2[8] = u08 - k08;
                dr2[9] = u09 - k09;
                dr2[10] = u10 - k10;
                dr2[11] = u11 - k11;
                dr2[12] = u12 - k12;
                dr2[13] = u13 - k13;
                dr2[14] = u14 - k14;
                dr2[15] = u15 - k15;
                dr2[16] = u16 - k16;
                dr2[17] = u17 - k17;
                dt.Rows.Add(dr2);

                DataRow dr = dt.NewRow();
                dr[0] = "9999";
                dr[1] = "当期利益率";
                dr[2] = u02 == 0 ? 0 : k02 / u02 * 100;
                dr[3] = u03 == 0 ? 0 : k03 / u03 * 100;
                dr[4] = u04 == 0 ? 0 : k04 / u04 * 100;
                dr[5] = u05 == 0 ? 0 : k05 / u05 * 100;
                dr[6] = u06 == 0 ? 0 : k06 / u06 * 100;
                dr[7] = u07 == 0 ? 0 : k07 / u07 * 100;
                dr[8] = u08 == 0 ? 0 : k08 / u08 * 100;
                dr[9] = u09 == 0 ? 0 : k09 / u09 * 100;
                dr[10] = u10 == 0 ? 0 : k10 / u10 * 100;
                dr[11] = u11 == 0 ? 0 : k11 / u11 * 100;
                dr[12] = u12 == 0 ? 0 : k12 / u12 * 100;
                dr[13] = u13 == 0 ? 0 : k13 / u13 * 100;
                dr[14] = u14 == 0 ? 0 : k14 / u14 * 100;
                dr[15] = u15 == 0 ? 0 : k15 / u15 * 100;
                dr[16] = u16 == 0 ? 0 : k16 / u16 * 100;
                dr[17] = u17 == 0 ? 0 : k17 / u17 * 100;
                dt.Rows.Add(dr);
            }

            //ソート
            DataView dv = new DataView(dt);
            dv.Sort = "科目コード";

            //dataGridView1.DataSource = dt;
            dataGridView1.DataSource = dv.ToTable();


            // 反転対象列：04月(2)～03月(13)、年間(14)、月平均(16)、月平均(前期)(17)
            int[] flipCols = { 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17 };

            foreach (var c in flipCols)
            {
                var col = dataGridView1.Columns[c];
                col.ValueType = typeof(decimal);
                col.DefaultCellStyle.NullValue = 0m;
                col.DefaultCellStyle.Format = "#,0.##";
            }

            // 売上比(%) は比率表示だけ調整（反転しない）
            dataGridView1.Columns[15].ValueType = typeof(decimal); // or double でもOK
            dataGridView1.Columns[15].DefaultCellStyle.Format = "N2";



            //表示処理
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 120;

            for (int i = 2; i < 18; i++)
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

                dataGridView1.Columns[i].Width = 75;

                //三桁区切り表示
                dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0.##";

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }

            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.Beige;

            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.Beige;
            dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.Beige;

            dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[14].HeaderCell.Style.BackColor = Color.AntiqueWhite;

            //売上比
            dataGridView1.Columns[15].DefaultCellStyle.Format = "N2";

            Com.InHistory("03_科目別損益", "", "");

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

            DataTable dt = new DataTable();
            string sql = "";

            sql += "select 科目名, 部門名, 現場名, 金額, 摘要文, 取引先名, 伝票日付, 伝票番号, 科目コード, 部門コード, 現場コード, 職種, 担当区分, 消費税額, 税区分コード, 税区分名,工種名　from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql += " where 科目コード between '8000' and '9900' ";
            if (col == "年間")
            {

                if (checkBox1.Checked)
                {
                    sql += " and 伝票日付 between '" + maey + "0401' and '" + atoy + "0331'";
                    //sql += " where 伝票日付 between '" + maeyex + "0401' and '" + atoy + "0331' and 科目コード between '8000' and '9900'";
                }
                else
                {
                    sql += " and 伝票日付 between '" + maey + "0401' and (select max(伝票日付) from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 where 科目コード = '8010') ";
                }
            }
            else
            {
                if (col == "01月" || col == "02月" || col == "03月")
                {
                    sql += " and 伝票日付 like '" + atoy + col.Replace("月", "") + "%'";
                }
                else
                {
                    sql += " and 伝票日付 like '" + maey + col.Replace("月", "") + "%'";
                }
            }

            if (row == "8300") //原価合計
            {
                sql += " and 科目コード between '8212' and '8298'";
            }
            else if (row == "8600") //販管費合計
            {
                sql += " and 科目コード between '8310' and '8599'";
            }
            else if (row == "9000") //営業外損益合計
            {
                sql += " and 科目コード between '8600' and '8999'";
            }
            else if (row == "9990") //経費総計
            {
                sql += " and 科目コード between '8200' and '9900'"; 
            }
            else
            {
                sql += " and 科目コード = '" + row + "'";
            }

            sql += GetTSG();


            //sql += " order by 科目コード, 部門コード, 現場コード,金額";
            sql += " order by 金額 desc";

            dt = Com.GetDB(sql);



            dataGridView2.DataSource = dt;

            //売上
            if (row.Substring(0, 1) == "0")
            {
                //MessageBox.Show("売上！");
                //TODO 
            }
            else
            {
                dataGridView2.Columns[0].Width = 120;//科目名
                dataGridView2.Columns[1].Width = 120;//部門名
                dataGridView2.Columns[2].Width = 250;//現場名
                dataGridView2.Columns[3].Width = 70;//金額
                dataGridView2.Columns[4].Width = 400;//摘要
                dataGridView2.Columns[5].Width = 250;//取引先
                dataGridView2.Columns[6].Width = 60;//
                dataGridView2.Columns[7].Width = 60;//
                dataGridView2.Columns[8].Width = 60;//
                dataGridView2.Columns[9].Width = 60;//
                dataGridView2.Columns[10].Width = 60;//
                dataGridView2.Columns[11].Width = 60;//

                //金額右寄
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //三桁区切り表示
                dataGridView2.Columns[3].DefaultCellStyle.Format = "#,0";
                //dataGridView2.Columns[4].DefaultCellStyle.Format = "#,0";
            }
        }

        //表示期間
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //期間指定
            maey = comboBox1.SelectedItem.ToString().Substring(4, 4);
            atoy = comboBox1.SelectedItem.ToString().Substring(9, 4);

            maeyex = (Convert.ToInt16(maey) - 1).ToString();
            atoyex = (Convert.ToInt16(atoy) - 1).ToString();

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

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //科目列
            if (e.ColumnIndex == 0)
            {
                if (e.Value.ToString() == "8010")　//売上
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PaleGreen;
                }
                else if (e.Value.ToString() == "8300" || e.Value.ToString() == "8600" || e.Value.ToString() == "9000" || e.Value.ToString() == "9990") //経費
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Khaki;
                }
                else if (e.Value.ToString() == "8301" || e.Value.ToString() == "8601" || e.Value.ToString() == "9001" || e.Value.ToString() == "9998") //利益
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PaleTurquoise;
                }
                else if (e.Value.ToString() == "8302" || e.Value.ToString() == "9999") //率
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.OldLace;
                    //dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Wheat;
                }
                //else if (e.Value.ToString() == "9990")
                //{
                //    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Wheat;
                //}

                //dgvtouzisseki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
                //dgvtouzisseki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
                //dgvtouzisseki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            }

            //セルの列を確認
            decimal val = 0;
            if (e.Value != null && e.ColumnIndex != 0 && e.ColumnIndex < 14 && decimal.TryParse(e.Value.ToString(), out val))
            {
                if (val == 0) return;

                decimal cell = Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                decimal av = Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells[16].Value);

                if (av < 0)
                {
                    if (cell < av * Convert.ToDecimal(red.Value / 100))
                    {
                        //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                        e.CellStyle.ForeColor = Color.Blue;
                    }
                    else if (cell > av * Convert.ToDecimal(blue.Value / 100))
                    {
                        e.CellStyle.ForeColor = Color.Red;
                    }
                }
                else
                {
                    if (cell < av * Convert.ToDecimal(blue.Value / 100))
                    {
                        //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                        e.CellStyle.ForeColor = Color.Blue;
                    }
                    else if (cell > av * Convert.ToDecimal(red.Value / 100))
                    {
                        e.CellStyle.ForeColor = Color.Red;
                    }
 
            }
        }
    }

    private void label11_Click(object sender, EventArgs e)
    {

    }

    private void red_ValueChanged(object sender, EventArgs e)
    {
        GetData();
    }

    private void blue_ValueChanged(object sender, EventArgs e)
    {
        GetData();
    }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // 反転対象列：04月(2)～03月(13)、年間(14)、月平均(16)、月平均(前期)(17)
            bool targetCol =
                (e.ColumnIndex >= 2 && e.ColumnIndex <= 13) ||
                e.ColumnIndex == 14 || e.ColumnIndex == 16 || e.ColumnIndex == 17;

            if (!targetCol) return;

            var dgv = (DataGridView)sender;
            if (e.RowIndex < 0) return;

            // 反転対象：8610～8990（8810・8990は除外）＋ 9000
            var codeObj = dgv.Rows[e.RowIndex].Cells[0].Value;
            if (codeObj == null) return;
            if (!int.TryParse(codeObj.ToString(), out var code)) return;
            // 反転対象：8610～8990（8810・8990は除外）＋ 9000, 9110, 9111, 9120, 9180
            bool targetRow = ((code >= 8610 && code <= 8990 && code != 8810 && code != 8990)
                              || code == 8296 || code == 9000 || code == 9110 || code == 9111 || code == 9120 || code == 9180);

            if (!targetRow) return;

            // 生値
            var raw = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (raw == null || raw == DBNull.Value) return;

            decimal val;
            try
            {
                if (raw is decimal d) val = d;
                else if (raw is double db) val = Convert.ToDecimal(db);
                else if (raw is float f) val = Convert.ToDecimal(f);
                else if (raw is long l) val = l;
                else if (raw is int i) val = i;
                else if (raw is short s) val = s;
                else if (raw is string str)
                {
                    if (!decimal.TryParse(str, out val)) return;
                }
                else
                {
                    val = Convert.ToDecimal(raw);
                }
            }
            catch { return; }

            // 列の ValueType は上で decimal に固定済みなので decimal で返す
            e.Value = -val;

            // 値だけ渡して、書式(#,0.##)は DataGridView 側に任せる
            e.FormattingApplied = false;
        }





        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // フォーマットエラーを無視（落ちないようにする）
            e.ThrowException = false;
        }


    }
}
