using Npgsql;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KamokuYosan : Form
    {
        //YM
        private string maey = "";
        private string atoy = "";

        private string maeyex = "";
        private string atoyex = "";

        //科目コードの幅
        //private string kamos = "8200";
        //private string kamoe = "9900";

        public KamokuYosan()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("53期(2024～2025)");
            //comboBox1.Items.Add("52期(2023～2024)");
            //comboBox1.Items.Add("51期(2022～2023)");

            comboBox1.SelectedIndex = 0;

            SetTiku();
            SetBumon();
            SetGenba();

            //フォントサイズの変更
            //dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            cbgoukei.Checked = true;
            //GetData();

            //dataGridView1でセル、行、列が複数選択されないようにする
            //dataGridView1.MultiSelect = false;
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            //TODO 次がうめこみ！！
            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.Yosankanri where 年月 between '" + maey + "04' and '" + atoy + "03' and 次 = '1' order by 担当区分 ";
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

            //TODO 次がうめこみ！！
            DataTable dt = new DataTable();
            string sql = "select distinct 職種 from dbo.Yosankanri where 年月 between '" + maey + "04' and '" + atoy + "03' and 次 = '1'";

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

            //TODO 次がうめこみ！！
            //売上
            string sql = "select distinct 現場CD,現場名 from dbo.Yosankanri　where 年月 between '" + maey + "04' and '" + atoy + "03' and 次 = '1'";

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += " order by 現場CD,現場名 ";


            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["現場CD"].ToString() + ' ' + row["現場名"].ToString());
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

            if (itemcount > 0 && ckcount == 0 )
            {
                sql = " and a.現場CD = '' ";
            }
            else
            if (itemcount / 2 > ckcount)
            {
                string sql3 = "";
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (checkedListBox3.GetItemChecked(i))
                    {
                        if (checkedListBox3.Items[i].ToString().Length < 5)
                        {
                            sql3 += " or isnull(a.現場CD,'') = '' ";
                        }
                        else
                        {
                            sql3 += " or isnull(a.現場CD,'') = '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
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
                            sql += " and isnull(a.現場CD,'') <> '' ";
                        }
                        else
                        {
                            sql += " and isnull(a.現場CD,'') <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
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

            //次うめこみ！！
            //string sql = "select 年月,sum(固定売上) as 固定売上,sum(臨時売上) as 臨時売上, sum(人件費) as 人件費,sum(賞与) as 賞与,sum(諸経費) as 諸経費,sum(管理人件費) as 管理人件費,sum(管理賞与) as 管理賞与,sum(管理諸経費) as 管理諸経費 ";
            string sql = "";

            if (cbsyouyo.Checked)
            {
                //賞与無

                //合算
                sql += "select 年月, ";
                if (!cbgoukei.Checked) sql += "sum(予算_固定売上) as 予算_固定売上, sum(予算_臨時売上) as 予算_臨時売上, ";
                sql += "sum(予算_売上) as 予算_売上, ";
                if (!cbgoukei.Checked) sql += "sum(予算_人件費) as 予算_人件費,sum(予算_諸経費) as 予算_諸経費,  ";
                sql += "sum(予算_経費) as 予算_経費,  sum(予算_現場利益) as 予算_現場利益, sum(予算_現場計数) as 予算_現場計数,   ";
                if (!cbgoukei.Checked) sql += "sum(予算_管理人件費) as 予算_管理人件費,sum(予算_管理諸経費) as 予算_管理諸経費, ";
                sql += "sum(予算_管理経費) as 予算_管理経費, sum(予算_管理計数) as 予算_管理計数, sum(予算_部門利益) as 予算_部門利益, sum(予算_部門計数) as 予算_部門計数,  ";

                if (!cbgoukei.Checked) sql += "sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, ";
                sql += "sum(売上) as 売上, ";
                if (!cbgoukei.Checked) sql += "sum(人件費) as 人件費,sum(諸経費) as 諸経費,  ";
                sql += "sum(経費) as 経費,  sum(現場利益) as 現場利益, sum(現場計数) as 現場計数,   ";
                if (!cbgoukei.Checked) sql += "sum(管理人件費) as 管理人件費,sum(管理諸経費) as 管理諸経費, ";
                sql += "sum(管理経費) as 管理経費,sum(管理計数) as 管理計数, sum(部門利益) as 部門利益, sum(部門計数) as 部門計数,  ";

                if (!cbgoukei.Checked) sql += "sum(前期_固定売上) as 前期_固定売上, sum(前期_臨時売上) as 前期_臨時売上, ";
                sql += "sum(前期_売上) as 前期_売上,";
                if (!cbgoukei.Checked) sql += "sum(前期_人件費) as 前期_人件費,sum(前期_諸経費) as 前期_諸経費,  ";
                sql += "sum(前期_経費) as 前期_経費,  sum(前期_現場利益) as 前期_現場利益, sum(前期_現場計数) as 前期_現場計数, ";
                if (!cbgoukei.Checked) sql += "sum(前期_管理人件費) as 前期_管理人件費,sum(前期_管理諸経費) as 前期_管理諸経費, ";
                sql += "sum(前期_管理経費) as 前期_管理経費, sum(前期_管理計数) as 前期_管理計数, sum(前期_部門利益) as 前期_部門利益, sum(前期_部門計数) as 前期_部門計数 ";
                sql += "from(select 年月, ";

                //予算
                sql += "sum(固定売上) as 予算_固定売上, sum(臨時売上) as 予算_臨時売上, sum(固定売上) + sum(臨時売上) as 予算_売上, ";
                sql += "sum(人件費) as 予算_人件費, sum(諸経費) as 予算_諸経費, ";
                sql += "sum(人件費) + sum(諸経費) as 予算_経費, sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) as 予算_現場利益, ";
                sql += "case when sum(固定売上) + sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_現場計数,   ";
                sql += "sum(管理人件費) as 予算_管理人件費,sum(管理諸経費) as 予算_管理諸経費, sum(管理人件費) + sum(管理諸経費) as 予算_管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) - sum(管理人件費) - sum(管理諸経費) as 予算_部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費) + sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_部門計数,  ";


                sql += "0 as 固定売上, 0 as 臨時売上, 0 as 売上, 0 as 人件費,0 as 諸経費, 0 as 経費,  0 as 現場利益, 0 as 現場計数,  ";
                sql += "0 as 管理人件費,0 as 管理諸経費, 0 as 管理経費, 0 as 管理計数, 0 as 部門利益, 0 as 部門計数,  ";

                sql += "0 as 前期_固定売上, 0 as 前期_臨時売上, 0 as 前期_売上, 0 as 前期_人件費,0 as 前期_諸経費, 0 as 前期_経費, 0 as 前期_現場利益, 0 as 前期_現場計数,   ";
                sql += "0 as 前期_管理人件費,0 as 前期_管理諸経費, 0 as 前期_管理経費, 0 as 前期_管理計数, 0 as 前期_部門利益, 0 as 前期_部門計数 ";
                sql += "from dbo.yosankanri a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 次 = '1' ";
                sql += GetTSG();
                sql += " group by 年月 ";

                sql += "union all ";

                //実績
                sql += "select 年月, ";
                sql += "0 as 予算_固定売上, 0 as 予算_臨時売上, 0 as 予算_売上, 0 as 予算_人件費,0 as 予算_諸経費, 0 as 予算_経費,  0 as 予算_現場利益, 0 as 予算_現場計数,   ";
                sql += "0 as 予算_管理人件費,0 as 予算_管理諸経費, 0 as 予算_管理経費, 0 as 予算_管理計数, 0 as 予算_部門利益, 0 as 予算_部門計数,  ";

                sql += "sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(固定売上) + sum(臨時売上) as 売上,   ";
                sql += "sum(人件費) as 人件費,sum(諸経費) as 諸経費,  ";
                sql += "sum(人件費) + sum(諸経費) as 経費,  sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) as 現場利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 現場計数,   ";
                sql += "sum(管理人件費) as 管理人件費,sum(管理諸経費) as 管理諸経費, sum(管理人件費) + sum(管理諸経費) as 管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) - sum(管理人件費) - sum(管理諸経費) as 部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費) + sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 部門計数, ";


                sql += "0 as 前期_固定売上, 0 as 前期_臨時売上, 0 as 前期_売上, 0 as 前期_人件費,0 as 前期_諸経費, 0 as 前期_経費, 0 as 前期_現場利益, 0 as 前期_現場計数,   ";
                sql += "0 as 前期_管理人件費,0 as 前期_管理諸経費, 0 as 前期_管理経費, 0 as 前期_管理計数, 0 as 前期_部門利益, 0 as 前期_部門計数 ";
                sql += "from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '202404' and '202503' ";
                sql += GetTSG();
                sql += " group by 年月 ";

                sql += "union all ";

                //前期
                sql += "select 年月 + 100, ";
                sql += "0 as 予算_固定売上, 0 as 予算_臨時売上, 0 as 予算_売上, 0 as 予算_人件費,0 as 予算_諸経費, 0 as 予算_経費,  0 as 予算_現場利益, 0 as 予算_現場計数,   ";
                sql += "0 as 予算_管理人件費,0 as 予算_管理諸経費, 0 as 予算_管理経費, 0 as 予算_管理計数, 0 as 予算_部門利益, 0 as 予算_部門計数,  ";

                sql += "0 as 固定売上, 0 as 臨時売上, 0 as 売上, 0 as 人件費,0 as 諸経費, 0 as 経費,  0 as 現場利益, 0 as 現場計数,  ";
                sql += "0 as 管理人件費,0 as 管理諸経費, 0 as 管理経費, 0 as 管理計数, 0 as 部門利益, 0 as 部門計数,   ";

                sql += "sum(固定売上) as 前期_固定売上, sum(臨時売上) as 前期_臨時売上, sum(固定売上) + sum(臨時売上) as 前期_売上,   ";
                sql += "sum(人件費) as 前期_人件費,sum(諸経費) as 前期_諸経費,  ";
                sql += "sum(人件費) + sum(諸経費) as 前期_経費,  sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) as 前期_現場利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_現場計数,   ";
                sql += "sum(管理人件費) as 前期_管理人件費,sum(管理諸経費) as 前期_管理諸経費, sum(管理人件費) + sum(管理諸経費) as 前期_管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(諸経費) - sum(管理人件費) - sum(管理諸経費) as 前期_部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(諸経費) + sum(管理人件費) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_部門計数 ";
                sql += "from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '202304' and '202403' ";
                sql += GetTSG();
                sql += " group by 年月 ";
                sql += ") temp group by 年月 ";
            }
            else
            {
                //合算
                sql += "select 年月, ";
                if (!cbgoukei.Checked) sql += "sum(予算_固定売上) as 予算_固定売上, sum(予算_臨時売上) as 予算_臨時売上, ";
                sql += "sum(予算_売上) as 予算_売上, ";
                if (!cbgoukei.Checked) sql += "sum(予算_人件費) as 予算_人件費, sum(予算_賞与) as 予算_賞与,sum(予算_諸経費) as 予算_諸経費,  ";
                sql += "sum(予算_経費) as 予算_経費,  sum(予算_現場利益) as 予算_現場利益, sum(予算_現場計数) as 予算_現場計数,   ";
                if (!cbgoukei.Checked) sql += "sum(予算_管理人件費) as 予算_管理人件費,sum(予算_管理賞与) as 予算_管理賞与,sum(予算_管理諸経費) as 予算_管理諸経費, ";
                sql += "sum(予算_管理経費) as 予算_管理経費, sum(予算_管理計数) as 予算_管理計数, sum(予算_部門利益) as 予算_部門利益, sum(予算_部門計数) as 予算_部門計数,  ";

                if (!cbgoukei.Checked) sql += "sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, ";
                sql += "sum(売上) as 売上, ";
                if (!cbgoukei.Checked) sql += "sum(人件費) as 人件費, sum(賞与) as 賞与,sum(諸経費) as 諸経費,  ";
                sql += "sum(経費) as 経費,  sum(現場利益) as 現場利益, sum(現場計数) as 現場計数,   ";
                if (!cbgoukei.Checked) sql += "sum(管理人件費) as 管理人件費,sum(管理賞与) as 管理賞与,sum(管理諸経費) as 管理諸経費, ";
                sql += "sum(管理経費) as 管理経費,sum(管理計数) as 管理計数, sum(部門利益) as 部門利益, sum(部門計数) as 部門計数,  ";

                if (!cbgoukei.Checked) sql += "sum(前期_固定売上) as 前期_固定売上, sum(前期_臨時売上) as 前期_臨時売上, ";
                sql += "sum(前期_売上) as 前期_売上,";
                if (!cbgoukei.Checked) sql += "sum(前期_人件費) as 前期_人件費, sum(前期_賞与) as 前期_賞与,sum(前期_諸経費) as 前期_諸経費,  ";
                sql += "sum(前期_経費) as 前期_経費,  sum(前期_現場利益) as 前期_現場利益, sum(前期_現場計数) as 前期_現場計数, ";
                if (!cbgoukei.Checked) sql += "sum(前期_管理人件費) as 前期_管理人件費,sum(前期_管理賞与) as 前期_管理賞与,sum(前期_管理諸経費) as 前期_管理諸経費, ";
                sql += "sum(前期_管理経費) as 前期_管理経費, sum(前期_管理計数) as 前期_管理計数, sum(前期_部門利益) as 前期_部門利益, sum(前期_部門計数) as 前期_部門計数 ";
                sql += "from(select 年月, ";

                //予算
                sql += "sum(固定売上) as 予算_固定売上, sum(臨時売上) as 予算_臨時売上, sum(固定売上) + sum(臨時売上) as 予算_売上, ";
                sql += "sum(人件費) as 予算_人件費, sum(賞与) as 予算_賞与, sum(諸経費) as 予算_諸経費, ";
                sql += "sum(人件費) + sum(賞与) + sum(諸経費) as 予算_経費, sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) as 予算_現場利益, ";
                sql += "case when sum(固定売上) + sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_現場計数,   ";
                sql += "sum(管理人件費) as 予算_管理人件費,sum(管理賞与) as 予算_管理賞与,sum(管理諸経費) as 予算_管理諸経費, sum(管理人件費) + sum(管理賞与) + sum(管理諸経費) as 予算_管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) - sum(管理人件費) - sum(管理賞与) - sum(管理諸経費) as 予算_部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費) + sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 予算_部門計数,  ";


                sql += "0 as 固定売上, 0 as 臨時売上, 0 as 売上, 0 as 人件費, 0 as 賞与,0 as 諸経費, 0 as 経費,  0 as 現場利益, 0 as 現場計数,  ";
                sql += "0 as 管理人件費,0 as 管理賞与,0 as 管理諸経費, 0 as 管理経費, 0 as 管理計数, 0 as 部門利益, 0 as 部門計数,  ";

                sql += "0 as 前期_固定売上, 0 as 前期_臨時売上, 0 as 前期_売上, 0 as 前期_人件費, 0 as 前期_賞与,0 as 前期_諸経費, 0 as 前期_経費, 0 as 前期_現場利益, 0 as 前期_現場計数,   ";
                sql += "0 as 前期_管理人件費,0 as 前期_管理賞与,0 as 前期_管理諸経費, 0 as 前期_管理経費, 0 as 前期_管理計数, 0 as 前期_部門利益, 0 as 前期_部門計数 ";
                sql += "from dbo.yosankanri a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 次 = '1' " ;
                sql += GetTSG();
                sql += " group by 年月 ";

                sql += "union all ";

                //実績
                sql += "select 年月, ";
                sql += "0 as 予算_固定売上, 0 as 予算_臨時売上, 0 as 予算_売上, 0 as 予算_人件費, 0 as 予算_賞与,0 as 予算_諸経費, 0 as 予算_経費,  0 as 予算_現場利益, 0 as 予算_現場計数,   ";
                sql += "0 as 予算_管理人件費,0 as 予算_管理賞与,0 as 予算_管理諸経費, 0 as 予算_管理経費, 0 as 予算_管理計数, 0 as 予算_部門利益, 0 as 予算_部門計数,  ";

                sql += "sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(固定売上) + sum(臨時売上) as 売上,   ";
                sql += "sum(人件費) as 人件費, sum(賞与) as 賞与,sum(諸経費) as 諸経費,  ";
                sql += "sum(人件費) + sum(賞与) + sum(諸経費) as 経費,  sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) as 現場利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 現場計数,   ";
                sql += "sum(管理人件費) as 管理人件費,sum(管理賞与) as 管理賞与,sum(管理諸経費) as 管理諸経費, sum(管理人件費) + sum(管理賞与) + sum(管理諸経費) as 管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) - sum(管理人件費) - sum(管理賞与) - sum(管理諸経費) as 部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費) + sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 部門計数, ";


                sql += "0 as 前期_固定売上, 0 as 前期_臨時売上, 0 as 前期_売上, 0 as 前期_人件費, 0 as 前期_賞与,0 as 前期_諸経費, 0 as 前期_経費, 0 as 前期_現場利益, 0 as 前期_現場計数,   ";
                sql += "0 as 前期_管理人件費,0 as 前期_管理賞与,0 as 前期_管理諸経費, 0 as 前期_管理経費, 0 as 前期_管理計数, 0 as 前期_部門利益, 0 as 前期_部門計数 ";
                sql += "from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '202404' and '202503' ";
                sql += GetTSG();
                sql += " group by 年月 ";

                sql += "union all ";

                //前期
                sql += "select 年月 + 100, ";
                sql += "0 as 予算_固定売上, 0 as 予算_臨時売上, 0 as 予算_売上, 0 as 予算_人件費, 0 as 予算_賞与,0 as 予算_諸経費, 0 as 予算_経費,  0 as 予算_現場利益, 0 as 予算_現場計数,   ";
                sql += "0 as 予算_管理人件費,0 as 予算_管理賞与,0 as 予算_管理諸経費, 0 as 予算_管理経費, 0 as 予算_管理計数, 0 as 予算_部門利益, 0 as 予算_部門計数,  ";

                sql += "0 as 固定売上, 0 as 臨時売上, 0 as 売上, 0 as 人件費, 0 as 賞与,0 as 諸経費, 0 as 経費,  0 as 現場利益, 0 as 現場計数,  ";
                sql += "0 as 管理人件費,0 as 管理賞与,0 as 管理諸経費, 0 as 管理経費, 0 as 管理計数, 0 as 部門利益, 0 as 部門計数,   ";

                sql += "sum(固定売上) as 前期_固定売上, sum(臨時売上) as 前期_臨時売上, sum(固定売上) + sum(臨時売上) as 前期_売上,   ";
                sql += "sum(人件費) as 前期_人件費, sum(賞与) as 前期_賞与,sum(諸経費) as 前期_諸経費,  ";
                sql += "sum(人件費) + sum(賞与) + sum(諸経費) as 前期_経費,  sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) as 前期_現場利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_現場計数,   ";
                sql += "sum(管理人件費) as 前期_管理人件費,sum(管理賞与) as 前期_管理賞与,sum(管理諸経費) as 前期_管理諸経費, sum(管理人件費) + sum(管理賞与) + sum(管理諸経費) as 前期_管理経費,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_管理計数,   ";
                sql += "sum(固定売上) + sum(臨時売上) - sum(人件費) - sum(賞与) - sum(諸経費) - sum(管理人件費) - sum(管理賞与) - sum(管理諸経費) as 前期_部門利益,   ";
                sql += "case when sum(固定売上) +sum(臨時売上) = 0 then 0 else (sum(人件費) + sum(賞与) + sum(諸経費) + sum(管理人件費) + sum(管理賞与) + sum(管理諸経費)) / (sum(固定売上) + sum(臨時売上)) * 100 end as 前期_部門計数 ";
                sql += "from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '202304' and '202403' ";
                sql += GetTSG();
                sql += " group by 年月 ";
                sql += ") temp group by 年月 ";
            }


            DataTable dtnew = new DataTable();
            dtnew = Com.GetDB(sql);

            DataTable gendt = new DataTable();
            gendt = MonthDataTable(dtnew);

            //string test = gendt.Rows[0]["現場計数"].ToString();
            string test2 = gendt.Rows[0]["年間"].ToString();

            //合計表示チェック
            if (cbgoukei.Checked)
            {
                //合計のみ表示

                //年間の計数
                if (Convert.ToDecimal(gendt.Rows[0]["年間"]) > 0)
                {
                    gendt.Rows[3]["年間"] = gendt.Rows[3]["月平均"] = Convert.ToDecimal(gendt.Rows[1]["年間"]) / Convert.ToDecimal(gendt.Rows[0]["年間"]) * 100;
                    gendt.Rows[5]["年間"] = gendt.Rows[5]["月平均"] = Convert.ToDecimal(gendt.Rows[4]["年間"]) / Convert.ToDecimal(gendt.Rows[0]["年間"]) * 100;
                    gendt.Rows[7]["年間"] = gendt.Rows[7]["月平均"] = (Convert.ToDecimal(gendt.Rows[1]["年間"]) + Convert.ToDecimal(gendt.Rows[4]["年間"])) / Convert.ToDecimal(gendt.Rows[0]["年間"]) * 100;
                }

                if (Convert.ToDecimal(gendt.Rows[8]["年間"]) > 0)
                {
                    gendt.Rows[11]["年間"] = gendt.Rows[11]["月平均"] = Convert.ToDecimal(gendt.Rows[9]["年間"]) / Convert.ToDecimal(gendt.Rows[8]["年間"]) * 100;
                    gendt.Rows[13]["年間"] = gendt.Rows[13]["月平均"] = Convert.ToDecimal(gendt.Rows[12]["年間"]) / Convert.ToDecimal(gendt.Rows[8]["年間"]) * 100;
                    gendt.Rows[15]["年間"] = gendt.Rows[15]["月平均"] = (Convert.ToDecimal(gendt.Rows[9]["年間"]) + Convert.ToDecimal(gendt.Rows[12]["年間"])) / Convert.ToDecimal(gendt.Rows[8]["年間"]) * 100;
                }

                if (Convert.ToDecimal(gendt.Rows[16]["年間"]) > 0)
                {
                    gendt.Rows[19]["年間"] = gendt.Rows[19]["月平均"] = Convert.ToDecimal(gendt.Rows[17]["年間"]) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                    gendt.Rows[21]["年間"] = gendt.Rows[21]["月平均"] = Convert.ToDecimal(gendt.Rows[20]["年間"]) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                    gendt.Rows[23]["年間"] = gendt.Rows[23]["月平均"] = (Convert.ToDecimal(gendt.Rows[17]["年間"]) + Convert.ToDecimal(gendt.Rows[20]["年間"])) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                }

            }
            else
            {

                //賞与チェック
                if (cbsyouyo.Checked)
                {
                    //賞与除く


                    if (Convert.ToDecimal(gendt.Rows[2]["年間"]) > 0)
                    {
                        gendt.Rows[7]["年間"] = gendt.Rows[7]["月平均"] = Convert.ToDecimal(gendt.Rows[5]["年間"]) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                        gendt.Rows[11]["年間"] = gendt.Rows[11]["月平均"] = Convert.ToDecimal(gendt.Rows[10]["年間"]) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                        gendt.Rows[13]["年間"] = gendt.Rows[13]["月平均"] = (Convert.ToDecimal(gendt.Rows[5]["年間"]) + Convert.ToDecimal(gendt.Rows[10]["年間"])) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                    }

                    if (Convert.ToDecimal(gendt.Rows[16]["年間"]) > 0)
                    {
                        gendt.Rows[21]["年間"] = gendt.Rows[21]["月平均"] = Convert.ToDecimal(gendt.Rows[19]["年間"]) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                        gendt.Rows[25]["年間"] = gendt.Rows[25]["月平均"] = Convert.ToDecimal(gendt.Rows[24]["年間"]) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                        gendt.Rows[27]["年間"] = gendt.Rows[27]["月平均"] = (Convert.ToDecimal(gendt.Rows[19]["年間"]) + Convert.ToDecimal(gendt.Rows[24]["年間"])) / Convert.ToDecimal(gendt.Rows[16]["年間"]) * 100;
                    }

                    if (Convert.ToDecimal(gendt.Rows[30]["年間"]) > 0)
                    {
                        gendt.Rows[35]["年間"] = gendt.Rows[35]["月平均"] = Convert.ToDecimal(gendt.Rows[33]["年間"]) / Convert.ToDecimal(gendt.Rows[30]["年間"]) * 100;
                        gendt.Rows[39]["年間"] = gendt.Rows[39]["月平均"] = Convert.ToDecimal(gendt.Rows[38]["年間"]) / Convert.ToDecimal(gendt.Rows[30]["年間"]) * 100;
                        gendt.Rows[41]["年間"] = gendt.Rows[41]["月平均"] = (Convert.ToDecimal(gendt.Rows[33]["年間"]) + Convert.ToDecimal(gendt.Rows[38]["年間"])) / Convert.ToDecimal(gendt.Rows[30]["年間"]) * 100;
                    }

                }
                else
                {
                    //賞与含む

                    if (Convert.ToDecimal(gendt.Rows[2]["年間"]) > 0)
                    {
                        gendt.Rows[8]["年間"] = gendt.Rows[8]["月平均"] = Convert.ToDecimal(gendt.Rows[6]["年間"]) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                        gendt.Rows[13]["年間"] = gendt.Rows[13]["月平均"] = Convert.ToDecimal(gendt.Rows[12]["年間"]) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                        gendt.Rows[15]["年間"] = gendt.Rows[15]["月平均"] = (Convert.ToDecimal(gendt.Rows[6]["年間"]) + Convert.ToDecimal(gendt.Rows[12]["年間"])) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;
                    }

                    if (Convert.ToDecimal(gendt.Rows[18]["年間"]) > 0)
                    {
                        gendt.Rows[24]["年間"] = gendt.Rows[24]["月平均"] = Convert.ToDecimal(gendt.Rows[22]["年間"]) / Convert.ToDecimal(gendt.Rows[18]["年間"]) * 100;
                        gendt.Rows[29]["年間"] = gendt.Rows[29]["月平均"] = Convert.ToDecimal(gendt.Rows[28]["年間"]) / Convert.ToDecimal(gendt.Rows[18]["年間"]) * 100;
                        gendt.Rows[31]["年間"] = gendt.Rows[31]["月平均"] = (Convert.ToDecimal(gendt.Rows[22]["年間"]) + Convert.ToDecimal(gendt.Rows[28]["年間"])) / Convert.ToDecimal(gendt.Rows[18]["年間"]) * 100;
                    }

                    if (Convert.ToDecimal(gendt.Rows[34]["年間"]) > 0)
                    {
                        gendt.Rows[40]["年間"] = gendt.Rows[40]["月平均"] = Convert.ToDecimal(gendt.Rows[38]["年間"]) / Convert.ToDecimal(gendt.Rows[34]["年間"]) * 100;
                        gendt.Rows[45]["年間"] = gendt.Rows[45]["月平均"] = Convert.ToDecimal(gendt.Rows[44]["年間"]) / Convert.ToDecimal(gendt.Rows[34]["年間"]) * 100;
                        gendt.Rows[47]["年間"] = gendt.Rows[47]["月平均"] = (Convert.ToDecimal(gendt.Rows[38]["年間"]) + Convert.ToDecimal(gendt.Rows[44]["年間"])) / Convert.ToDecimal(gendt.Rows[34]["年間"]) * 100;
                    }
                }

            }

            //TODO 
            //前期月平均列を追加
            //gendt.Columns.Add("前期月平均");

            //gendt.Rows[8]["前期月平均"] = 0;

            dataGridView1.DataSource = gendt;



            //表示処理
            dataGridView1.Columns[0].Width = 100;

            for (int i = 1; i < 15; i++)
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
            //dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.Beige;

            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.Beige;
            //dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.Beige;

            //dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            //dataGridView1.Columns[14].HeaderCell.Style.BackColor = Color.AntiqueWhite;


            dataGridView1.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dataGridView1.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dataGridView1.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益


            Com.InHistory("予算縦横", "", "");

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

            sql += "select 科目名, 部門名, 現場名, 金額, 摘要文, 取引先名, 伝票日付, 伝票番号, 科目コード, 部門コード, 現場コード, 担当事務, 担当区分　from dbo.PCA会計仕訳データ_貸借区分_科目別損益用 ";
            sql += " where 科目コード between '8000' and '9900' ";
            if (col == "年間")
            {
                sql += " and 伝票日付 between '" + maey + "0401' and '" + atoy + "0331'";
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

            if (row == "8299") //現場経費
            {
                sql += " and 科目コード between '8200' and '8298'";
            }
            else if (row == "9980") //管理経費
            {
                sql += " and 科目コード between '8300' and '8999'";
            }
            else if (row == "9990") //全体経費
            {
                sql += " and 科目コード between '8200' and '8999'";
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

        private DataTable MonthDataTable(DataTable dt)
        {

            DataTable wkdt = new DataTable();
            wkdt = Com.replaceDataTable(dt);

            DataTable gendt = new DataTable();
            gendt.Columns.Add("項目", typeof(string));
            gendt.Columns.Add("４月", typeof(decimal));
            gendt.Columns.Add("５月", typeof(decimal));
            gendt.Columns.Add("６月", typeof(decimal));
            gendt.Columns.Add("７月", typeof(decimal));
            gendt.Columns.Add("８月", typeof(decimal));
            gendt.Columns.Add("９月", typeof(decimal));
            gendt.Columns.Add("１０月", typeof(decimal));
            gendt.Columns.Add("１１月", typeof(decimal));
            gendt.Columns.Add("１２月", typeof(decimal));
            gendt.Columns.Add("１月", typeof(decimal));
            gendt.Columns.Add("２月", typeof(decimal));
            gendt.Columns.Add("３月", typeof(decimal));
            gendt.Columns.Add("年間", typeof(decimal));
            gendt.Columns.Add("月平均", typeof(decimal));

            DataTable zendt = gendt.Clone();
            DataTable zenzendt = gendt.Clone();

            foreach (DataRow row in wkdt.Rows)
            {
                //年間
                decimal sum = 0;

                //列名の対応と、合算列の対応と、
                DataRow nr = gendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains("202404")) { nr["４月"] = row["202404"]; sum += Convert.ToDecimal(row["202404"]); }
                if (row.Table.Columns.Contains("202405")) { nr["５月"] = row["202405"]; sum += Convert.ToDecimal(row["202405"]); }
                if (row.Table.Columns.Contains("202406")) { nr["６月"] = row["202406"]; sum += Convert.ToDecimal(row["202406"]); }
                if (row.Table.Columns.Contains("202407")) { nr["７月"] = row["202407"]; sum += Convert.ToDecimal(row["202407"]); }
                if (row.Table.Columns.Contains("202408")) { nr["８月"] = row["202408"]; sum += Convert.ToDecimal(row["202408"]); }
                if (row.Table.Columns.Contains("202409")) { nr["９月"] = row["202409"]; sum += Convert.ToDecimal(row["202409"]); }
                if (row.Table.Columns.Contains("202410")) { nr["１０月"] = row["202410"]; sum += Convert.ToDecimal(row["202410"]); }
                if (row.Table.Columns.Contains("202411")) { nr["１１月"] = row["202411"]; sum += Convert.ToDecimal(row["202411"]); }
                if (row.Table.Columns.Contains("202412")) { nr["１２月"] = row["202412"]; sum += Convert.ToDecimal(row["202412"]); }
                if (row.Table.Columns.Contains("202501")) { nr["１月"] = row["202501"]; sum += Convert.ToDecimal(row["202501"]); }
                if (row.Table.Columns.Contains("202502")) { nr["２月"] = row["202502"]; sum += Convert.ToDecimal(row["202502"]); }
                if (row.Table.Columns.Contains("202503")) { nr["３月"] = row["202503"]; sum += Convert.ToDecimal(row["202503"]); }


                if (row[0].ToString() == "現場計数" || row[0].ToString() == "部門計数")
                {

                }
                else
                {
                    nr["年間"] = sum;
                    nr["月平均"] = Math.Round(sum / 12);
                }

                //if (row[0].ToString() == "従業員数" || row[0].ToString() == "労働生産性")
                //{
                //    Int16 ct_sum = 0;

                //    if (row.Table.Columns.Contains("202404")) if (Convert.ToDecimal(row["202404"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202405")) if (Convert.ToDecimal(row["202405"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202406")) if (Convert.ToDecimal(row["202406"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202407")) if (Convert.ToDecimal(row["202407"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202408")) if (Convert.ToDecimal(row["202408"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202409")) if (Convert.ToDecimal(row["202409"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202410")) if (Convert.ToDecimal(row["202410"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202411")) if (Convert.ToDecimal(row["202411"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202412")) if (Convert.ToDecimal(row["202412"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202501")) if (Convert.ToDecimal(row["202501"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202502")) if (Convert.ToDecimal(row["202502"]) > 0) { ct_sum++; }
                //    if (row.Table.Columns.Contains("202503")) if (Convert.ToDecimal(row["202503"]) > 0) { ct_sum++; }

                //    nr["年間"] = ct_sum == 0 ? sum : sum / ct_sum;
                //}
                //else
                //{
                //    nr["年間"] = sum;
                //    nr["年間"] = sum / 12;
                //}

                gendt.Rows.Add(nr);


            }
            return gendt;
        }

        private void dataGridView1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {

        }
    }
}
