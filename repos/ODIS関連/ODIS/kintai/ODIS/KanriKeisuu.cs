using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.IO;
using C1.C1Excel;

namespace ODIS.ODIS
{
    public partial class KanriKeisuu : Form
    {


        //期
        private string maey = "2025";
        private string atoy = "2026";

        private string maeyex = "";
        private string atoyex = "";

        private decimal souuri = 0;
        private decimal sourieki = 0;

        private string result;

        string[,] ar = new string[500, 2];

        //次
        private string zi = "";

        private string bumoncd = "";
        private string genbacd = "";


        public KanriKeisuu()
        {
            InitializeComponent();

            //TODO 毎年追加
            cbki.Items.Add("53期(2024)");
            cbki.Items.Add("54期(2025)");

            cbki.SelectedIndex = cbki.Items.Count - 1;

            IniSet();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dgvlist.Font = new Font(dgvlist.Font.Name, 9);

            dgvkoyosan.Font = new Font(dgvkoyosan.Font.Name, 9);
            dgvkozenki.Font = new Font(dgvkozenki.Font.Name, 9);
            dgvkozisseki.Font = new Font(dgvkozisseki.Font.Name, 9);

            dgvtouyosan.Font = new Font(dgvtouyosan.Font.Name, 9);
            dgvtouzenki.Font = new Font(dgvtouzenki.Font.Name, 9);
            dgvtouzisseki.Font = new Font(dgvtouzisseki.Font.Name, 9);

            //行ヘッダを非表示
            //dataGridView1.RowHeadersVisible = false;

            dgvkoyosan.RowHeadersWidth = 33;
            dgvkozenki.RowHeadersWidth = 33;

            dgvtouyosan.RowHeadersWidth = 33;
            dgvtouzenki.RowHeadersWidth = 33;

            //グリッドビューのコピー
            dgvlist.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            dgvkozisseki.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvkoyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvkozenki.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            dgvtouzisseki.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvtouyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvtouzisseki.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //行非表示にさせるためフォーカス無に
            //dgvlist.CurrentCell = null;



            Com.InHistory("51_管理計数", "","");

        }

        private void IniSet()
        {
            maey = cbki.SelectedItem.ToString().Substring(4, 4);
            atoy = (Convert.ToInt16(maey) + 1).ToString();

            DataTable mokutable = new DataTable();
            mokutable = Com.GetDB("select 次 from dbo.y予算マスタ where 始年 = '" + maey + "' order by 次 ");

            cbzi.Items.Clear();

            foreach (DataRow row in mokutable.Rows)
            {
                cbzi.Items.Add("第" + row[0] + "次");
                zi = row[0].ToString();
            }

            cbzi.SelectedIndex = cbzi.Items.Count - 1;


            ////予算の次を取得
            //DataTable dtzi = new DataTable();
            //dtzi = Com.GetDB(" select max(次) from dbo.y予算マスタ where 始年 = '" + maey + "' ");
            //zi = dtzi.Rows[0][0].ToString();

            maeyex = (Convert.ToInt16(maey) - 1).ToString();
            atoyex = (Convert.ToInt16(atoy) - 1).ToString();


            //構成比で使用 年間総合計額
            GetSumData();

            SetTiku();
            SetBumon();
            SetGenba();
            GetData();
        }

        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD  where 年月 between '" + maey + "04' and '" + atoy + "03' order by 担当区分";

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
            string sql = "select distinct 職種 from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD  where 年月 between '" + maey + "04' and '" + atoy + "03'";

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
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct b.現場CD, b.現場名 from dbo.kanrikeisuu a left join dbo.担当テーブル b on a.部門CD = b.組織CD and a.現場CD = b.現場CD where 年月 between '" + maey + "04' and '" + atoy + "03'";

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

        //構成比出力のために使用　年間総合計額
        private void GetSumData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable sumdt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select sum(売上) as 売上, sum(利益) as 利益 from dbo.kanrikeisuu where 年月 between '" + maey + "04' and '" + atoy + "03'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(sumdt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            //データ無場合のエラー回避
            if (sumdt.Rows.Count == 0)
            {
                //総売上と総利益
                foreach (DataRow row in sumdt.Rows)
                {
                    souuri = Convert.ToDecimal(row[0]);
                    sourieki = Convert.ToDecimal(row[1]);
                }
            }
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
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //部門
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        result += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            //職種
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 職種 <> '" + checkedListBox2.Items[i].ToString() + "'";
                    //flg = true;
                }
            }

            //現場
            int itemcount = checkedListBox3.Items.Count; //項目数合計
            int ckcount = checkedListBox3.CheckedItems.Count; //チェック項目数
            //アイテム数合計(200)の半分(100)より、チェック数が少ない場合はチェック無を条件にsql作成

            if (ckcount == 0)
            {
                result = " and 現場CD = '99999' ";
                return;
            }

            if (itemcount / 2 > ckcount)
            {
                string sql3 = "";
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (checkedListBox3.GetItemChecked(i))
                    {
                        if (checkedListBox3.Items[i].ToString().Length < 5)
                        {
                            sql3 += " or isnull(現場CD,'') = '' ";
                        }
                        else
                        {
                            sql3 += " or isnull(現場CD,'') = '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                        }
                    }
                }

                if (sql3.Length > 0) result += " and ( " + sql3.Substring(4) + " ) ";

            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (!checkedListBox3.GetItemChecked(i))
                    {
                        if (checkedListBox3.Items[i].ToString().Length < 5)
                        {
                            result += " and isnull(現場CD,'') <> '' ";
                        }
                        else
                        {
                            result += " and isnull(現場CD,'') <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                        }
                    }
                }
            }


        }

        //検索結果一覧表示
        private void GetData()
        {
            DataTable dt = new DataTable();

            //TODO 引当対応
            if (cbhikiate.Checked)
            {
                dt = Com.GetDB("select * from [dbo].[k管理計数_年間集計_引当金無]('" + maey + "04', '" + atoy + "03','" + zi + "') order by 部門CD,現場CD");
            }
            else
            { 
                dt = Com.GetDB("select * from [dbo].[k管理計数_年間集計]('" + maey + "04', '" + atoy + "03','" + zi + "') order by 部門CD,現場CD");
            }

            DataTable sumdt = new DataTable();

            //検索文字列処理
            ResultStr();

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            //if (dtrow.Length == 0) return;

            DataTable Disp = new DataTable();
            Disp.Columns.Add("組織CD", typeof(string));
            Disp.Columns.Add("現場CD", typeof(string));
            //Disp.Columns.Add("地区名", typeof(string));
            Disp.Columns.Add("組織名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));

            Disp.Columns.Add("売上", typeof(decimal));
            Disp.Columns.Add("経費", typeof(decimal));
            Disp.Columns.Add("利益", typeof(decimal));

            //Disp.Columns.Add("労働生産性", typeof(decimal));
            //Disp.Columns.Add("売上順位", typeof(decimal));
            //Disp.Columns.Add("利益順位", typeof(decimal));

            //Disp.Columns.Add("実績利益", typeof(decimal));
            Disp.Columns.Add("計数", typeof(decimal));

            Disp.Columns.Add("予算利益", typeof(decimal));
            Disp.Columns.Add("予算利益差", typeof(decimal));
            Disp.Columns.Add("前期利益", typeof(decimal));
            Disp.Columns.Add("前期利益差", typeof(decimal));

            Disp.Columns.Add("予算計数", typeof(decimal));
            Disp.Columns.Add("予算計数差", typeof(decimal));
            Disp.Columns.Add("前期計数", typeof(decimal));
            Disp.Columns.Add("前期計数差", typeof(decimal));
            Disp.Columns.Add("部門", typeof(string));
            Disp.Columns.Add("職種", typeof(string));


            //今期
            decimal uriage = 0;
            //decimal keihi = 0;
            decimal rieki = 0;

            decimal uriagekou = 0;
            decimal riekikou = 0;


            int i = 0;
            foreach (DataRow row in dtrow)
            {
                DataRow nr = Disp.NewRow();
                //nr["地区名"] = row["地区名"];

                nr["組織CD"] = row["部門CD"];
                nr["現場CD"] = row["現場CD"];

                nr["組織名"] = row["部門名"];
                nr["現場名"] = row["現場名"];
                nr["売上"] = Convert.ToDecimal(row["売上"]);
                nr["経費"] = Convert.ToDecimal(row["経費"]);
                nr["利益"] = Convert.ToDecimal(row["利益"]);
                nr["計数"] = Convert.ToDecimal(row["実績計数"]);
                //nr["労働生産性"] = Convert.ToDecimal(row["労働生産性"]);
                //nr["売上順位"] = Convert.ToDecimal(row["売上順位"]);
                //nr["利益順位"] = Convert.ToDecimal(row["利益順位"]);

                //nr["実績計数"] = row["実績計数"];

                nr["予算利益"] = row["予算利益"];
                nr["予算利益差"] = row["予算利益差"];
                nr["前期利益"] = row["前期利益"];
                nr["前期利益差"] = row["前期利益差"];

                nr["予算計数"] = row["予算計数"];
                nr["予算計数差"] = row["予算計数差"];
                nr["前期計数"] = row["前期計数"];
                nr["前期計数差"] = row["前期計数差"];

                //nr["実績利益"] = row["実績利益"];
                nr["部門"] = row["担当区分"];
                nr["職種"] = row["職種"];



                ar[i, 0] = row["部門CD"].ToString();
                ar[i, 1] = row["現場CD"].ToString();

                Disp.Rows.Add(nr);

                //今期
                uriage += Convert.ToDecimal(row["売上"]);
                rieki += Convert.ToDecimal(row["実績利益"]);

                uriagekou += souuri == 0 ? 0 : Convert.ToDecimal(row["売上"]) / souuri;
                riekikou += sourieki == 0 ? 0 : Convert.ToDecimal(row["実績利益"]) / sourieki;

                i++;
            }

            dgvlist.DataSource = Disp;

            dgvlist.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft; 
            dgvlist.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft; 
            
            dgvlist.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dgvlist.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            
            dgvlist.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvlist.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dgvlist.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvlist.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dgvlist.Columns[4].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[5].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[6].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[7].DefaultCellStyle.Format = "0.00\'%\'";

            dgvlist.Columns[8].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[9].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[10].DefaultCellStyle.Format = "#,0";
            dgvlist.Columns[11].DefaultCellStyle.Format = "#,0";

            dgvlist.Columns[12].DefaultCellStyle.Format = "0.00\'%\'";
            dgvlist.Columns[13].DefaultCellStyle.Format = "0.00\'%\'";
            dgvlist.Columns[14].DefaultCellStyle.Format = "0.00\'%\'";
            dgvlist.Columns[15].DefaultCellStyle.Format = "0.00\'%\'";


            //dgvlist.Columns[1].Width = 80;
            dgvlist.Columns[2].Width = 80;
            dgvlist.Columns[3].Width = 300;

            dgvlist.Columns[4].Width = 80;
            dgvlist.Columns[5].Width = 80;
            dgvlist.Columns[6].Width = 80;
            dgvlist.Columns[7].Width = 60;

            dgvlist.Columns[8].Width = 80;
            dgvlist.Columns[9].Width = 70;
            dgvlist.Columns[10].Width = 80;
            dgvlist.Columns[11].Width = 70;

            dgvlist.Columns[12].Width = 60;
            dgvlist.Columns[13].Width = 60;
            dgvlist.Columns[14].Width = 60;
            dgvlist.Columns[15].Width = 60;

            dgvlist.Columns[0].Visible = false;
            dgvlist.Columns[1].Visible = false;

            //利益
            dgvlist.Columns[8].HeaderCell.Style.BackColor = Color.OldLace;
            dgvlist.Columns[9].HeaderCell.Style.BackColor = Color.OldLace;
            dgvlist.Columns[10].HeaderCell.Style.BackColor = Color.OldLace;
            dgvlist.Columns[11].HeaderCell.Style.BackColor = Color.OldLace;

            dgvlist.Columns[12].HeaderCell.Style.BackColor = Color.PaleTurquoise;
            dgvlist.Columns[13].HeaderCell.Style.BackColor = Color.PaleTurquoise;
            dgvlist.Columns[14].HeaderCell.Style.BackColor = Color.PaleTurquoise;
            dgvlist.Columns[15].HeaderCell.Style.BackColor = Color.PaleTurquoise;

        }

        //検索ボタンクリック
        private void GetDataFirst()
        {
            Cursor.Current = Cursors.WaitCursor;

            //個別
            GetData();

            if (tabControl1.SelectedTab == tabPage2)
            {
                //合計
                GetDataTotal();
            }


            if (dgvlist.RowCount == 0)
            {
                label2.Text = "";
                label4.Text = "";
                dgvkozisseki.DataSource = null;
                dgvtouzisseki.DataSource = null;
            }

            //カーソル変更・メッセージキュー処理
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void GetToukeiZissekiData()
        {
            DataTable dtSum = new DataTable();
            String sql = "";

            if (cbhikiate.Checked)
            {
                sql = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(売上) as 売上 ";
                sql += ", sum(人件費) as 人件費, sum(諸経費) as 諸経費, sum(人件費 + 諸経費) as 経費, sum(売上) - sum(人件費 + 諸経費) as 利益 ";

                sql += ", case when SUM(売上) = 0 then 0 else sum(人件費 + 諸経費) / sum(売上) * 100 end as 計数 ";
                sql += ", case when SUM(売上) = 0 then 0 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 100 then 1 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 90 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 100 then 2 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 85 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 90 then 3 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 80 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 85 then 4 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 70 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 80 then 5 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 60 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 70 then 6 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 50 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 60 then 7 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 < 50 then 8 ";
                sql += "else 9 end as 評価 ";

                sql += ", sum(管理人件費) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理人件費 + 管理諸経費) as 管理経費, sum(売上 - 人件費 - 諸経費 - 管理人件費 - 管理諸経費) as 管理利益 ";

                sql += ", case when SUM(売上) = 0 then 0 else sum(人件費 + 諸経費 + 管理人件費 + 管理諸経費) / sum(売上) * 100 end as 管理計数 ";

                sql += ", SUM(従業員数) as 従業員数 ";
                sql += ", case when SUM(従業員数) = 0 then 0 else (SUM(売上)-SUM(諸経費))/SUM(従業員数) end as 労働生産性 , SUM(労働時間) as 労働時間 ";
                sql += " from dbo.c管理計数 where 年月 between '" + maey + "04' and '" + atoy + "03' ";
            }
            else
            {
                sql = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(売上) as 売上 ";
                sql += ", sum(人件費 + 賞与 + 退職金) as 人件費, sum(諸経費) as 諸経費, sum(経費) as 経費, sum(利益) as 利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(経費) / sum(売上) * 100 end as 計数 ";
                sql += ", case when SUM(売上) = 0 then 0 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 100 then 1 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 90 and SUM(経費)/SUM(売上)*100 < 100 then 2 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 85 and SUM(経費)/SUM(売上)*100 < 90 then 3 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 80 and SUM(経費)/SUM(売上)*100 < 85 then 4 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 70 and SUM(経費)/SUM(売上)*100 < 80 then 5 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 60 and SUM(経費)/SUM(売上)*100 < 70 then 6 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 50 and SUM(経費)/SUM(売上)*100 < 60 then 7 ";
                sql += "when SUM(経費)/SUM(売上)*100 < 50 then 8 ";
                sql += "else 9 end as 評価 ";
                sql += ", sum(管理人件費 + 管理賞与 + 管理退職金) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理経費) as 管理経費, sum(利益 - 管理経費) as 管理利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(経費 + 管理経費) / sum(売上) * 100 end as 管理計数 ";

                sql += ", SUM(従業員数) as 従業員数 ";
                sql += ", case when SUM(従業員数) = 0 then 0 else (SUM(売上)-SUM(諸経費))/SUM(従業員数) end as 労働生産性 , SUM(労働時間) as 労働時間 ";
                sql += " from dbo.c管理計数 where 年月 between '" + maey + "04' and '" + atoy + "03' ";
            }

            //検索文字列処理
            ResultStr();

            sql += " " + result + " group by 年月 ";

            dtSum = Com.GetDB(sql);

            //縦横変換
            DataTable dtTotal = new DataTable();
            dtTotal = MonthDataTable(dtSum);

            #region 年間～下期までの計数と評価の設定
            //decimal nen_mokuromi_keisuu = 0;
            decimal nen_zisseki_keisuu = 0;
            int nen_zisseki_rank = 0;

            decimal nen_zisseki_kanrikeisuu = 0;

            nen_zisseki_keisuu = Convert.ToDecimal(dtTotal.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(dtTotal.Rows[5]["年間"]) / Convert.ToDecimal(dtTotal.Rows[2]["年間"]) * 100;
            nen_zisseki_rank = Com.GetLevel(nen_zisseki_keisuu);

            nen_zisseki_kanrikeisuu = Convert.ToDecimal(dtTotal.Rows[2]["年間"]) == 0 ? 0 : (Convert.ToDecimal(dtTotal.Rows[5]["年間"]) + Convert.ToDecimal(dtTotal.Rows[11]["年間"])) / Convert.ToDecimal(dtTotal.Rows[2]["年間"]) * 100;

            dtTotal.Rows[7]["年間"] = nen_zisseki_keisuu;
            dtTotal.Rows[8]["年間"] = nen_zisseki_rank;

            dtTotal.Rows[13]["年間"] = nen_zisseki_kanrikeisuu;
            #endregion


            dgvtouzisseki.DataSource = dtTotal;

            //表示加工処理
            for (int i = 0; i < dtTotal.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvtouzisseki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvtouzisseki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvtouzisseki.Columns[i].Width = 80;

                //三桁区切り表示
                dgvtouzisseki.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvtouzisseki.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvtouzisseki.Columns[13].Width = 90;

            //行非表示にさせるためフォーカス無に
            dgvtouzisseki.CurrentCell = null;

            dgvtouzisseki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvtouzisseki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvtouzisseki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvtouzisseki.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            dgvtouzisseki.Rows[11].DefaultCellStyle.BackColor = Color.Khaki; //管理経費
            dgvtouzisseki.Rows[12].DefaultCellStyle.BackColor = Color.PaleTurquoise;//管理利益
            dgvtouzisseki.Rows[13].DefaultCellStyle.Format = "0.00\'%\'"; //管理計数

            dgvtouzisseki.Rows[16].DefaultCellStyle.Format = "N1";

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvtouzisseki.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;


            //DataGridView1の左側2列を固定する
            dgvtouzisseki.Columns[0].Frozen = true;
        }

        private void GetToukeiYosanData()
        {
            DataTable dtSumm = new DataTable();
            String sqlm = "";

            if (cbhikiate.Checked)
            {
                sqlm = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(固定売上 + 臨時売上) as 売上 ";
                sqlm += ", sum(人件費) as 人件費, sum(諸経費) as 諸経費, sum(人件費 + 諸経費) as 経費, sum(固定売上 + 臨時売上 - 人件費 - 諸経費) as 利益 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 else sum(人件費 + 諸経費) / sum(固定売上 + 臨時売上) * 100 end as 計数 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 100 then 1 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 90 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 100 then 2 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 85 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 90 then 3 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 80 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 85 then 4 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 70 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 80 then 5 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 60 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 70 then 6 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 50 and SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 60 then 7 ";
                sqlm += "when SUM(人件費 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 50 then 8 ";
                sqlm += "else 9 end as 評価 ";
                sqlm += ", sum(管理人件費) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理人件費 + 管理諸経費) as 管理経費, sum(固定売上 + 臨時売上 - 人件費 - 諸経費 - (管理人件費 + 管理諸経費)) as 管理利益 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 else sum(人件費 + 諸経費 + 管理人件費 + 管理諸経費) / sum(固定売上 + 臨時売上) * 100 end as 管理計数 ";
            }
            else
            {
                sqlm = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(固定売上 + 臨時売上) as 売上 ";
                sqlm += ", sum(人件費 + 賞与 + 退職金) as 人件費, sum(諸経費) as 諸経費, sum(人件費  + 賞与 + 退職金 + 諸経費) as 経費, sum(固定売上 + 臨時売上 - 人件費 - 賞与 - 退職金 - 諸経費) as 利益 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 else sum(人件費 + 賞与 + 退職金 + 諸経費) / sum(固定売上 + 臨時売上) * 100 end as 計数 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 100 then 1 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 90 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 100 then 2 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 85 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 90 then 3 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 80 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 85 then 4 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 70 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 80 then 5 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 60 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 70 then 6 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 >= 50 and SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 60 then 7 ";
                sqlm += "when SUM(人件費 + 賞与 + 退職金 + 諸経費)/SUM(固定売上 + 臨時売上)*100 < 50 then 8 ";
                sqlm += "else 9 end as 評価 ";
                sqlm += ", sum(管理人件費 + 管理賞与 + 管理退職金) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理人件費 + 管理賞与 + 管理退職金 + 管理諸経費) as 管理経費, sum(固定売上 + 臨時売上 - 人件費 - 賞与 - 退職金 - 諸経費 - (管理人件費 + 管理賞与 + 管理退職金 + 管理諸経費)) as 管理利益 ";
                sqlm += ", case when SUM(固定売上 + 臨時売上) = 0 then 0 else sum(人件費 + 賞与 + 退職金 + 諸経費 + 管理人件費 + 管理賞与 + 管理退職金 + 管理諸経費) / sum(固定売上 + 臨時売上) * 100 end as 管理計数 ";
            }
            sqlm += " from dbo.y予算 where 年月 between '" + maey + "04' and '" + atoy + "03' and 次 = '" + zi + "'";

            //検索文字列処理
            ResultStr();

            sqlm += " " + result + " group by 年月 ";

            dtSumm = Com.GetDB(sqlm);

            //縦横変換
            DataTable dtTotalm = new DataTable();
            dtTotalm = MonthDataTable(dtSumm);

            #region 年間～下期までの計数と評価の設定
            //decimal nen_mokuromi_keisuu = 0;
            decimal nen_zisseki_keisuum = 0;
            int nen_zisseki_rankm = 0;

            decimal nen_zisseki_kanrikeisuum = 0;

            nen_zisseki_keisuum = Convert.ToDecimal(dtTotalm.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(dtTotalm.Rows[5]["年間"]) / Convert.ToDecimal(dtTotalm.Rows[2]["年間"]) * 100;
            nen_zisseki_rankm = Com.GetLevel(nen_zisseki_keisuum);

            nen_zisseki_kanrikeisuum = Convert.ToDecimal(dtTotalm.Rows[2]["年間"]) == 0 ? 0 : (Convert.ToDecimal(dtTotalm.Rows[5]["年間"]) + Convert.ToDecimal(dtTotalm.Rows[11]["年間"])) / Convert.ToDecimal(dtTotalm.Rows[2]["年間"]) * 100;

            dtTotalm.Rows[7]["年間"] = nen_zisseki_keisuum;
            dtTotalm.Rows[8]["年間"] = nen_zisseki_rankm;

            dtTotalm.Rows[13]["年間"] = nen_zisseki_kanrikeisuum;
            #endregion


            dgvtouyosan.DataSource = dtTotalm;

            //表示加工処理
            for (int i = 0; i < dtTotalm.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvtouyosan.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvtouyosan.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvtouyosan.Columns[i].Width = 80;

                //三桁区切り表示
                dgvtouyosan.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvtouyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvtouyosan.Columns[13].Width = 90;

            //行非表示にさせるためフォーカス無に
            dgvtouyosan.CurrentCell = null;

            dgvtouyosan.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvtouyosan.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvtouyosan.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvtouyosan.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            dgvtouyosan.Rows[11].DefaultCellStyle.BackColor = Color.Khaki; //管理経費
            dgvtouyosan.Rows[12].DefaultCellStyle.BackColor = Color.PaleTurquoise;//管理利益

            dgvtouyosan.Rows[13].DefaultCellStyle.Format = "0.00\'%\'"; //管理計数

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvtouyosan.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;


            //DataGridView1の左側2列を固定する
            dgvtouyosan.Columns[0].Frozen = true;
        }

        private void GetToukeiZenkiData()
        {
            DataTable dtSum = new DataTable();
            String sql = "";

            if (cbhikiate.Checked)
            {
                sql = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(売上) as 売上 ";
                sql += ", sum(人件費) as 人件費, sum(諸経費) as 諸経費, sum(人件費 + 諸経費) as 経費, sum(売上) - sum(人件費 + 諸経費) as 利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(人件費 + 諸経費) / sum(売上) * 100 end as 計数 ";
                sql += ", case when SUM(売上) = 0 then 0 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 100 then 1 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 90 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 100 then 2 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 85 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 90 then 3 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 80 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 85 then 4 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 70 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 80 then 5 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 60 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 70 then 6 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 >= 50 and SUM(人件費 + 諸経費)/SUM(売上)*100 < 60 then 7 ";
                sql += "when SUM(人件費 + 諸経費)/SUM(売上)*100 < 50 then 8 ";
                sql += "else 9 end as 評価 ";
                sql += ", sum(管理人件費) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理人件費 + 管理諸経費) as 管理経費, sum(売上 - 人件費 - 諸経費 - 管理人件費 - 管理諸経費) as 管理利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(人件費 + 諸経費 + 管理人件費 + 管理諸経費) / sum(売上) * 100 end as 管理計数 ";
                sql += ", SUM(従業員数) as 従業員数 ";
                sql += ", case when SUM(従業員数) = 0 then 0 else (SUM(売上)-SUM(諸経費))/SUM(従業員数) end as 労働生産性 , SUM(労働時間) as 労働時間 ";
                sql += " from dbo.c管理計数 where 年月 between '" + maeyex + "04' and '" + atoyex + "03' ";
            }
            else
            {
                sql = "select 年月, sum(固定売上) as 固定売上, sum(臨時売上) as 臨時売上, sum(売上) as 売上 ";
                sql += ", sum(人件費 + 賞与 + 退職金) as 人件費, sum(諸経費) as 諸経費, sum(経費) as 経費, sum(利益) as 利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(経費) / sum(売上) * 100 end as 計数 ";
                sql += ", case when SUM(売上) = 0 then 0 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 100 then 1 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 90 and SUM(経費)/SUM(売上)*100 < 100 then 2 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 85 and SUM(経費)/SUM(売上)*100 < 90 then 3 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 80 and SUM(経費)/SUM(売上)*100 < 85 then 4 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 70 and SUM(経費)/SUM(売上)*100 < 80 then 5 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 60 and SUM(経費)/SUM(売上)*100 < 70 then 6 ";
                sql += "when SUM(経費)/SUM(売上)*100 >= 50 and SUM(経費)/SUM(売上)*100 < 60 then 7 ";
                sql += "when SUM(経費)/SUM(売上)*100 < 50 then 8 ";
                sql += "else 9 end as 評価 ";
                sql += ", sum(管理人件費 + 管理賞与 + 管理退職金) as 管理人件費, sum(管理諸経費) as 管理諸経費, sum(管理経費) as 管理経費, sum(利益 - 管理経費) as 管理利益 ";
                sql += ", case when SUM(売上) = 0 then 0 else sum(経費 + 管理経費) / sum(売上) * 100 end as 管理計数 ";
                sql += ", SUM(従業員数) as 従業員数 ";
                sql += ", case when SUM(従業員数) = 0 then 0 else (SUM(売上)-SUM(諸経費))/SUM(従業員数) end as 労働生産性 , SUM(労働時間) as 労働時間 ";
                sql += " from dbo.c管理計数 where 年月 between '" + maeyex + "04' and '" + atoyex + "03' ";
            }
            //検索文字列処理
            ResultStr();

            sql += " " + result + " group by 年月 ";

            dtSum = Com.GetDB(sql);

            //縦横変換
            DataTable dtTotal = new DataTable();
            dtTotal = MonthDataTableEx(dtSum);

            #region 年間～下期までの計数と評価の設定
            //decimal nen_mokuromi_keisuu = 0;
            decimal nen_zisseki_keisuu = 0;
            int nen_zisseki_rank = 0;

            decimal nen_zisseki_kanrikeisuu = 0;

            nen_zisseki_keisuu = Convert.ToDecimal(dtTotal.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(dtTotal.Rows[5]["年間"]) / Convert.ToDecimal(dtTotal.Rows[2]["年間"]) * 100;
            nen_zisseki_rank = Com.GetLevel(nen_zisseki_keisuu);

            nen_zisseki_kanrikeisuu = Convert.ToDecimal(dtTotal.Rows[2]["年間"]) == 0 ? 0 : (Convert.ToDecimal(dtTotal.Rows[5]["年間"]) + Convert.ToDecimal(dtTotal.Rows[11]["年間"])) / Convert.ToDecimal(dtTotal.Rows[2]["年間"]) * 100;

            dtTotal.Rows[7]["年間"] = nen_zisseki_keisuu;
            dtTotal.Rows[8]["年間"] = nen_zisseki_rank;

            dtTotal.Rows[13]["年間"] = nen_zisseki_kanrikeisuu;
            #endregion


            dgvtouzenki.DataSource = dtTotal;

            //表示加工処理
            for (int i = 0; i < dtTotal.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvtouzenki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvtouzenki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvtouzenki.Columns[i].Width = 80;

                //三桁区切り表示
                dgvtouzenki.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvtouzenki.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvtouzenki.Columns[13].Width = 90;

            //行非表示にさせるためフォーカス無に
            dgvtouzenki.CurrentCell = null;

            dgvtouzenki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvtouzenki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvtouzenki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvtouzenki.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            dgvtouzenki.Rows[11].DefaultCellStyle.BackColor = Color.Khaki; //管理経費
            dgvtouzenki.Rows[12].DefaultCellStyle.BackColor = Color.PaleTurquoise;//管理利益

            dgvtouzenki.Rows[13].DefaultCellStyle.Format = "0.00\'%\'"; //管理計数

            dgvtouzenki.Rows[16].DefaultCellStyle.Format = "N1";

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvtouzenki.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;


            //DataGridView1の左側2列を固定する
            dgvtouzenki.Columns[0].Frozen = true;
        }

        //合計データ
        private void GetDataTotal()
        {
            GetToukeiZissekiData();

            if (tabControl3.SelectedIndex.ToString() == "0")
            {
                //予算
                GetToukeiYosanData();
            }
            else
            {
                //前期
                GetToukeiZenkiData();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //カーソル変更
                Cursor.Current = Cursors.WaitCursor;

                //データ処理
                GetData();

                //合計 or 予算
                if (tabControl1.SelectedTab == tabPage2)
                {
                    //合計
                    GetDataTotal();
                }

                //カーソル変更・メッセージキュー処理
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }
        }

        private DataRowView _drv = null;

        private void GetKobetsuZissekiData()
        {
            //【実績表示】
            DataTable dt = new DataTable();

            if (cbhikiate.Checked)
            {
                dt = Com.GetDB("select * from dbo.c管理計数_個別_引当無('" + maey + "', '" + atoy + "', '" + bumoncd + "','" + genbacd + "')");
            }
            else
            {
                dt = Com.GetDB("select * from dbo.c管理計数_個別('" + maey + "', '" + atoy + "', '" + bumoncd + "','" + genbacd + "')");
            }

            DataTable gendt = new DataTable();
            gendt = MonthDataTable(dt);

            #region 年間～下期までの計数と評価の設定
            decimal nen_zisseki_keisuu = 0;
            int nen_zisseki_rank = 0;

            //decimal nen_zisseki_kanrikeisuu = 0;

            nen_zisseki_keisuu = Convert.ToDecimal(gendt.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(gendt.Rows[5]["年間"]) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;

            nen_zisseki_rank = Com.GetLevel(nen_zisseki_keisuu);

            //nen_zisseki_kanrikeisuu = Convert.ToDecimal(gendt.Rows[2]["年間"]) == 0 ? 0 : (Convert.ToDecimal(gendt.Rows[5]["年間"]) + Convert.ToDecimal(gendt.Rows[11]["年間"])) / Convert.ToDecimal(gendt.Rows[2]["年間"]) * 100;

            gendt.Rows[7]["年間"] = nen_zisseki_keisuu;
            gendt.Rows[8]["年間"] = nen_zisseki_rank;

            //gendt.Rows[13]["年間"] = nen_zisseki_kanrikeisuu;

            #endregion

            //月別タブのデータグリッドビュー
            dgvkozisseki.DataSource = gendt;

            //月別タブの表示設定
            for (int i = 0; i < gendt.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvkozisseki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvkozisseki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvkozisseki.Columns[i].Width = 80;

                //三桁区切り表示
                dgvkozisseki.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvkozisseki.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvkozisseki.Columns[13].Width = 90;

            dgvkozisseki.Rows[11].DefaultCellStyle.Format = "N1";

            //行非表示にさせるためフォーカス無に
            //dgvkozisseki.CurrentCell = null;

            //dgvkozisseki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            //dgvkozisseki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            //dgvkozisseki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            //dgvkozisseki.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            ////ソート禁止設定
            //foreach (DataGridViewColumn c in dgvkozisseki.Columns)
            //    c.SortMode = DataGridViewColumnSortMode.NotSortable;

            ////DataGridView1の左側2列を固定する
            //dgvkozisseki.Columns[0].Frozen = true;

        }

        private void GetKobetsuYosanData()
        {
            //【予算表示】
            DataTable dtsub = new DataTable();

            if (cbhikiate.Checked)
            {
                dtsub = Com.GetDB("select * from dbo.c管理計数_予算_引当無('" + maey + "', '" + atoy + "', '" + bumoncd + "','" + genbacd + "', '" + zi + "')");
            }
            else
            {
                dtsub = Com.GetDB("select * from dbo.c管理計数_予算('" + maey + "', '" + atoy + "', '" + bumoncd + "','" + genbacd + "', '" + zi + "')");
            }
            DataTable gendtsub = new DataTable();
            gendtsub = MonthDataTable(dtsub);


            //年間の計数と評価の設定
            decimal nen_zisseki_keisuu_sub = 0;
            int nen_zisseki_rank_sub = 0;

            nen_zisseki_keisuu_sub = Convert.ToDecimal(gendtsub.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(gendtsub.Rows[5]["年間"]) / Convert.ToDecimal(gendtsub.Rows[2]["年間"]) * 100;

            nen_zisseki_rank_sub = Com.GetLevel(nen_zisseki_keisuu_sub);

            gendtsub.Rows[7]["年間"] = nen_zisseki_keisuu_sub;
            gendtsub.Rows[8]["年間"] = nen_zisseki_rank_sub;


            dgvkoyosan.DataSource = gendtsub;

            //表示設定
            for (int i = 0; i < gendtsub.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvkoyosan.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvkoyosan.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvkoyosan.Columns[i].Width = 80;

                //三桁区切り表示
                dgvkoyosan.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvkoyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvkoyosan.Columns[13].Width = 90;

            //行非表示にさせるためフォーカス無に
            dgvkoyosan.CurrentCell = null;

        }

        private void GetKobetsuZenkiData()
        {
            //【前期表示】
            DataTable dtzen = new DataTable();

            if (cbhikiate.Checked)
            {
                dtzen = Com.GetDB("select * from dbo.c管理計数_個別_引当無('" + maeyex + "', '" + atoyex + "', '" + bumoncd + "','" + genbacd + "')");
            }
            else
            {
                dtzen = Com.GetDB("select * from dbo.c管理計数_個別('" + maeyex + "', '" + atoyex + "', '" + bumoncd + "','" + genbacd + "')");
            }

            DataTable dtzenchange = new DataTable();
            dtzenchange = MonthDataTableEx(dtzen);

            //年間の計数と評価の設定
            decimal nen_kozenki = 0;
            int nen_kozenki_rank_sub = 0;

            nen_kozenki = Convert.ToDecimal(dtzenchange.Rows[2]["年間"]) == 0 ? 0 : Convert.ToDecimal(dtzenchange.Rows[5]["年間"]) / Convert.ToDecimal(dtzenchange.Rows[2]["年間"]) * 100;

            nen_kozenki_rank_sub = Com.GetLevel(nen_kozenki);

            dtzenchange.Rows[7]["年間"] = nen_kozenki;
            dtzenchange.Rows[8]["年間"] = nen_kozenki_rank_sub;

            dgvkozenki.DataSource = dtzenchange;

            //表示設定
            for (int i = 0; i < dtzenchange.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dgvkozenki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dgvkozenki.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                dgvkozenki.Columns[i].Width = 80;

                //三桁区切り表示
                dgvkozenki.Columns[i].DefaultCellStyle.Format = "#,0";

                //ヘッダーの中央表示
                dgvkozenki.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvkozenki.Columns[13].Width = 90;

            //行非表示にさせるためフォーカス無に
            dgvkozenki.CurrentCell = null;

            dgvkozenki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvkozenki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvkozenki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvkozenki.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";


            dgvkozenki.Rows[11].DefaultCellStyle.Format = "N1";

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvkozenki.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;

            //DataGridView1の左側2列を固定する
            dgvkozenki.Columns[0].Frozen = true;
        }



        //検索結果での選択
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvlist.RowCount == 0)
            {
                label2.Text = "";
                label4.Text = "";
                return;
            }

            if (tabControl1.SelectedIndex.ToString() == "1") return;

            //ソート対応
            BindingManagerBase bm = dgvlist.BindingContext[dgvlist.DataSource, dgvlist.DataMember];

            if (bm.Count == 0) return;

            DataRowView drv = (DataRowView)bm.Current;

            //TODO 240524コメントアウト 
            //前回と同じならスルー
            if (_drv == drv)
            {
                //TODO 同じ行を選択したときも削除されるので、コメントアウト
                //dgvkozisseki.DataSource = "";
                //dgvkoyosan.DataSource = "";
                //label2.Text = "";
                //label4.Text = "";

                return;
            }

            _drv = drv;

            //個別
            label2.Text = drv.Row.ItemArray[2].ToString();
            label4.Text = drv.Row.ItemArray[3].ToString();

            bumoncd = drv.Row.ItemArray[0].ToString();
            genbacd = drv.Row.ItemArray[1].ToString();


            //GetKobetsuZissekiData(bumoncd, genbacd);
            GetKobetsuZissekiData();

            if (tabControl2.SelectedIndex.ToString() == "0")
            {
                //予算
                GetKobetsuYosanData();
            }
            else
            {
                //前期
                GetKobetsuZenkiData();
            }
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

            DataTable zendt = gendt.Clone();
            DataTable zenzendt = gendt.Clone();

            foreach (DataRow row in wkdt.Rows)
            {
                //年間
                decimal sum = 0;
                
                //列名の対応と、合算列の対応と、
                DataRow nr = gendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains(maey + "04")) { nr["４月"] = row[maey + "04"]; sum += Convert.ToDecimal(row[maey + "04"]);  }
                if (row.Table.Columns.Contains(maey + "05")) { nr["５月"] = row[maey + "05"]; sum += Convert.ToDecimal(row[maey + "05"]);  }
                if (row.Table.Columns.Contains(maey + "06")) { nr["６月"] = row[maey + "06"]; sum += Convert.ToDecimal(row[maey + "06"]);  }
                if (row.Table.Columns.Contains(maey + "07")) { nr["７月"] = row[maey + "07"]; sum += Convert.ToDecimal(row[maey + "07"]);  }
                if (row.Table.Columns.Contains(maey + "08")) { nr["８月"] = row[maey + "08"]; sum += Convert.ToDecimal(row[maey + "08"]);  }
                if (row.Table.Columns.Contains(maey + "09")) { nr["９月"] = row[maey + "09"]; sum += Convert.ToDecimal(row[maey + "09"]);  }
                if (row.Table.Columns.Contains(maey + "10")) { nr["１０月"] = row[maey + "10"]; sum += Convert.ToDecimal(row[maey + "10"]);  }
                if (row.Table.Columns.Contains(maey + "11")) { nr["１１月"] = row[maey + "11"]; sum += Convert.ToDecimal(row[maey + "11"]);  }
                if (row.Table.Columns.Contains(maey + "12")) { nr["１２月"] = row[maey + "12"]; sum += Convert.ToDecimal(row[maey + "12"]);  }
                if (row.Table.Columns.Contains(atoy + "01")) { nr["１月"] = row[atoy + "01"]; sum += Convert.ToDecimal(row[atoy + "01"]);  }
                if (row.Table.Columns.Contains(atoy + "02")) { nr["２月"] = row[atoy + "02"]; sum += Convert.ToDecimal(row[atoy + "02"]);  }
                if (row.Table.Columns.Contains(atoy + "03")) { nr["３月"] = row[atoy + "03"]; sum += Convert.ToDecimal(row[atoy + "03"]);  }

                if (row[0].ToString() == "従業員数" || row[0].ToString() == "労働生産性")
                {
                    Int16 ct_sum = 0;

                    if (row.Table.Columns.Contains(maey + "04")) if (Convert.ToDecimal(row[maey + "04"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "05")) if (Convert.ToDecimal(row[maey + "05"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "06")) if (Convert.ToDecimal(row[maey + "06"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "07")) if (Convert.ToDecimal(row[maey + "07"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "08")) if (Convert.ToDecimal(row[maey + "08"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "09")) if (Convert.ToDecimal(row[maey + "09"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "10")) if (Convert.ToDecimal(row[maey + "10"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "11")) if (Convert.ToDecimal(row[maey + "11"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maey + "12")) if (Convert.ToDecimal(row[maey + "12"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoy + "01")) if (Convert.ToDecimal(row[atoy + "01"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoy + "02")) if (Convert.ToDecimal(row[atoy + "02"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoy + "03")) if (Convert.ToDecimal(row[atoy + "03"]) > 0) { ct_sum++; }

                    nr["年間"] = ct_sum == 0 ? sum : sum / ct_sum;
                }
                else
                {
                    nr["年間"] = sum;
                }

                gendt.Rows.Add(nr);


            }
            return gendt;
        }


        private DataTable MonthDataTableEx(DataTable dt)
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

            DataTable zendt = gendt.Clone();
            DataTable zenzendt = gendt.Clone();

            foreach (DataRow row in wkdt.Rows)
            {
                //年間
                decimal sum = 0;

                //列名の対応と、合算列の対応と、
                DataRow nr = gendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains(maeyex + "04")) { nr["４月"] = row[maeyex + "04"]; sum += Convert.ToDecimal(row[maeyex + "04"]); }
                if (row.Table.Columns.Contains(maeyex + "05")) { nr["５月"] = row[maeyex + "05"]; sum += Convert.ToDecimal(row[maeyex + "05"]); }
                if (row.Table.Columns.Contains(maeyex + "06")) { nr["６月"] = row[maeyex + "06"]; sum += Convert.ToDecimal(row[maeyex + "06"]); }
                if (row.Table.Columns.Contains(maeyex + "07")) { nr["７月"] = row[maeyex + "07"]; sum += Convert.ToDecimal(row[maeyex + "07"]); }
                if (row.Table.Columns.Contains(maeyex + "08")) { nr["８月"] = row[maeyex + "08"]; sum += Convert.ToDecimal(row[maeyex + "08"]); }
                if (row.Table.Columns.Contains(maeyex + "09")) { nr["９月"] = row[maeyex + "09"]; sum += Convert.ToDecimal(row[maeyex + "09"]); }
                if (row.Table.Columns.Contains(maeyex + "10")) { nr["１０月"] = row[maeyex + "10"]; sum += Convert.ToDecimal(row[maeyex + "10"]); }
                if (row.Table.Columns.Contains(maeyex + "11")) { nr["１１月"] = row[maeyex + "11"]; sum += Convert.ToDecimal(row[maeyex + "11"]); }
                if (row.Table.Columns.Contains(maeyex + "12")) { nr["１２月"] = row[maeyex + "12"]; sum += Convert.ToDecimal(row[maeyex + "12"]); }
                if (row.Table.Columns.Contains(atoyex + "01")) { nr["１月"] = row[atoyex + "01"]; sum += Convert.ToDecimal(row[atoyex + "01"]); }
                if (row.Table.Columns.Contains(atoyex + "02")) { nr["２月"] = row[atoyex + "02"]; sum += Convert.ToDecimal(row[atoyex + "02"]); }
                if (row.Table.Columns.Contains(atoyex + "03")) { nr["３月"] = row[atoyex + "03"]; sum += Convert.ToDecimal(row[atoyex + "03"]); }

                if (row[0].ToString() == "従業員数" || row[0].ToString() == "労働生産性")
                {
                    Int16 ct_sum = 0;

                    if (row.Table.Columns.Contains(maeyex + "04")) if (Convert.ToDecimal(row[maeyex + "04"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "05")) if (Convert.ToDecimal(row[maeyex + "05"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "06")) if (Convert.ToDecimal(row[maeyex + "06"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "07")) if (Convert.ToDecimal(row[maeyex + "07"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "08")) if (Convert.ToDecimal(row[maeyex + "08"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "09")) if (Convert.ToDecimal(row[maeyex + "09"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "10")) if (Convert.ToDecimal(row[maeyex + "10"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "11")) if (Convert.ToDecimal(row[maeyex + "11"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(maeyex + "12")) if (Convert.ToDecimal(row[maeyex + "12"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoyex + "01")) if (Convert.ToDecimal(row[atoyex + "01"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoyex + "02")) if (Convert.ToDecimal(row[atoyex + "02"]) > 0) { ct_sum++; }
                    if (row.Table.Columns.Contains(atoyex + "03")) if (Convert.ToDecimal(row[atoyex + "03"]) > 0) { ct_sum++; }

                    nr["年間"] = ct_sum == 0 ? sum : sum / ct_sum;
                }
                else
                {
                    nr["年間"] = sum;
                }

                gendt.Rows.Add(nr);


            }
            return gendt;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //セルの列を確認
            decimal val = 0;
            if (e.Value != null && e.ColumnIndex > 3 && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }
            }
        }


        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }
        }



        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            return;

            //TODO 来年までコメントアウト
            //string bumon = _drv.Row.ItemArray[2].ToString();
            //string genba = _drv.Row.ItemArray[3].ToString();

            //string bumonCD = _drv.Row.ItemArray[0].ToString();
            //string genbaCD = _drv.Row.ItemArray[1].ToString();

            //KeisuuYM keiYM = new KeisuuYM(bumon, genba, bumonCD, genbaCD);
            //keiYM.ShowDialog();
        }


        private string ymChange(DataGridView dg)
        {
            string my = "";
            string ay = "";
            if (dg.Name == "dgvkozisseki" || dg.Name == "dgvtouzisseki")
            {
                my = maey;
                ay = atoy;
            }
            else
            {
                my = maeyex;
                ay = atoyex;
            }

            string ym = "";


            switch (dg.CurrentCell.OwningColumn.HeaderText)
            {
                case "４月": ym = my + "04"; break;
                case "５月": ym = my + "05"; break;
                case "６月": ym = my + "06"; break;
                case "７月": ym = my + "07"; break;
                case "８月": ym = my + "08"; break;
                case "９月": ym = my + "09"; break;
                case "１０月": ym = my + "10"; break;
                case "１１月": ym = my + "11"; break;
                case "１２月": ym = my + "12"; break;
                case "１月": ym = ay + "01"; break;
                case "２月": ym = ay + "02"; break;
                case "３月": ym = ay + "03"; break;
                case "年間": ym = "年間"; break;
                case "項目": ym = ""; break;
                default: ym = dg.CurrentCell.OwningColumn.HeaderText; break;
            }
            return ym;
        }




        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //タブコントロール無効化・カーソル変更
            tabControl1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string dum = tabControl1.SelectedIndex.ToString();

            if (dum == "0")
            {
                //TODO くるしまぎれ
                _drv = null;
                dataGridView1_SelectionChanged(sender, e);
            }
            else
            {
                GetDataTotal();
            }

            //カーソル変更・メッセージキュー処理・タブコントロール有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            tabControl1.Enabled = true;
        }

        //地区
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
            GetDataFirst();
        }

        //部門
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
            GetDataFirst();
        }

        //現場
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

            GetDataFirst();
        }


        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dgvkozisseki.GetClipboardContent());
            dgvkozisseki.ClearSelection();
        }

        //地区
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetBumon();
            SetGenba();
            GetDataFirst();
        }

        //部門
        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetGenba();
            GetDataFirst();
        }

        //現場
        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetDataFirst();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetDataFirst();
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //タブコントロール無効化・カーソル変更
            tabControl3.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string dum = tabControl3.SelectedIndex.ToString();

            if (dum == "0")
            {
                //予算
                GetToukeiYosanData();
            }
            else
            {
                //前期
                GetToukeiZenkiData();
            }

            //カーソル変更・メッセージキュー処理・タブコントロール有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            tabControl3.Enabled = true;
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //タブコントロール無効化・カーソル変更
            tabControl2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string dum = tabControl2.SelectedIndex.ToString();

            if (dum == "0")
            {
                //予算
                GetKobetsuYosanData();
            }
            else
            {
                //前期
                GetKobetsuZenkiData();
            }

            //カーソル変更・メッセージキュー処理・タブコントロール有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            tabControl2.Enabled = true;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button2.Enabled = false;

            string fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\計数\現場計数.xlsx";

            //手順1：新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            c1XLBook1.Load(fileName);

            string localPass = @"C:\ODIS\KEISUU\";
            string exlName = localPass + "計数" + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            // 手順3：ファイルを保存します。
            c1XLBook1.Save(exlName + ".xlsx");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;

            //excel出力
            System.Diagnostics.Process.Start(exlName + ".xlsx");

            Com.InHistory("Excel計数を開いた。。", "", "");
        }

        private void dgvtouzenki_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }
        }

        private void dgvkozenki_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;

            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }
        }

        private void dgvkozissekizenki_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //【個別】実績と前期の詳細

            Cursor.Current = Cursors.WaitCursor;
            decimal result;

            string group = ((DataGridView)sender).Rows[((DataGridView)sender).CurrentCell.RowIndex].Cells[0].Value.ToString();

            if (group == "固定売上" | group == "臨時売上" | group == "売上" | group == "人件費" | group == "諸経費" | group == "従業員数")
            {
                if (group == "人件費" && _drv.Row.ItemArray[1].ToString().Substring(1, 4) == "9900")
                {
                    if (Convert.ToInt16(Program.access) < 5) //部門長未満はみれない
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + label2.Text + " " + label4.Text, "");
                        return;
                    }
                    else if (Convert.ToInt16(Program.access) < 9 && _drv.Row.ItemArray[0].ToString() == "11000")
                    {
                        MessageBox.Show("参照権限がありません");
                        Com.InHistory("予算更新画面-権限制限", group + " " + label2.Text + " " + label4.Text, "");
                        return;
                    }
                }

                //数値以外はスルー
                if (decimal.TryParse(((DataGridView)sender).CurrentCell.Value.ToString(), out result))
                {
                    //ゼロはスルー
                    if (result != 0)
                    {
                        //年間の従業員はスルー
                        if (((DataGridView)sender).CurrentCell.OwningColumn.HeaderText == "年間" && group == "従業員数") return;

                        string bumonCD = _drv.Row.ItemArray[0].ToString();
                        string genbaCD = _drv.Row.ItemArray[1].ToString();

                        string ym = ymChange(((DataGridView)sender));

                        string ys = "";
                        string ye = "";

                        //合計の場合はbetween対応
                        if (ym == "年間")
                        {
                            if (((DataGridView)sender).Name == "dgvkozisseki")
                            {
                                ys = maey + "04";
                                ye = atoy + "03";
                            }
                            else
                            { 
                                ys = maeyex + "04"; 
                                ye = atoyex + "03";
                            }
                        }
                        else
                        {
                            ys = ym;
                            ye = ym;
                        }

                        //別フォームで表示
                        DetailsS_KK Detail = new DetailsS_KK(group, ys, ye, bumonCD, genbaCD, label2.Text, label4.Text, cbhikiate.Checked);
                        Detail.Show();
                    }
                }
            }

            //固定売上、臨時売上、物品売上、売上、人件費、諸経費以外はスルー
            //201605 従業員数追加

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void cbhikiate_CheckedChanged(object sender, EventArgs e)
        {
            GetData();

            string dum = tabControl1.SelectedIndex.ToString();

            if (dum == "0")
            {
                //TODO くるしまぎれ
                _drv = null;
                dataGridView1_SelectionChanged(sender, e);
            }
            else
            {
                GetDataTotal();
            }

            if (cbhikiate.Checked)
            {
                button2.Visible = true;
            }
            else
            {
                button2.Visible = false;
            }
        }

        private void dgvkozisseki_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }

            dgvkozisseki.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvkozisseki.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvkozisseki.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvkozisseki.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvkozisseki.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;

            //DataGridView1の左側2列を固定する
            dgvkozisseki.Columns[0].Frozen = true;
        }

        private void dgvkoyosan_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                Com.GetLevelDispPCA(e, val);
            }


            dgvkoyosan.Rows[2].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            dgvkoyosan.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費
            dgvkoyosan.Rows[6].DefaultCellStyle.BackColor = Color.PaleTurquoise;//利益

            //計数表示
            dgvkoyosan.Rows[7].DefaultCellStyle.Format = "0.00\'%\'";

            //ソート禁止設定
            foreach (DataGridViewColumn c in dgvkoyosan.Columns)
                c.SortMode = DataGridViewColumnSortMode.NotSortable;

            //DataGridView1の左側2列を固定する
            dgvkoyosan.Columns[0].Frozen = true;
        }

        private void dgvtouzenki_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //合計前期の諸経費をダブルクリックした場合の処理
            dgvCellDK(sender, e);
        }

        private void dgvCellDK(object sender, DataGridViewCellEventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;
            decimal zerock;

            string group = ((DataGridView)sender).Rows[((DataGridView)sender).CurrentCell.RowIndex].Cells[0].Value.ToString();

            //従業員数
            string count = ((DataGridView)sender).Rows[14].Cells[((DataGridView)sender).CurrentCell.ColumnIndex].Value.ToString();

            if (group == "諸経費" || group == "管理諸経費")
            {
                //数値以外はスルー
                if (decimal.TryParse(((DataGridView)sender).CurrentCell.Value.ToString(), out zerock))
                {
                    //ゼロはスルー
                    if (zerock != 0)
                    {
                        string ym = ymChange(((DataGridView)sender));

                        string ys = "";
                        string ye = "";

                        //合計の場合はbetween対応
                        if (ym == "年間")
                        {
                            if (((DataGridView)sender).Name == "dgvkozisseki" || ((DataGridView)sender).Name == "dgvtouzisseki")
                            {
                                ys = maey + "04";
                                ye = atoy + "03";
                            }
                            else
                            {
                                ys = maeyex + "04";
                                ye = atoyex + "03";
                            }

                            DataTable dtym = Com.GetDB("select max(年月) from dbo.kanrikeisuu where 年月 between '" + ys + "' and '" + ye + "'");

                            ye = dtym.Rows[0][0].ToString();
                        }
                        else
                        {
                            ys = ym;
                            ye = ym;
                        }

                        //別フォームで表示
                        DetailsSKKSum Detail = new DetailsSKKSum(group, ys, ye, result, count);
                        Detail.Show();
                    }
                }


            }

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void dgvtouzisseki_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //合計実績の諸経費をダブルクリックした場合の処理
            dgvCellDK(sender, e);
        }

        private void cbzi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null)
            {
                zi = cbzi.SelectedItem.ToString().Substring(1, 1);
                GetData();
            }
        }

        private void cbki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null)
            {
                IniSet();
            }
        }
    }
}
