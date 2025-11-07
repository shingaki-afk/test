using C1.C1Excel;
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
    public partial class AfterCalc_PCA : Form
    {
        private TargetDays td = new TargetDays();
        //private int kizyun = 0;
        private DataTable dt = new DataTable();
        private DataTable sumdt = new DataTable();

        private string y = "";
        private string m = "";

        public AfterCalc_PCA()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);
            dataGridView2.Font = new Font(dataGridView1.Font.Name, 12);
            dataGridView3.Font = new Font(dataGridView1.Font.Name, 12);
            
            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            for (int i = 2023; i <= Convert.ToInt16(td.StartYMD.AddMonths(1).ToString("yyyy")); i++)
            {
                comboBox1.Items.Add(i.ToString());
            }

            comboBox2.Items.Add("01");
            comboBox2.Items.Add("02");
            comboBox2.Items.Add("03");
            comboBox2.Items.Add("04");
            comboBox2.Items.Add("05");
            comboBox2.Items.Add("06");
            comboBox2.Items.Add("07");
            comboBox2.Items.Add("08");
            comboBox2.Items.Add("09");
            comboBox2.Items.Add("10");
            comboBox2.Items.Add("11");
            comboBox2.Items.Add("12");

            comboBox3.Items.Add("給与集計表");
            comboBox3.Items.Add("給与預かり金");
            comboBox3.Items.Add("給与預かり金_地区別");
            comboBox3.Items.Add("金種表_組織別");
            comboBox3.Items.Add("金種表_個別");
            comboBox3.Items.Add("財形生保控除");
            comboBox3.Items.Add("その他控除");
            comboBox3.Items.Add("友の会費");
            comboBox3.Items.Add("組織別人数");
            
            comboBox3.Items.Add("住民税一覧");
            comboBox3.Items.Add("住民税合計");
            comboBox3.Items.Add("住民税先月比");
            comboBox3.Items.Add("PPP/PFI人件費");
            comboBox3.Items.Add("PPP/PFI人件費_現場別");
            comboBox3.Items.Add("客室人件費");
            comboBox3.Items.Add("80超改善指導書");

            comboBox3.Items.Add("退職者源泉徴収一覧");

            checkedListBox1.Items.Add("本社");
            checkedListBox1.Items.Add("那覇");
            checkedListBox1.Items.Add("八重山");
            checkedListBox1.Items.Add("北部");
            checkedListBox1.Items.Add("広域");
            checkedListBox1.Items.Add("宮古島");
            checkedListBox1.Items.Add("久米島");

            //条件コントロール、消すー
            checkedListBox1.Visible = false;
            tikulbl.Visible = false;
            printbtn.Visible = false;

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }


            comboBox1.SelectedItem = td.StartYMD.AddMonths(1).ToString("yyyy");
            comboBox2.SelectedItem = td.StartYMD.AddMonths(1).ToString("MM");

            DataTable ctdt = new DataTable();
            ctdt = Com.GetDB("select count(*) from QUATRO.dbo.QCMTSHIKYU where 会社コード = 'E0' and 年 = '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "' and 月 = '" + td.StartYMD.AddMonths(1).ToString("MM") + "' and 完了 = '0'");

            if (ctdt.Rows[0][0].ToString() != "0")
            {
                label1.Text = td.StartYMD.AddMonths(1).ToString("yyyy") + "年" + td.StartYMD.AddMonths(1).Month.ToString() + "月支払給与はまだ未確定です。";
                //comboBox2.SelectedItem = td.StartYMD.ToString("MM");
                //kizyun = Convert.ToInt32(td.StartYMD.ToString("yyyyMM"));
            }
            else
            {
                //kizyun = Convert.ToInt32(td.StartYMD.AddMonths(1).ToString("yyyyMM"));
            }

            //出勤簿 TODO 
            comboBox5.Items.Add("2025/01");
            comboBox5.Items.Add("2025/02");
            comboBox5.Items.Add("2025/03");
            comboBox5.Items.Add("2025/04");
            comboBox5.Items.Add("2025/05");
            comboBox5.Items.Add("2025/06");
            comboBox5.Items.Add("2025/07");
            comboBox5.Items.Add("2025/08");
            comboBox5.Items.Add("2025/09");
            comboBox5.Items.Add("2025/10");
            comboBox5.Items.Add("2025/11");
            comboBox5.Items.Add("2025/12");
            comboBox5.Items.Add("2026/01");
            comboBox5.Items.Add("2026/02");
            comboBox5.Items.Add("2026/03");
            comboBox5.SelectedIndex = 0;

            comboBox4.Items.Add("01_現業");
            comboBox4.Items.Add("02_客室");
            comboBox4.Items.Add("03_施設");
            comboBox4.Items.Add("04_エンジ");
            comboBox4.Items.Add("05_PPP/PFI");
            comboBox4.Items.Add("11_北部");
            comboBox4.Items.Add("12_八重山");
            comboBox4.Items.Add("13_広域");
            comboBox4.Items.Add("14_宮古島");
            comboBox4.Items.Add("15_久米島");
            comboBox4.Items.Add("21_役員室");
            comboBox4.Items.Add("22_営業");
            comboBox4.Items.Add("23_経営企画");
            comboBox4.Items.Add("24_総務");
            comboBox4.SelectedIndex = 0;

            Com.InHistory("45_給与計算後資料", "", "");
        }

        private DataTable GetData(string sql)
        {
            DataTable dt = new DataTable();
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        //Cmd.CommandText = "select * from dbo.金種表_那覇職種別";
                        Cmd.CommandText = sql;
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

            return dt;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void GetComboData()
        {
            if (comboBox1.SelectedItem == null || comboBox2.SelectedItem == null || comboBox3.SelectedItem == null) return;

            //コンボボックス無効化・カーソル変更
            comboBox3.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;

            y = comboBox1.SelectedItem.ToString();
            m = comboBox2.SelectedItem.ToString();
            string str = y + "年" + m + "月" + "処理";

            //string tiku = comboBox4.SelectedItem.ToString();
            //string sqltiku = "";

            //switch (tiku)
            //{
            //    case "全地区": sqltiku = ""; break;
            //    case "本社": sqltiku = " where 地区名 = '本社' or 地区名 = '那覇' "; break;
            //    case "八重山": sqltiku = " where 地区名 = '八重山' "; break;
            //    case "北部": sqltiku = " where 地区名 = '北部' "; break;
            //}

            string tiku = " where 地区名 <> 'dummy' ";
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) tiku += " and 地区名 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            DateTime ymd = new DateTime(Convert.ToInt16(y), Convert.ToInt16(m), 01);
            DateTime ymd_ex = ymd.AddMonths(-1);
            DateTime ymdend = ymd.AddDays(-1);


            switch (comboBox3.SelectedItem.ToString())
            {
                case "給与集計表":
                    DataTable wkdt = new DataTable();
                    wkdt = Com.GetDB("select * from dbo.k給与集計表('" + y + "','" + m + "')");
                    dt = Com.replaceDataTable(wkdt);
                    sumdt = null;
                    splitContainer1.SplitterDistance = 400;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "給与預かり金":
                    dt = Com.GetDB("select 項目, 内容, 金額 from dbo.k給与預金一覧取得_地区別廃止('" + y + "','" + m + "') where 金額 is not null order by 項目, 内容");
                    //dt = Com.GetDB("select 項目, 内容, 金額 from dbo.k給与預金一覧取得_地区別廃止('" + y + "','" + m + "') order by 項目, 内容");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 600;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "給与預かり金_地区別":
                    dt = Com.GetDB("select 経理地区, 項目, 内容, 金額 from dbo.k給与預金一覧取得('" + y + "','" + m + "') where 金額 is not null order by 経理地区, 項目, 内容");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 600;
                    printbtn.Visible = true ;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "金種表_組織別":
                    dt = Com.GetDB("select * from dbo.k金種表_組織別('" + y + "','" + m + "')" + tiku);
                    sumdt = Com.GetDB("select 地区名, '-' as 組織名, sum(人数) as [人数], sum(現金支給額) as [現金支給額], sum(万) as [万], sum([5千]) as [5千], sum([千]) as [千], sum([500]) as [500], sum([100]) as [100], sum([50]) as [50], sum([10]) as [10], sum([5]) as [5], sum([1]) as [1] from dbo.k金種表_組織別('" + y + "','" + m + "')" + tiku + " group by 地区名");
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = true;
                    tikulbl.Visible = true;
                    break;
                case "金種表_個別":
                    dt = Com.GetDB("select * from dbo.k金種表('" + y + "','" + m + "')" + tiku);
                    sumdt = Com.GetDB("select '-' as 地区名,'-' as 組織名,'-' as 現場CD,'-' as 現場名,'-' as 社員番号, count(*) as [人数], sum(現金支給額) as [現金支給額], sum(万) as [万], sum([5千]) as [5千], sum([千]) as [千], sum([500]) as [500], sum([100]) as [100], sum([50]) as [50], sum([10]) as [10], sum([5]) as [5], sum([1]) as [1] from dbo.k金種表('" + y + "','" + m + "')");
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = true;
                    tikulbl.Visible = true;
                    break;
                case "財形生保控除":
                    //dt = Com.GetDB("select * from dbo.z財形生保控除取得('" + y + "','" + m + "')" + tiku);
                    dt = Com.GetDB("select * from dbo.z財形生保控除取得('" + y + "','" + m + "') order by 経理地区");
                    //sumdt = Com.GetDB("select '-' as '地区','-' as '社員番号','-' as '漢字氏名' ,'-' as '退職年月日' ,sum(琉銀財形) as '琉銀財形',sum(沖銀財形) as '沖銀財形',sum(財形積立) as '財形積立',sum(日本生命) as '日本生命',sum(住友生命) as '住友生命',sum(日動火災) as '日動火災',sum(ＡＦＬＡＣ) as 'ＡＦＬＡＣ',sum(団体傷害) as '団体傷害',sum(アライアンス) as 'アライアンス',sum(朝日生命) as '朝日生命',sum(フコク生命) as 'フコク生命',sum(ＴＯＰ２１) as 'ＴＯＰ２１',sum(生命保険合計) as '生命保険合計' from dbo.z財形生保控除取得('" + y + "','" + m + "')" + tiku);
                    sumdt = Com.GetDB("select 経理地区,'-' as '社員番号','-' as '漢字氏名' ,'-' as '退職年月日' ,sum(琉銀財形) as '琉銀財形',sum(沖銀財形) as '沖銀財形',sum(日本生命) as '日本生命',sum(住友生命) as '住友生命',sum(日動火災) as '日動火災',sum(ＡＦＬＡＣ) as 'ＡＦＬＡＣ',sum(団体傷害) as '団体傷害',sum(アライアンス) as 'アライアンス',sum(朝日生命) as '朝日生命',sum(フコク生命) as 'フコク生命',sum(ＴＯＰ２１) as 'ＴＯＰ２１' from dbo.z財形生保控除取得('" + y + "','" + m + "') group by 経理地区 order by 経理地区");
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "その他控除":
                    dt = Com.GetDB("select * from dbo.sその他控除取得('" + y + "','" + m + "')" + tiku);
                    sumdt = Com.GetDB("select '-' as 地区名, '-' as 組織名, '-' as 社員番号, '-' as 漢字氏名, sum(積立金) as 積立金, sum([前払金(-)]) as [前払金(-)], sum([固定他１]) as [固定他１], sum([固定他２]) as [固定他２], sum([変動他１]) as [変動他１], sum([変動他２]) as [変動他２], sum(差押金) as 差押金, sum(退職積立) as 退職積立, '-' as [固1内容], '-' as [固2内容], '-' as [変1内容], '-' as [変2内容] from dbo.sその他控除取得('" + y + "','" + m + "')" + tiku);
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = true;
                    tikulbl.Visible = true;
                    break;
                case "友の会費":
                    dt = Com.GetDB("select * from dbo.t友の会費('" + y + "','" + m + "','" + ymdend + "') order by 地区名 desc");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "組織別人数":
                    dt = Com.GetDB("select * from dbo.s組織別人数('" + y + "','" + m + "','" + ymdend + "') order by 組織CD");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "住民税一覧":
                    dt = Com.GetDB("select * from dbo.住民税一覧取得('" + y + "','" + m + "')" + tiku + " order by 納付先コード, 地区名");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = true;
                    tikulbl.Visible = true;
                    break;
                case "住民税合計":
                    dt = Com.GetDB("select 納付先コード, 納付先名, sum(isnull(住民税, 0) + isnull(退職一括徴収, 0)) as 住民税合計 from dbo.住民税一覧取得('" + y + "','" + m + "') " + tiku + " group by 納付先コード, 納付先名 order by 納付先コード");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = true;
                    checkedListBox1.Visible = true;
                    break;
                case "住民税先月比":
                    dt = Com.GetDB("select * from dbo.住民税前月比('" + y + "', '" + m + "', '" + ymd_ex.ToString("yyyy") + "', '" + ymd_ex.ToString("MM") + "', '" + ymd.AddDays(-1).ToString("yyyy/MM/dd") + "') " + tiku + " order by 退職年月日, 納付先コード");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = true;
                    tikulbl.Visible = true;
                    break;
                case "PPP/PFI人件費":
                    dt = Com.GetDB("select * from dbo.PPPPFI人件費('" + y + "','" + m + "', '" + ymdend.ToString("yyyy/MM/dd") + "') order by 組織CD, 現場CD");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "PPP/PFI人件費_現場別":
                    dt = Com.GetDB("select * from dbo.PPPPFI人件費_現場別('" + y + "','" + m + "', '" + ymdend.ToString("yyyy/MM/dd") + "') order by 現場CD");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "客室人件費":
                    dt = Com.GetDB("select * from dbo.k客室人件費('" + y + "','" + m + "', '" + ymdend.ToString("yyyy/MM/dd") + "') order by 組織CD, 現場CD");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "80超改善指導書":
                    dt = Com.GetDB("select * from dbo.z残業80超改善指導書('" + y + "','" + m + "', '" + ymdend.ToString("yyyy/MM/dd") + "') ");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = false;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
                case "退職者源泉徴収一覧":
                    dt = Com.GetDB("select top 1 * from dbo.t退職者源泉徴収一覧('" + ymd_ex.ToString("yyyy/MM/dd") + "','" + ymdend.ToString("yyyy/MM/dd") + "') order by 組織名, 社員番号 ");
                    sumdt = null;
                    splitContainer1.SplitterDistance = 1800;
                    printbtn.Visible = true;
                    checkedListBox1.Visible = false;
                    tikulbl.Visible = false;
                    break;
            }

            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.DataSource = dt;

            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView2.DataSource = sumdt;

            DispChange();

            //全て入力した後に列幅を自動調節する
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            comboBox3.Enabled = true;

            Com.InHistory(comboBox3.SelectedItem.ToString(), str, "");
        }

        private void DispChange()
        {
            if (dt.Rows.Count == 0) return;

            if (comboBox3.SelectedItem.ToString() == "金種表_組織別")
            {
                dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox3.SelectedItem.ToString() == "金種表_個別")
            {
                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox3.SelectedItem.ToString() == "財形生保控除")
            {
                //三桁区切り表示
                for (int i = 3; i < 15; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView2.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "その他控除")
            {
                //三桁区切り表示
                for (int i = 4; i < 12; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView2.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "友の会費")
            {
                //三桁区切り表示
                for (int i = 0; i < 3; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "住民税一覧")
            {
                dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            }
            else if (comboBox3.SelectedItem.ToString() == "住民税合計")
            {
                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            }
            else if (comboBox3.SelectedItem.ToString() == "住民税先月比")
            {
                //三桁区切り表示
                for (int i = 6; i < 10; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "PPP/PFI人件費")
            {
                for (int i = 7; i <= 30; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "PPP/PFI人件費_現場別")
            {
                for (int i = 2; i <= 5; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "給与集計表")
            {
                dataGridView1.Columns[0].Width = 150;

                dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox3.SelectedItem.ToString() == "給与預かり金")
            {
                dataGridView1.Columns[0].Width = 80;
                dataGridView1.Columns[1].Width = 200;
                dataGridView1.Columns[2].Width = 80;
                //dataGridView1.Columns[3].Width = 80;

                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox3.SelectedItem.ToString() == "給与預かり金_地区別")
            {
                dataGridView1.Columns[0].Width = 80;
                dataGridView1.Columns[1].Width = 80;
                dataGridView1.Columns[2].Width = 200;
                dataGridView1.Columns[3].Width = 80;

                dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.DataSource = null;

            string sql = "";
            sql += "select * from(select ";
            sql += "a.社員番号, a.地区名, a.組織名, a.現場名, a.漢字氏名, a.時給, a.労働時間, a.週労働数, 健保, 年齢, a.入社年月日, a.在籍状況, ";
            sql += "isnull(d.SYSTR00001_支給合計額, 0) - isnull(d.A8500CG_通勤課税, 0) - isnull(d.A3700CG_通勤非課税, 0) - isnull(d.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-11).ToString("MM") + "月], ";
            sql += "isnull(f.SYSTR00001_支給合計額, 0) - isnull(f.A8500CG_通勤課税, 0) - isnull(f.A3700CG_通勤非課税, 0) - isnull(f.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-10).ToString("MM") + "月], ";
            sql += "isnull(h.SYSTR00001_支給合計額, 0) - isnull(h.A8500CG_通勤課税, 0) - isnull(h.A3700CG_通勤非課税, 0) - isnull(h.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-9).ToString("MM") + "月], ";
            sql += "isnull(j.SYSTR00001_支給合計額, 0) - isnull(j.A8500CG_通勤課税, 0) - isnull(j.A3700CG_通勤非課税, 0) - isnull(j.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-8).ToString("MM") + "月], ";
            sql += "isnull(l.SYSTR00001_支給合計額, 0) - isnull(l.A8500CG_通勤課税, 0) - isnull(l.A3700CG_通勤非課税, 0) - isnull(l.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-7).ToString("MM") + "月], ";
            sql += "isnull(n.SYSTR00001_支給合計額, 0) - isnull(n.A8500CG_通勤課税, 0) - isnull(n.A3700CG_通勤非課税, 0) - isnull(n.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-6).ToString("MM") + "月], ";
            sql += "isnull(p.SYSTR00001_支給合計額, 0) - isnull(p.A8500CG_通勤課税, 0) - isnull(p.A3700CG_通勤非課税, 0) - isnull(p.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-5).ToString("MM") + "月], ";
            sql += "isnull(r.SYSTR00001_支給合計額, 0) - isnull(r.A8500CG_通勤課税, 0) - isnull(r.A3700CG_通勤非課税, 0) - isnull(r.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-4).ToString("MM") + "月], ";
            sql += "isnull(t.SYSTR00001_支給合計額, 0) - isnull(t.A8500CG_通勤課税, 0) - isnull(t.A3700CG_通勤非課税, 0) - isnull(t.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-3).ToString("MM") + "月], ";
            sql += "isnull(v.SYSTR00001_支給合計額, 0) - isnull(v.A8500CG_通勤課税, 0) - isnull(v.A3700CG_通勤非課税, 0) - isnull(v.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-2).ToString("MM") + "月], ";
            sql += "isnull(x.SYSTR00001_支給合計額, 0) - isnull(x.A8500CG_通勤課税, 0) - isnull(x.A3700CG_通勤非課税, 0) - isnull(x.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(-1).ToString("MM") + "月], ";
            sql += "isnull(b.SYSTR00001_支給合計額, 0) - isnull(b.A8500CG_通勤課税, 0) - isnull(b.A3700CG_通勤非課税, 0) - isnull(b.A7100CG_深夜手当, 0) as [" + td.StartYMD.AddMonths(0).ToString("MM") + "月] ";
            sql += "from dbo.accessNew a ";
            sql += "left ";
            sql += "join dbo.KYTTKREKR4_E0 d on a.社員番号 = d.社員番号 and d.会社コード = 'E0' and d.処理年 = '" + td.StartYMD.AddMonths(-10).ToString("yyyy") + "' and d.処理月 = '" + td.StartYMD.AddMonths(-10).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 f on a.社員番号 = f.社員番号 and f.会社コード = 'E0' and f.処理年 = '" + td.StartYMD.AddMonths(-9).ToString("yyyy") + "' and f.処理月 = '" + td.StartYMD.AddMonths(-9).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 h on a.社員番号 = h.社員番号 and h.会社コード = 'E0' and h.処理年 = '" + td.StartYMD.AddMonths(-8).ToString("yyyy") + "' and h.処理月 = '" + td.StartYMD.AddMonths(-8).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 j on a.社員番号 = j.社員番号 and j.会社コード = 'E0' and j.処理年 = '" + td.StartYMD.AddMonths(-7).ToString("yyyy") + "' and j.処理月 = '" + td.StartYMD.AddMonths(-7).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 l on a.社員番号 = l.社員番号 and l.会社コード = 'E0' and l.処理年 = '" + td.StartYMD.AddMonths(-6).ToString("yyyy") + "' and l.処理月 = '" + td.StartYMD.AddMonths(-6).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 n on a.社員番号 = n.社員番号 and n.会社コード = 'E0' and n.処理年 = '" + td.StartYMD.AddMonths(-5).ToString("yyyy") + "' and n.処理月 = '" + td.StartYMD.AddMonths(-5).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 p on a.社員番号 = p.社員番号 and p.会社コード = 'E0' and p.処理年 = '" + td.StartYMD.AddMonths(-4).ToString("yyyy") + "' and p.処理月 = '" + td.StartYMD.AddMonths(-4).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 r on a.社員番号 = r.社員番号 and r.会社コード = 'E0' and r.処理年 = '" + td.StartYMD.AddMonths(-3).ToString("yyyy") + "' and r.処理月 = '" + td.StartYMD.AddMonths(-3).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 t on a.社員番号 = t.社員番号 and t.会社コード = 'E0' and t.処理年 = '" + td.StartYMD.AddMonths(-2).ToString("yyyy") + "' and t.処理月 = '" + td.StartYMD.AddMonths(-2).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 v on a.社員番号 = v.社員番号 and v.会社コード = 'E0' and v.処理年 = '" + td.StartYMD.AddMonths(-1).ToString("yyyy") + "' and v.処理月 = '" + td.StartYMD.AddMonths(-1).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 x on a.社員番号 = x.社員番号 and x.会社コード = 'E0' and x.処理年 = '" + td.StartYMD.AddMonths(0).ToString("yyyy") + "' and x.処理月 = '" + td.StartYMD.AddMonths(0).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 b on a.社員番号 = b.社員番号 and b.会社コード = 'E0' and b.処理年 = '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "' and b.処理月 = '" + td.StartYMD.AddMonths(1).ToString("MM") + "' ";
            sql += "where a.給与支給区分 in ('E1', 'F1') ";
            sql += ") temp where ([01月] + [02月] + [03月] + [04月] + [05月] + [06月] + [07月] + [08月] + [09月] + [10月] + [11月] + [12月]) / 12 > 0 ";

            dataGridView1.DataSource = Com.GetDB(sql);

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.DataSource = null;

            string strc = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-11).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string stre = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-10).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strg = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-9).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string stri = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-8).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strk = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-7).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strm = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-6).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string stro = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-5).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strq = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-4).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strs = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-3).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string stru = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-2).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string strw = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(-1).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string stry = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(0).ToString("yyyyMM") + "'").Rows[0][0].ToString();

            string sql = "";
            sql += "select * from ( ";
            sql += "select a.社員番号, a.地区名, a.組織名, a.現場名, a.給与支給区分名 as 社員区分, a.役職名, a.氏名, ";
            
            sql += " case when b.給与支給区分 is null then 0 when b.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(c.F0100_残業時間 + c.F0300_法休時間 + c.F0400_所休時間 + c.F0500_延長時間 - (c.F1600_届欠 + c.F1700_無届) * b.A5700CG_勤務時間 - c.F0700_遅刻時間, 0) else ";
            sql += "isnull(c.F0100_残業時間 + c.F0300_法休時間 + c.F0400_所休時間 + c.F0500_延長時間 + c.F0800_所定 * b.A5700CG_勤務時間 - c.F0700_遅刻時間, 0) - " + strc + " end as [" + td.StartYMD.AddMonths(-11).ToString("MM") + "月(" + strc + ")],  ";

            sql += " case when d.給与支給区分 is null then 0  when d.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(e.F0100_残業時間 + e.F0300_法休時間 + e.F0400_所休時間 + e.F0500_延長時間 - (e.F1600_届欠 + e.F1700_無届) * d.A5700CG_勤務時間 - e.F0700_遅刻時間, 0) else ";
            sql += "isnull(e.F0100_残業時間 + e.F0300_法休時間 + e.F0400_所休時間 + e.F0500_延長時間 + e.F0800_所定 * d.A5700CG_勤務時間 - e.F0700_遅刻時間, 0) - " + stre + " end as [" + td.StartYMD.AddMonths(-10).ToString("MM") + "月(" + stre + ")],  ";

            sql += " case when f.給与支給区分 is null then 0  when f.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(g.F0100_残業時間 + g.F0300_法休時間 + g.F0400_所休時間 + g.F0500_延長時間 - (g.F1600_届欠 + g.F1700_無届) * f.A5700CG_勤務時間 - g.F0700_遅刻時間, 0) else ";
            sql += "isnull(g.F0100_残業時間 + g.F0300_法休時間 + g.F0400_所休時間 + g.F0500_延長時間 + g.F0800_所定 * f.A5700CG_勤務時間 - g.F0700_遅刻時間, 0) - " + strg + " end as [" + td.StartYMD.AddMonths(-9).ToString("MM") + "月(" + strg + ")],  ";

            sql += " case when h.給与支給区分 is null then 0  when h.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(i.F0100_残業時間 + i.F0300_法休時間 + i.F0400_所休時間 + i.F0500_延長時間 - (i.F1600_届欠 + i.F1700_無届) * h.A5700CG_勤務時間 - i.F0700_遅刻時間, 0) else ";
            sql += "isnull(i.F0100_残業時間 + i.F0300_法休時間 + i.F0400_所休時間 + i.F0500_延長時間 + i.F0800_所定 * h.A5700CG_勤務時間 - i.F0700_遅刻時間, 0) - " + stri + " end as [" + td.StartYMD.AddMonths(-8).ToString("MM") + "月(" + stri + ")],  ";

            sql += " case when j.給与支給区分 is null then 0  when j.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(k.F0100_残業時間 + k.F0300_法休時間 + k.F0400_所休時間 + k.F0500_延長時間 - (k.F1600_届欠 + k.F1700_無届) * j.A5700CG_勤務時間 - k.F0700_遅刻時間, 0) else ";
            sql += "isnull(k.F0100_残業時間 + k.F0300_法休時間 + k.F0400_所休時間 + k.F0500_延長時間 + k.F0800_所定 * j.A5700CG_勤務時間 - k.F0700_遅刻時間, 0) - " + strk + " end as [" + td.StartYMD.AddMonths(-7).ToString("MM") + "月(" + strk + ")],  ";

            sql += " case when l.給与支給区分 is null then 0  when l.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(m.F0100_残業時間 + m.F0300_法休時間 + m.F0400_所休時間 + m.F0500_延長時間 - (m.F1600_届欠 + m.F1700_無届) * l.A5700CG_勤務時間 - m.F0700_遅刻時間, 0) else ";
            sql += "isnull(m.F0100_残業時間 + m.F0300_法休時間 + m.F0400_所休時間 + m.F0500_延長時間 + m.F0800_所定 * l.A5700CG_勤務時間 - m.F0700_遅刻時間, 0) - " + strm + " end as [" + td.StartYMD.AddMonths(-6).ToString("MM") + "月(" + strm + ")],  ";

            sql += " case when n.給与支給区分 is null then 0  when n.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(o.F0100_残業時間 + o.F0300_法休時間 + o.F0400_所休時間 + o.F0500_延長時間 - (o.F1600_届欠 + o.F1700_無届) * n.A5700CG_勤務時間 - o.F0700_遅刻時間, 0) else ";
            sql += "isnull(o.F0100_残業時間 + o.F0300_法休時間 + o.F0400_所休時間 + o.F0500_延長時間 + o.F0800_所定 * n.A5700CG_勤務時間 - o.F0700_遅刻時間, 0) - " + stro + " end as [" + td.StartYMD.AddMonths(-5).ToString("MM") + "月(" + stro + ")],  ";

            sql += " case when p.給与支給区分 is null then 0  when p.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(q.F0100_残業時間 + q.F0300_法休時間 + q.F0400_所休時間 + q.F0500_延長時間 - (q.F1600_届欠 + q.F1700_無届) * p.A5700CG_勤務時間 - q.F0700_遅刻時間, 0) else ";
            sql += "isnull(q.F0100_残業時間 + q.F0300_法休時間 + q.F0400_所休時間 + q.F0500_延長時間 + q.F0800_所定 * p.A5700CG_勤務時間 - q.F0700_遅刻時間, 0) - " + strq + " end as [" + td.StartYMD.AddMonths(-4).ToString("MM") + "月(" + strq + ")],  ";

            sql += " case when r.給与支給区分 is null then 0  when r.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(s.F0100_残業時間 + s.F0300_法休時間 + s.F0400_所休時間 + s.F0500_延長時間 - (s.F1600_届欠 + s.F1700_無届) * r.A5700CG_勤務時間 - s.F0700_遅刻時間, 0) else ";
            sql += "isnull(s.F0100_残業時間 + s.F0300_法休時間 + s.F0400_所休時間 + s.F0500_延長時間 + s.F0800_所定 * r.A5700CG_勤務時間 - s.F0700_遅刻時間, 0) - " + strs + " end as [" + td.StartYMD.AddMonths(-3).ToString("MM") + "月(" + strs + ")],  ";

            sql += " case when t.給与支給区分 is null then 0  when t.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(u.F0100_残業時間 + u.F0300_法休時間 + u.F0400_所休時間 + u.F0500_延長時間 - (u.F1600_届欠 + u.F1700_無届) * t.A5700CG_勤務時間 - u.F0700_遅刻時間, 0) else ";
            sql += "isnull(u.F0100_残業時間 + u.F0300_法休時間 + u.F0400_所休時間 + u.F0500_延長時間 + u.F0800_所定 * t.A5700CG_勤務時間 - u.F0700_遅刻時間, 0) - " + stru + " end as [" + td.StartYMD.AddMonths(-2).ToString("MM") + "月(" + stru + ")],  ";

            sql += " case when v.給与支給区分 is null then 0  when v.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(w.F0100_残業時間 + w.F0300_法休時間 + w.F0400_所休時間 + w.F0500_延長時間 - (w.F1600_届欠 + w.F1700_無届) * v.A5700CG_勤務時間 - w.F0700_遅刻時間, 0) else ";
            sql += "isnull(w.F0100_残業時間 + w.F0300_法休時間 + w.F0400_所休時間 + w.F0500_延長時間 + w.F0800_所定 * v.A5700CG_勤務時間 - w.F0700_遅刻時間, 0) - " + strw + " end as [" + td.StartYMD.AddMonths(-1).ToString("MM") + "月(" + strw + ")],  ";
  
            sql += " case when x.給与支給区分 is null then 0  when x.給与支給区分 in ('A1','B1','C1') then ";
            sql += "isnull(y.F0100_残業時間 + y.F0300_法休時間 + y.F0400_所休時間 + y.F0500_延長時間 - (y.F1600_届欠 + y.F1700_無届) * x.A5700CG_勤務時間 - y.F0700_遅刻時間, 0) else ";
            sql += "isnull(y.F0100_残業時間 + y.F0300_法休時間 + y.F0400_所休時間 + y.F0500_延長時間 + y.F0800_所定 * x.A5700CG_勤務時間 - y.F0700_遅刻時間, 0) - " + stry + " end as [" + td.StartYMD.AddMonths(0).ToString("MM") + "月(" + stry + ")] ";

            sql += "from dbo.s社員基本情報_期間指定('" + td.StartYMD.AddMonths(0).ToString("yyyy/MM/dd") + "') a left ";
            sql += "join QUATRO.dbo.KYTTKREKR4_E0 x on x.会社コード = 'E0' and a.社員番号 = x.社員番号 left ";
            sql += "join QUATRO.dbo.KYTTKINTR3_E0 y on y.会社コード = 'E0' and a.社員番号 = y.社員番号 ";
            sql += "left join dbo.KYTTKREKR4_E0 b on a.社員番号 = b.社員番号 and b.会社コード = 'E0' and b.処理年 = '" + td.StartYMD.AddMonths(-10).ToString("yyyy") + "' and b.処理月 = '" + td.StartYMD.AddMonths(-10).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 c on a.社員番号 = c.社員番号 and c.会社コード = 'E0' and c.処理年 = '" + td.StartYMD.AddMonths(-10).ToString("yyyy") + "' and c.処理月 = '" + td.StartYMD.AddMonths(-10).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 d on a.社員番号 = d.社員番号 and d.会社コード = 'E0' and d.処理年 = '" + td.StartYMD.AddMonths(-9).ToString("yyyy") + "' and d.処理月 = '" + td.StartYMD.AddMonths(-9).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 e on a.社員番号 = e.社員番号 and e.会社コード = 'E0' and e.処理年 = '" + td.StartYMD.AddMonths(-9).ToString("yyyy") + "' and e.処理月 = '" + td.StartYMD.AddMonths(-9).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 f on a.社員番号 = f.社員番号 and f.会社コード = 'E0' and f.処理年 = '" + td.StartYMD.AddMonths(-8).ToString("yyyy") + "' and f.処理月 = '" + td.StartYMD.AddMonths(-8).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 g on a.社員番号 = g.社員番号 and g.会社コード = 'E0' and g.処理年 = '" + td.StartYMD.AddMonths(-8).ToString("yyyy") + "' and g.処理月 = '" + td.StartYMD.AddMonths(-8).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 h on a.社員番号 = h.社員番号 and h.会社コード = 'E0' and h.処理年 = '" + td.StartYMD.AddMonths(-7).ToString("yyyy") + "' and h.処理月 = '" + td.StartYMD.AddMonths(-7).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 i on a.社員番号 = i.社員番号 and i.会社コード = 'E0' and i.処理年 = '" + td.StartYMD.AddMonths(-7).ToString("yyyy") + "' and i.処理月 = '" + td.StartYMD.AddMonths(-7).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 j on a.社員番号 = j.社員番号 and j.会社コード = 'E0' and j.処理年 = '" + td.StartYMD.AddMonths(-6).ToString("yyyy") + "' and j.処理月 = '" + td.StartYMD.AddMonths(-6).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 k on a.社員番号 = k.社員番号 and k.会社コード = 'E0' and k.処理年 = '" + td.StartYMD.AddMonths(-6).ToString("yyyy") + "' and k.処理月 = '" + td.StartYMD.AddMonths(-6).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 l on a.社員番号 = l.社員番号 and l.会社コード = 'E0' and l.処理年 = '" + td.StartYMD.AddMonths(-5).ToString("yyyy") + "' and l.処理月 = '" + td.StartYMD.AddMonths(-5).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 m on a.社員番号 = m.社員番号 and m.会社コード = 'E0' and m.処理年 = '" + td.StartYMD.AddMonths(-5).ToString("yyyy") + "' and m.処理月 = '" + td.StartYMD.AddMonths(-5).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 n on a.社員番号 = n.社員番号 and n.会社コード = 'E0' and n.処理年 = '" + td.StartYMD.AddMonths(-4).ToString("yyyy") + "' and n.処理月 = '" + td.StartYMD.AddMonths(-4).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 o on a.社員番号 = o.社員番号 and o.会社コード = 'E0' and o.処理年 = '" + td.StartYMD.AddMonths(-4).ToString("yyyy") + "' and o.処理月 = '" + td.StartYMD.AddMonths(-4).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 p on a.社員番号 = p.社員番号 and p.会社コード = 'E0' and p.処理年 = '" + td.StartYMD.AddMonths(-3).ToString("yyyy") + "' and p.処理月 = '" + td.StartYMD.AddMonths(-3).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 q on a.社員番号 = q.社員番号 and q.会社コード = 'E0' and q.処理年 = '" + td.StartYMD.AddMonths(-3).ToString("yyyy") + "' and q.処理月 = '" + td.StartYMD.AddMonths(-3).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 r on a.社員番号 = r.社員番号 and r.会社コード = 'E0' and r.処理年 = '" + td.StartYMD.AddMonths(-2).ToString("yyyy") + "' and r.処理月 = '" + td.StartYMD.AddMonths(-2).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 s on a.社員番号 = s.社員番号 and s.会社コード = 'E0' and s.処理年 = '" + td.StartYMD.AddMonths(-2).ToString("yyyy") + "' and s.処理月 = '" + td.StartYMD.AddMonths(-2).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 t on a.社員番号 = t.社員番号 and t.会社コード = 'E0' and t.処理年 = '" + td.StartYMD.AddMonths(-1).ToString("yyyy") + "' and t.処理月 = '" + td.StartYMD.AddMonths(-1).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 u on a.社員番号 = u.社員番号 and u.会社コード = 'E0' and u.処理年 = '" + td.StartYMD.AddMonths(-1).ToString("yyyy") + "' and u.処理月 = '" + td.StartYMD.AddMonths(-1).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKREKR4_E0 v on a.社員番号 = v.社員番号 and v.会社コード = 'E0' and v.処理年 = '" + td.StartYMD.AddMonths(0).ToString("yyyy") + "' and v.処理月 = '" + td.StartYMD.AddMonths(0).ToString("MM") + "' ";
            sql += "left join dbo.KYTTKINTR3_E0 w on a.社員番号 = w.社員番号 and w.会社コード = 'E0' and w.処理年 = '" + td.StartYMD.AddMonths(0).ToString("yyyy") + "' and w.処理月 = '" + td.StartYMD.AddMonths(0).ToString("MM") + "' ";
            //sql += "where a.在籍区分 <> '9' and a.給与支給区分 in ('A1','B1','C1') ";
            sql += "where a.在籍区分 <> '9' ";

            //sql += "union all ";

            //sql += "select a.社員番号, a.地区名, a.組織名, a.現場名, a.給与支給区分名 as 社員区分, a.役職名, a.氏名, ";
            //sql += "isnull(c.F0100_残業時間 + c.F0300_法休時間 + c.F0400_所休時間 + c.F0500_延長時間 + c.F0800_所定 * b.A5700CG_勤務時間 - c.F0700_遅刻時間, 0) - " + strc + " as [" + td.StartYMD.AddMonths(-11).ToString("MM") + "月(" + strc + ")],  ";
            //sql += "isnull(e.F0100_残業時間 + e.F0300_法休時間 + e.F0400_所休時間 + e.F0500_延長時間 + e.F0800_所定 * d.A5700CG_勤務時間 - e.F0700_遅刻時間, 0) - " + stre + " as [" + td.StartYMD.AddMonths(-10).ToString("MM") + "月(" + stre + ")],  ";
            //sql += "isnull(g.F0100_残業時間 + g.F0300_法休時間 + g.F0400_所休時間 + g.F0500_延長時間 + g.F0800_所定 * f.A5700CG_勤務時間 - g.F0700_遅刻時間, 0) - " + strg + " as [" + td.StartYMD.AddMonths(-9).ToString("MM") + "月(" + strg + ")],  ";
            //sql += "isnull(i.F0100_残業時間 + i.F0300_法休時間 + i.F0400_所休時間 + i.F0500_延長時間 + i.F0800_所定 * h.A5700CG_勤務時間 - i.F0700_遅刻時間, 0) - " + stri + " as [" + td.StartYMD.AddMonths(-8).ToString("MM") + "月(" + stri + ")],  ";
            //sql += "isnull(k.F0100_残業時間 + k.F0300_法休時間 + k.F0400_所休時間 + k.F0500_延長時間 + k.F0800_所定 * j.A5700CG_勤務時間 - k.F0700_遅刻時間, 0) - " + strk + " as [" + td.StartYMD.AddMonths(-7).ToString("MM") + "月(" + strk + ")],  ";
            //sql += "isnull(m.F0100_残業時間 + m.F0300_法休時間 + m.F0400_所休時間 + m.F0500_延長時間 + m.F0800_所定 * l.A5700CG_勤務時間 - m.F0700_遅刻時間, 0) - " + strm + " as [" + td.StartYMD.AddMonths(-6).ToString("MM") + "月(" + strm + ")],  ";
            //sql += "isnull(o.F0100_残業時間 + o.F0300_法休時間 + o.F0400_所休時間 + o.F0500_延長時間 + o.F0800_所定 * n.A5700CG_勤務時間 - o.F0700_遅刻時間, 0) - " + stro + " as [" + td.StartYMD.AddMonths(-5).ToString("MM") + "月(" + stro + ")],  ";
            //sql += "isnull(q.F0100_残業時間 + q.F0300_法休時間 + q.F0400_所休時間 + q.F0500_延長時間 + q.F0800_所定 * p.A5700CG_勤務時間 - q.F0700_遅刻時間, 0) - " + strq + " as [" + td.StartYMD.AddMonths(-4).ToString("MM") + "月(" + strq + ")],  ";
            //sql += "isnull(s.F0100_残業時間 + s.F0300_法休時間 + s.F0400_所休時間 + s.F0500_延長時間 + s.F0800_所定 * r.A5700CG_勤務時間 - s.F0700_遅刻時間, 0) - " + strs + " as [" + td.StartYMD.AddMonths(-3).ToString("MM") + "月(" + strs + ")],  ";
            //sql += "isnull(u.F0100_残業時間 + u.F0300_法休時間 + u.F0400_所休時間 + u.F0500_延長時間 + u.F0800_所定 * t.A5700CG_勤務時間 - u.F0700_遅刻時間, 0) - " + stru + " as [" + td.StartYMD.AddMonths(-2).ToString("MM") + "月(" + stru + ")],  ";
            //sql += "isnull(w.F0100_残業時間 + w.F0300_法休時間 + w.F0400_所休時間 + w.F0500_延長時間 + w.F0800_所定 * v.A5700CG_勤務時間 - w.F0700_遅刻時間, 0) - " + strw + " as [" + td.StartYMD.AddMonths(-1).ToString("MM") + "月(" + strw + ")],  ";
            //sql += "isnull(y.F0100_残業時間 + y.F0300_法休時間 + y.F0400_所休時間 + y.F0500_延長時間 + y.F0800_所定 * x.A5700CG_勤務時間 - y.F0700_遅刻時間, 0) - " + stry + " as [" + td.StartYMD.AddMonths(0).ToString("MM") + "月(" + stry + ")] ";
            //sql += "from dbo.s社員基本情報_期間指定('" + td.StartYMD.AddMonths(0).ToString("yyyy/MM/dd") + "') a ";
            //sql += "left join QUATRO.dbo.KYTTKREKR4_E0 x on x.会社コード = 'E0' and x.会社コード = 'E0' and a.社員番号 = x.社員番号 ";
            //sql += "left join QUATRO.dbo.KYTTKINTR3_E0 y on y.会社コード = 'E0' and y.会社コード = 'E0' and a.社員番号 = y.社員番号 ";
            //sql += "left join dbo.KYTTKREKR4_E0 b on a.社員番号 = b.社員番号 and b.会社コード = 'E0' and b.処理年 = '" + td.StartYMD.AddMonths(-10).ToString("yyyy") + "' and b.処理月 = '" + td.StartYMD.AddMonths(-10).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 c on a.社員番号 = c.社員番号 and c.会社コード = 'E0' and c.処理年 = '" + td.StartYMD.AddMonths(-10).ToString("yyyy") + "' and c.処理月 = '" + td.StartYMD.AddMonths(-10).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 d on a.社員番号 = d.社員番号 and d.会社コード = 'E0' and d.処理年 = '" + td.StartYMD.AddMonths(-9).ToString("yyyy") + "' and d.処理月 = '" + td.StartYMD.AddMonths(-9).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 e on a.社員番号 = e.社員番号 and e.会社コード = 'E0' and e.処理年 = '" + td.StartYMD.AddMonths(-9).ToString("yyyy") + "' and e.処理月 = '" + td.StartYMD.AddMonths(-9).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 f on a.社員番号 = f.社員番号 and f.会社コード = 'E0' and f.処理年 = '" + td.StartYMD.AddMonths(-8).ToString("yyyy") + "' and f.処理月 = '" + td.StartYMD.AddMonths(-8).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 g on a.社員番号 = g.社員番号 and g.会社コード = 'E0' and g.処理年 = '" + td.StartYMD.AddMonths(-8).ToString("yyyy") + "' and g.処理月 = '" + td.StartYMD.AddMonths(-8).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 h on a.社員番号 = h.社員番号 and h.会社コード = 'E0' and h.処理年 = '" + td.StartYMD.AddMonths(-7).ToString("yyyy") + "' and h.処理月 = '" + td.StartYMD.AddMonths(-7).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 i on a.社員番号 = i.社員番号 and i.会社コード = 'E0' and i.処理年 = '" + td.StartYMD.AddMonths(-7).ToString("yyyy") + "' and i.処理月 = '" + td.StartYMD.AddMonths(-7).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 j on a.社員番号 = j.社員番号 and j.会社コード = 'E0' and j.処理年 = '" + td.StartYMD.AddMonths(-6).ToString("yyyy") + "' and j.処理月 = '" + td.StartYMD.AddMonths(-6).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 k on a.社員番号 = k.社員番号 and k.会社コード = 'E0' and k.処理年 = '" + td.StartYMD.AddMonths(-6).ToString("yyyy") + "' and k.処理月 = '" + td.StartYMD.AddMonths(-6).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 l on a.社員番号 = l.社員番号 and l.会社コード = 'E0' and l.処理年 = '" + td.StartYMD.AddMonths(-5).ToString("yyyy") + "' and l.処理月 = '" + td.StartYMD.AddMonths(-5).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 m on a.社員番号 = m.社員番号 and m.会社コード = 'E0' and m.処理年 = '" + td.StartYMD.AddMonths(-5).ToString("yyyy") + "' and m.処理月 = '" + td.StartYMD.AddMonths(-5).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 n on a.社員番号 = n.社員番号 and n.会社コード = 'E0' and n.処理年 = '" + td.StartYMD.AddMonths(-4).ToString("yyyy") + "' and n.処理月 = '" + td.StartYMD.AddMonths(-4).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 o on a.社員番号 = o.社員番号 and o.会社コード = 'E0' and o.処理年 = '" + td.StartYMD.AddMonths(-4).ToString("yyyy") + "' and o.処理月 = '" + td.StartYMD.AddMonths(-4).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 p on a.社員番号 = p.社員番号 and p.会社コード = 'E0' and p.処理年 = '" + td.StartYMD.AddMonths(-3).ToString("yyyy") + "' and p.処理月 = '" + td.StartYMD.AddMonths(-3).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 q on a.社員番号 = q.社員番号 and q.会社コード = 'E0' and q.処理年 = '" + td.StartYMD.AddMonths(-3).ToString("yyyy") + "' and q.処理月 = '" + td.StartYMD.AddMonths(-3).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 r on a.社員番号 = r.社員番号 and r.会社コード = 'E0' and r.処理年 = '" + td.StartYMD.AddMonths(-2).ToString("yyyy") + "' and r.処理月 = '" + td.StartYMD.AddMonths(-2).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 s on a.社員番号 = s.社員番号 and s.会社コード = 'E0' and s.処理年 = '" + td.StartYMD.AddMonths(-2).ToString("yyyy") + "' and s.処理月 = '" + td.StartYMD.AddMonths(-2).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 t on a.社員番号 = t.社員番号 and t.会社コード = 'E0' and t.処理年 = '" + td.StartYMD.AddMonths(-1).ToString("yyyy") + "' and t.処理月 = '" + td.StartYMD.AddMonths(-1).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 u on a.社員番号 = u.社員番号 and u.会社コード = 'E0' and u.処理年 = '" + td.StartYMD.AddMonths(-1).ToString("yyyy") + "' and u.処理月 = '" + td.StartYMD.AddMonths(-1).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKREKR4_E0 v on a.社員番号 = v.社員番号 and v.会社コード = 'E0' and v.処理年 = '" + td.StartYMD.AddMonths(0).ToString("yyyy") + "' and v.処理月 = '" + td.StartYMD.AddMonths(0).ToString("MM") + "' ";
            //sql += "left join dbo.KYTTKINTR3_E0 w on a.社員番号 = w.社員番号 and w.会社コード = 'E0' and w.処理年 = '" + td.StartYMD.AddMonths(0).ToString("yyyy") + "' and w.処理月 = '" + td.StartYMD.AddMonths(0).ToString("MM") + "' ";
            //sql += "where a.在籍区分 <> '9' and a.給与支給区分 in ('D1','E1','F1') ";


            sql += ") temp where[04月(168)] > 45 or[05月(176)] > 45 or[06月(168)] > 45 or[07月(176)] > 45 or[08月(176)] > 45 or[09月(168)] > 45 or[10月(176)] > 45 or[11月(168)] > 45 or[12月(176)] > 45 or[01月(176)] > 45 or[02月(160)] > 45 or[03月(176)] > 45 ";

            dataGridView1.DataSource = Com.GetDB(sql);

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button3.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.DataSource = null;

            string str = Com.GetDB("select 労働時間 from dbo.c所定労働テーブル where 年月 = '" + td.StartYMD.AddMonths(0).ToString("yyyyMM") + "'").Rows[0][0].ToString();
            string sql = "";

            sql += "select c.社員番号, c.地区名, c.現場名, c.組織名, c.給与支給区分名,c.氏名, F0500_延長時間 as 延長時間, F0800_所定 as 所定日数, b.A5700CG_勤務時間 as 勤務時間, b.A5800CG_時給 as 時給, F0700_遅刻時間 as 遅刻時間, ";
            sql += "((F0500_延長時間 + (b.A5700CG_勤務時間 * a.F0800_所定) - F0700_遅刻時間) - " + str + ") as 未支給対象時間, ";
            sql += "((F0500_延長時間 + (b.A5700CG_勤務時間 * a.F0800_所定) - F0700_遅刻時間) - " + str + ") * b.A5800CG_時給 * 0.25 as 未支給額 ";

            sql += "from QUATRO.dbo.KYTTKINTR3_E0 a left join dbo.KYTTKREKR4_E0 b on a.社員番号 = b.社員番号 and b.会社コード = 'E0' and b.処理年 = '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "' and b.処理月 = '" + td.StartYMD.AddMonths(1).ToString("MM") + "' ";
            sql += "left join dbo.社員基本情報 c on a.社員番号 = c.社員番号 where a.会社コード = 'E0' and a.処理年 = '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "' and a.処理月 = '" + td.StartYMD.AddMonths(1).ToString("MM") + "' and(F0500_延長時間 + (b.A5700CG_勤務時間 * a.F0800_所定) - F0700_遅刻時間) - " + str + " > 0 ";

            dataGridView1.DataSource = Com.GetDB(sql);

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button3.Enabled = true;
        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void label9_Click(object sender, EventArgs e)
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

            GetComboData();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem == null) return;
            //if (comboBox3.SelectedItem.ToString() == "給与預かり金" && dataGridView1.Columns[0].HeaderText == "経理地区")
            if (comboBox3.SelectedItem.ToString() == "給与預かり金")
            {

                //ソートエラー対応
                DataGridViewRow dgr = dataGridView1.CurrentRow;
                if (dgr == null) return;

                //string col = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderCell.Value.ToString();
                //string tiku = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                string koumoku = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                string naiyou = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();

                DataTable dt = new DataTable();
                string sql = "";
                if (koumoku.Substring(0, 1) == "B")
                {
                    //他
                    sql = "select * from dbo.k給与預金詳細一覧取得_固_地区廃止('" + y + "','" + m + "','" + koumoku + "')";
                }
                else
                {
                    //固定
                    sql = "select * from dbo.k給与預金詳細一覧取得_他_地区廃止('" + y + "','" + m + "','" + koumoku + "','" + naiyou + "')";
                }

                dt = Com.GetDB(sql);
                dataGridView3.DataSource = dt;

                //詳細
                dataGridView3.Columns[0].Width = 80;
                dataGridView3.Columns[1].Width = 200;
                dataGridView3.Columns[2].Width = 75;
                dataGridView3.Columns[3].Width = 120;
                dataGridView3.Columns[4].Width = 120;
                dataGridView3.Columns[5].Width = 120;
                dataGridView3.Columns[6].Width = 80;
                dataGridView3.Columns[7].Width = 250;

                dataGridView3.Columns[6].DefaultCellStyle.Format = "#,0";
                dataGridView3.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            printbtn.Enabled = false;

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();
            c1XLBook1.KeepFormulas = true;

            //-----------ここから
            if (comboBox3.SelectedItem.ToString() == "給与集計表")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\51_給与集計表.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["給与集計表"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "給与預かり金")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\52_給与預かり金.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["給与預かり金"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                            ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "給与預かり金_地区別")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\52_給与預かり金_地区別.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["給与預かり金_地区別"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "友の会費")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\54_友の会.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["友の会費"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "組織別人数")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\58_組織別人数.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["組織別人数"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "財形生保控除")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\55_財形生保控除.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["財形生保控除"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 12, j + 0].Value = dt.Rows[i][j];
                    }
                }

                int rows_sum = sumdt.Rows.Count;
                int cols_sum = sumdt.Columns.Count;
                for (int i = 0; i < rows_sum; i++)
                {
                    for (int j = 0; j < cols_sum; j++)
                    {
                        ls[i + 4, j + 0].Value = sumdt.Rows[i][j];
                    }
                }

            }
            
            else if (comboBox3.SelectedItem.ToString() == "金種表_個別")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\56_金種表_個別.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["金種表"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "金種表_組織別")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\57_金種表_組織別.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["金種表"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else if (comboBox3.SelectedItem.ToString() == "退職者源泉徴収一覧")
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\80_退職者源泉徴収一覧.xlsx");

                //リストシート
                XLSheet ls = c1XLBook1.Sheets["退職者源泉徴収一覧"];

                int rows = dt.Rows.Count;
                int cols = dt.Columns.Count;

                ls[0, 2].Value = y + "年" + m + "月支給";

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        ls[i + 3, j + 0].Value = dt.Rows[i][j];
                    }
                }
            }
            else
            {
                MessageBox.Show("おかしー。連絡ください。");
            }

            //ここまで

            string localPass = @"C:\ODIS\CalAfter\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");

            //PDF出力
            System.Diagnostics.Process.Start(exlName + ".xlsx");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            printbtn.Enabled = true;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button11.Enabled = false;

            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s出勤簿データ取得('" + comboBox5.SelectedItem.ToString().Replace("/", "") + "') where 担当区分 = '" + comboBox4.SelectedItem + "'order by 組織名, 現場名, カナ名");

            bool flg = true;
            if (comboBox4.SelectedItem.ToString() == "03_施設") flg = false;
            Com.GetSyukkinbo(dt, Convert.ToDateTime(comboBox5.SelectedItem.ToString() + "/01"), flg, false);

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button11.Enabled = true;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧 order by 組織CD, 現場CD");

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.[z残業一覧_組織別現場別] order by 組織CD, 現場CD");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_部門別 order by 担当区分");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_組織別 order by 組織CD");
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_残業のみ order by 組織CD, 現場CD");
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.[z残業一覧_組織別現場別_残業のみ] order by 組織CD, 現場CD");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_組織別_残業のみ order by 組織CD");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_部門別_残業のみ order by 担当区分");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_休出のみ order by 組織CD, 現場CD");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.[z残業一覧_組織別現場別_休出のみ] order by 組織CD, 現場CD");
        }

        private void button16_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_組織別_休出のみ order by 組織CD");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Com.GetDB("select * from dbo.z残業一覧_部門別_休出のみ order by 担当区分");
        }
    }
}
