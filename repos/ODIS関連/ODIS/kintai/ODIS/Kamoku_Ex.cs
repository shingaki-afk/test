using Npgsql;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Kamoku_Ex : Form
    {
        //期が変わるタイミングで変更する必要がある
        private string maey = "2021";
        private string atoy = "2022";

        //科目コードの幅
        private string kamos = "8200";
        private string kamoe = "9900";

        public Kamoku_Ex()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("51期(2022～2023)");
            comboBox1.Items.Add("50期(2021～2022)");
            comboBox1.Items.Add("49期(2020～2021)");
            comboBox1.Items.Add("48期(2019～2020)");
            comboBox1.Items.Add("47期(2018～2019)");
            comboBox1.Items.Add("46期(2017～2018)");
            comboBox1.Items.Add("45期(2016～2017)");
            comboBox1.Items.Add("44期(2015～2016)");
            comboBox1.Items.Add("43期(2014～2015)");
            comboBox1.Items.Add("42期(2013～2014)");

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

            checkedListBox1.Items.Add("1_本社");
            checkedListBox1.Items.Add("2_那覇");
            checkedListBox1.Items.Add("3_八重山");
            checkedListBox1.Items.Add("4_北部");
            checkedListBox1.Items.Add("5_広域");
            checkedListBox1.Items.Add("6_宮古島");
            //checkedListBox1.Items.Add("7_久米島");

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
            int nRet;
            //string sql = "select distinct bumonkubun from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8000' and '8600' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' ";
            string sql = "select distinct bumonkubun from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '8000' and '9900' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' ";

            if (!checkedListBox1.GetItemChecked(0)) sql += " and bumoncode not like '1%' "; //本社
            if (!checkedListBox1.GetItemChecked(1)) sql += " and bumoncode not like '2%' "; //那覇
            if (!checkedListBox1.GetItemChecked(2)) sql += " and bumoncode not like '3%' "; //八重山
            if (!checkedListBox1.GetItemChecked(3)) sql += " and bumoncode not like '4%' "; //北部
            if (!checkedListBox1.GetItemChecked(4)) sql += " and bumoncode not like '5%' "; //広域
            if (!checkedListBox1.GetItemChecked(5)) sql += " and bumoncode not like '6%' "; //宮古島
            //if (!checkedListBox1.GetItemChecked(6)) sql += " and bumoncode not like '7%' "; //久米島
            sql += " order by bumonkubun ";
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    //string sql = "select distinct bumonkubun from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' order by bumonkubun";
                    //string sql = "select distinct bumonname from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' order by bumonkubun";
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

        private void SetGenba()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();
            int nRet;
            //売上
            string sql = "select distinct a.koujicode, b.koujiname from kpcp01.ctmgyoumukanritbl a left join kpcp01.wkkoujiview b on a.koujicode = b.koujicode and b.koujiedacode = '000' LEFT JOIN kpcp01.\"CtmSoshikiCD\" c ON a.bumoncode = c.bumoncode where a.sakujyo <> '1' and a.uriagecheck = '1' and a.uriageym between '" + maey + "04' and '" + atoy + "03'";
            if (!checkedListBox1.GetItemChecked(0)) sql += " and a.bumoncode not like '1%' "; //本社
            if (!checkedListBox1.GetItemChecked(1)) sql += " and a.bumoncode not like '2%' "; //那覇
            if (!checkedListBox1.GetItemChecked(2)) sql += " and a.bumoncode not like '3%' "; //八重山
            if (!checkedListBox1.GetItemChecked(3)) sql += " and a.bumoncode not like '4%' "; //北部
            if (!checkedListBox1.GetItemChecked(4)) sql += " and a.bumoncode not like '5%' "; //多面
            if (!checkedListBox1.GetItemChecked(5)) sql += " and a.bumoncode not like '6%' "; //宮古島
            //if (!checkedListBox1.GetItemChecked(6)) sql += " and a.bumoncode not like '7%' "; //久米島

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and c.bumonkubun <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += " union ";

            //経費
            //sql += "select distinct case when a.koujicode = '' then c.genbacode else a.koujicode end as koujicode, case when b.koujiname is null then c.genbaname else b.koujiname end as koujiname from kpcp01.\"CostomGetDenpyouDataDetails\" a left join kpcp01.wkkoujiview b on a.koujicode = b.koujicode and b.koujiedacode = '000' left join kpcp01.\"CtmGenbaNull\" c on a.bumoncode = c.bumoncode where a.kamokucode between '" + kamos + "' and '" + kamoe + "' and a.suitouymd between'" + maey + "0401' and '" + atoy + "0331' ";
            sql += "select distinct a.koujicode, b.koujiname from kpcp01.\"CostomGetDenpyouDataDetails\" a left join kpcp01.wkkoujiview b on a.koujicode = b.koujicode and b.koujiedacode = '000' where a.kamokucode between '" + kamos + "' and '" + kamoe + "' and a.suitouymd between'" + maey + "0401' and '" + atoy + "0331' ";
            sql += "";
            if (!checkedListBox1.GetItemChecked(0)) sql += " and a.bumoncode not like '1%' "; //本社
            if (!checkedListBox1.GetItemChecked(1)) sql += " and a.bumoncode not like '2%' "; //那覇
            if (!checkedListBox1.GetItemChecked(2)) sql += " and a.bumoncode not like '3%' "; //八重山
            if (!checkedListBox1.GetItemChecked(3)) sql += " and a.bumoncode not like '4%' "; //北部
            if (!checkedListBox1.GetItemChecked(4)) sql += " and a.bumoncode not like '5%' "; //多面
            if (!checkedListBox1.GetItemChecked(5)) sql += " and a.bumoncode not like '6%' "; //宮古島
            //if (!checkedListBox1.GetItemChecked(6)) sql += " and a.bumoncode not like '7%' "; //久米島

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and a.bumonkubun <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }
                
            sql += " order by koujicode ";

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    //string sql = "select distinct bumonkubun from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' order by bumonkubun";
                    //string sql = "select distinct bumonname from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' and suitouymd between'" + maey + "0401' and '" + atoy + "0331' order by bumonkubun";
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
                //checkedListBox3.Items.Add(row["koujicode"].ToString() + ' ' + row["koujiname"].ToString());
                checkedListBox3.Items.Add(row["koujicode"].ToString() + ' ' + row["koujiname"].ToString());
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

            DataTable dt = new DataTable();

            //期間指定
            maey = comboBox1.SelectedItem.ToString().Substring(4, 4);
            atoy = comboBox1.SelectedItem.ToString().Substring(9, 4);

            int nRet;

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    string sql = "";
                    sql += "select * from ( ";

                    //売上を業務管理テーブルから引っ張ってくる
                    sql += " select keiyakukubun || '-' || sagyoukubun as \"CD\", ";
                    sql += " CASE WHEN keiyakukubun = '01' and sagyoukubun = '0' THEN '契約固定(自社)'";
                    sql += " WHEN keiyakukubun = '02' and sagyoukubun = '0' THEN '契約臨時(自社)'";
                    sql += " WHEN keiyakukubun = '03' and sagyoukubun = '0' THEN '臨時(自社)'";
                    sql += " WHEN keiyakukubun = '04' THEN '物品'";
                    sql += " WHEN keiyakukubun = '01' and sagyoukubun = '1' THEN '契約固定(外注)'";
                    sql += " WHEN keiyakukubun = '02' and sagyoukubun = '1' THEN '契約臨時(外注)'";
                    sql += " WHEN keiyakukubun = '03' and sagyoukubun = '1' THEN '臨時(外注)' ELSE 'Error' END AS 科目名";
                    sql += " , sum(case when uriageym = '" + maey + "04' then uriagekingaku else 0 end) as \"04月\"";
                    sql += " , sum(case when uriageym = '" + maey + "05' then uriagekingaku else 0 end) as \"05月\"";
                    sql += " , sum(case when uriageym = '" + maey + "06' then uriagekingaku else 0 end) as \"06月\"";
                    sql += " , sum(case when uriageym = '" + maey + "07' then uriagekingaku else 0 end) as \"07月\"";
                    sql += " , sum(case when uriageym = '" + maey + "08' then uriagekingaku else 0 end) as \"08月\"";
                    sql += " , sum(case when uriageym = '" + maey + "09' then uriagekingaku else 0 end) as \"09月\"";
                    sql += " , sum(case when uriageym = '" + maey + "10' then uriagekingaku else 0 end) as \"10月\"";
                    sql += " , sum(case when uriageym = '" + maey + "11' then uriagekingaku else 0 end) as \"11月\"";
                    sql += " , sum(case when uriageym = '" + maey + "12' then uriagekingaku else 0 end) as \"12月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "01' then uriagekingaku else 0 end) as \"01月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "02' then uriagekingaku else 0 end) as \"02月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "03' then uriagekingaku else 0 end) as \"03月\"";
                    sql += " , sum(case when uriageym between '" + maey + "04' and '" + atoy + "03' then uriagekingaku else 0 end) as \"年間\"";
                    sql += " , sum(case when uriageym between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "04' and '" + maey + "03' then uriagekingaku else 0 end)/12 as \"前年月Ave\"";
                    sql += " from kpcp01.ctmgyoumukanritbl a LEFT JOIN kpcp01.\"CtmSoshikiCD\" b ON a.bumoncode = b.bumoncode";
                    sql += " where sakujyo <> '1' and uriagecheck = '1'";
                    //地区
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (!checkedListBox1.GetItemChecked(i))
                        {
                            if (!checkedListBox1.GetItemChecked(i))
                            {
                                sql += " and a.bumoncode not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
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

                    //現場
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        if (!checkedListBox3.GetItemChecked(i))
                        {
                            if (checkedListBox3.Items[i].ToString().Length < 5)
                            {
                                //sql += " and a.koujicode is not null ";
                                sql += " and a.koujicode <> '' ";
                            }
                            else
                            { 
                                sql += " and a.koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                            }
                        }
                    }


                    sql += " group by keiyakukubun, sagyoukubun ";

                    //売上合計
                    sql += " union all select '2000' as \"CD\", '【売上合計】' as 科目名 ";
                    sql += " , sum(case when uriageym = '" + maey + "04' then uriagekingaku else 0 end) as \"04月\"";
                    sql += " , sum(case when uriageym = '" + maey + "05' then uriagekingaku else 0 end) as \"05月\"";
                    sql += " , sum(case when uriageym = '" + maey + "06' then uriagekingaku else 0 end) as \"06月\"";
                    sql += " , sum(case when uriageym = '" + maey + "07' then uriagekingaku else 0 end) as \"07月\"";
                    sql += " , sum(case when uriageym = '" + maey + "08' then uriagekingaku else 0 end) as \"08月\"";
                    sql += " , sum(case when uriageym = '" + maey + "09' then uriagekingaku else 0 end) as \"09月\"";
                    sql += " , sum(case when uriageym = '" + maey + "10' then uriagekingaku else 0 end) as \"10月\"";
                    sql += " , sum(case when uriageym = '" + maey + "11' then uriagekingaku else 0 end) as \"11月\"";
                    sql += " , sum(case when uriageym = '" + maey + "12' then uriagekingaku else 0 end) as \"12月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "01' then uriagekingaku else 0 end) as \"01月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "02' then uriagekingaku else 0 end) as \"02月\"";
                    sql += " , sum(case when uriageym = '" + atoy + "03' then uriagekingaku else 0 end) as \"03月\"";
                    sql += " , sum(case when uriageym between '" + maey + "04' and '" + atoy + "03' then uriagekingaku else 0 end) as \"年間\"";
                    sql += " , sum(case when uriageym between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "04' and '" + maey + "03' then uriagekingaku else 0 end)/12 as \"前年月Ave\"";
                    sql += " from kpcp01.ctmgyoumukanritbl a LEFT JOIN kpcp01.\"CtmSoshikiCD\" b ON a.bumoncode = b.bumoncode";
                    sql += " where sakujyo <> '1' and uriagecheck = '1'";

                    //地区
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (!checkedListBox1.GetItemChecked(i))
                        {
                            if (!checkedListBox1.GetItemChecked(i))
                            {
                                sql += " and a.bumoncode not like '" + checkedListBox1.Items[i].ToString().Substring(0, 1) + "%'";
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

                    //現場
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        if (!checkedListBox3.GetItemChecked(i))
                        {
                            if (checkedListBox3.Items[i].ToString().Length < 5)
                            {
                                //sql += " and a.koujicode is not null ";
                                sql += " and a.koujicode <> '' ";
                            }
                            else
                            {
                                sql += " and a.koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                            }
                        }
                    }



                    sql += "  union all ";


                    //sql += "select kamokucode as \"CD\", max(case when uchiwakecode = '0000' then kamokuname else '' end) as 科目名";
                    sql += " select kamokucode as \"CD\", max(kamokuname) as 科目名";
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
                    sql += " , sum(case when suitouymd between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "0401' and '" + maey + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "0401' and '" + maey + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end)/12 as \"前年月Ave\"";
                    //sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" a left join kpcp01.\"CtmGenbaNull\" c on a.bumoncode = c.bumoncode where kamokucode between '" + kamos + "' and '" + kamoe + "' ";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' ";
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

                    //現場
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        if (!checkedListBox3.GetItemChecked(i))
                        {
                            if (checkedListBox3.Items[i].ToString().Length < 5)
                            {
                                //sql += " and koujicode is not null ";
                                //sql += " and case when koujicode = '' then c.genbacode else a.koujicode end <> '' ";
                                sql += " and koujicode <> '' ";
                            }
                            else
                            {
                                //sql += " and case when koujicode = '' then c.genbacode else a.koujicode end <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                                sql += " and koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                            }
                        }
                    }

                    sql += " group by kamokucode";
                    sql += " having sum(case when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + maey + "0401' and '" + atoy + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end) <> 0";

                    sql += " union all select '9990' as \"CD\", '【経費合計】' as 科目名 ";
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
                    sql += " , sum(case when suitouymd between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "0401' and '" + maey + "0331' and taisyakukubunb = '1' then denpyoukingaku when suitouymd between '" + Convert.ToString(Convert.ToInt16(maey) - 1) + "0401' and '" + maey + "0331' and taisyakukubunb = '2' then denpyoukingaku * -1 else 0 end)/12 as \"前年月Ave\"";
                    //sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" a left join kpcp01.\"CtmGenbaNull\" c on a.bumoncode = c.bumoncode where kamokucode between '" + kamos + "' and '" + kamoe + "' ";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "' ";

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

                    //現場
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        if (!checkedListBox3.GetItemChecked(i))
                        {
                            if (checkedListBox3.Items[i].ToString().Length < 5)
                            {

                                //sql += " and case when koujicode = '' then c.genbacode else a.koujicode end <> '' ";
                                sql += " and koujicode <> '' ";
                            }
                            else
                            {
                                //sql += " and case when koujicode = '' then c.genbacode else a.koujicode end <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                                sql += " and koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                            }
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



            if (dt.Rows.Count == 0)
            {
                //return;
            }
            else
            {

                //抽出
                DataRow[] keihiall;
                keihiall = dt.Select("CD = '9990'", "");
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

                DataRow[] uriall;
                uriall = dt.Select("CD = '2000'", "");
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


                //利益
                DataRow dr2 = dt.NewRow();
                dr2[0] = "9995";
                dr2[1] = "利益";
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
                dt.Rows.Add(dr2);

                //計数
                DataRow dr = dt.NewRow();
                dr[0] = "9999";
                dr[1] = "計数";
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
                dt.Rows.Add(dr);
            }

            dataGridView1.DataSource = dt;

            ////構造だけコピー
            //DataTable dt2 = new DataTable();
            //dt2 = dt.Clone();
            //dt2.Rows.Add(dtrow);

            //dataGridView3.DataSource = dt2;
            //foreach (DataRow row in dtrow)
            //{
            //    DataRow nr = dt2.NewRow();
            //    nr["社員番号"] = row["社員番号"];
            //    nr["氏名"] = row["氏名"];
            //    nr["組織名"] = row["組織名"];
            //    nr["現場名"] = row["現場名"];
            //    nr["登録状況"] = row["登録フラグ"].ToString() == "1" ? "登録済" : "";

            //    dt2.Rows.Add(nr);
            //}

            ///comboBox1.

            //表示処理
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 120;

            for (int i = 2; i < 16; i++)
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


            //dataGridView1.Columns[15].DefaultCellStyle.Format = "#,0";


            Com.InHistory("03_科目別損益(～23/03)", "", "");

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

                    //売上
                    if (row.Substring(0, 1) == "0"　|| row == "2000")
                    {
                        sql += "select uriageym as 売上年月, 契約区分, 作業区分, bumonname as 部門名, koujiname as 現場名, keiyakukoumoku as 契約項目,  ";
                        sql += "uriagekingaku as 売上額, torihikisakiname as 取引先名,  担当者氏名, 入力者, 更新者 ";
                        //bumoncode = '21010' and koujicode = '10101' and uriageym = '201904'";
                        sql += " from kpcp01.\"CostomGetUriageDataDetails\" ";
                        if (col == "年間")
                        {
                            sql += " where uriageym between '" + maey + "04' and '" + atoy + "03'";
                        }
                        else
                        {
                            if (col == "01月" || col == "02月" || col == "03月")
                            {
                                sql += " where uriageym = '" + atoy + col.Replace("月", "") + "'";
                            }
                            else
                            {
                                sql += " where uriageym = '" + maey + col.Replace("月", "") + "'";
                            }
                        }

                        if (row == "2000")
                        {

                        }
                        else
                        {
                            if (row == "01-0")
                            {
                                sql += " and 契約区分 = '契約固定' and 作業区分 = '自社' ";
                            }
                            else if (row == "01-1")
                            {
                                sql += " and 契約区分 = '契約固定' and 作業区分 = '外注' ";
                            }
                            else if (row == "02-0")
                            {
                                sql += " and 契約区分 = '契約臨時' and 作業区分 = '自社' ";
                            }
                            else if (row == "02-1")
                            {
                                sql += " and 契約区分 = '契約臨時' and 作業区分 = '外注' ";
                            }
                            else if (row == "03-0")
                            {
                                sql += " and 契約区分 = '臨時' and 作業区分 = '自社' ";
                            }
                            else if (row == "03-1")
                            {
                                sql += " and 契約区分 = '臨時' and 作業区分 = '外注' ";
                            }
                            else if (row == "04-0")
                            {
                                sql += " and 契約区分 = '物品' and 作業区分 = '自社' ";
                            }
                            else if (row == "04-1")
                            {
                                sql += " and 契約区分 = '物品' and 作業区分 = '外注' ";
                            }
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

                        //現場
                        for (int i = 0; i < checkedListBox3.Items.Count; i++)
                        {
                            if (!checkedListBox3.GetItemChecked(i))
                            {
                                if (checkedListBox3.Items[i].ToString().Length < 5)
                                {
                                }
                                else
                                {
                                    sql += " and koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                                }
                            }
                        }

                    }
                    //else if (row == "8212")
                    //{
                    //    //8212 現場給与
                    //    dt = Com.GetDB("select * from dbo.KM");
                    //}
                    else
                    {
                        //経費
                        sql += "select kamokuname as 科目名, bumonname as 部門名, koujiname as 現場名,";
                        sql += " case when taisyakukubunb = '1' then denpyoukingaku else denpyoukingaku * -1 end as 金額, ";
                        sql += " case when taisyakukubunb = '1' then denpyoukingaku + syouhizeikingaku else (denpyoukingaku + syouhizeikingaku) * -1 end as 税込額, ";
                        sql += " tekiyou as 摘要, torihikisakiname as 取引先名";
                        sql += " , suitouymd as 日付, denpyounumber as 伝票番号, gyounumber as 行番";
                        sql += " , inputtantousyaname as 入力者, registrationtantousyaname as 更新者";
                        sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where kamokucode between '" + kamos + "' and '" + kamoe + "'";
                        if (col == "年間")
                        {
                            sql += " and suitouymd between '" + maey + "0401' and '" + atoy + "0331'";
                        }
                        else
                        {
                            if (col == "01月" || col == "02月" || col == "03月")
                            {
                                sql += " and suitouymd like '" + atoy + col.Replace("月", "") + "%'";
                            }
                            else
                            {
                                sql += " and suitouymd like '" + maey + col.Replace("月", "") + "%'";
                            }
                        }

                        if (row == "9990") //経費合計、粗利、計数
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

                        //現場
                        for (int i = 0; i < checkedListBox3.Items.Count; i++)
                        {
                            if (!checkedListBox3.GetItemChecked(i))
                            {
                                if (checkedListBox3.Items[i].ToString().Length < 5)
                                {
                                }
                                else
                                {
                                    sql += " and koujicode <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                                }
                            }
                        }

                        sql += " order by kamokucode, bumoncode, koujicode";

                    }

                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            dataGridView2.DataSource = dt;

            //売上
            if (row.Substring(0, 1) == "0")
            {
                //MessageBox.Show("売上！");
                //TODO 
            }
            else
            {
                dataGridView2.Columns[0].Width = 100;//科目名
                dataGridView2.Columns[1].Width = 100;//部門名
                dataGridView2.Columns[2].Width = 250;//現場名
                dataGridView2.Columns[3].Width = 70;//金額
                dataGridView2.Columns[4].Width = 70;//金額(税込)
                dataGridView2.Columns[5].Width = 400;//摘要
                dataGridView2.Columns[6].Width = 250;//取引先

                //金額右寄
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //三桁区切り表示
                dataGridView2.Columns[3].DefaultCellStyle.Format = "#,0";
                dataGridView2.Columns[4].DefaultCellStyle.Format = "#,0";
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
    }
}
