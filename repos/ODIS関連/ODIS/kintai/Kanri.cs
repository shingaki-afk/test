using Npgsql;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Kanri : Form
    {
        private string yy = ""; //2020
        private string MM = ""; //06
        private string yyyyMM = ""; //202006
        private string yyadd = ""; //2020
        private string MMadd = ""; //07
        private string ymd = ""; //2020/06/30

        public Kanri()
        {
            if (Convert.ToInt16(Program.access) == 1)
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("現場損益権限無", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("2020");
            comboBox1.Items.Add("2021");
            comboBox1.Items.Add("2022");
            comboBox1.Items.Add("2023");
            comboBox1.SelectedIndex =2;

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
            string m = DateTime.Now.AddDays(-30).ToString("MM");

            comboBox2.SelectedIndex = comboBox2.FindString(m);

            GetYM();
        }

        private void GetYM()
        {
            DateTime dt = Convert.ToDateTime(this.comboBox1.SelectedItem.ToString() + "/" + this.comboBox2.SelectedItem.ToString() + "/01");
            yy = dt.ToString("yyyy");
            MM = dt.ToString("MM");
            yyyyMM = dt.ToString("yyyyMM");
            yyadd = dt.AddMonths(1).ToString("yyyy");
            MMadd = dt.AddMonths(1).ToString("MM");
            ymd = dt.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd");
            GetData();
        }

        private void GetData()
        {
            //ボタン無効化・カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            DataTable dt = new DataTable();

            //プロステージからデータをとってくる
            DataTable dttemp = new DataTable();
            dttemp = Com.GetPosDB("select ym,bumoncode, koujicode, sum(売上) 売上, sum(現場経費) 現場経費, sum(管理経費) 管理経費 from kpcp01.uriageGenbaKanrikeihi('" + yyyyMM + "') group by ym, bumoncode, koujicode");

            //一旦データ削除
            DataTable dtr = new DataTable();
            dtr = Com.GetDB("delete from u売上と現場経費と管理経費 where 年月 = '" + yyyyMM + "'");

            //ZeeMDBにインサート
            SqlConnection Cn;
            SqlCommand Cmd;

            try
            {
                using (Cn = new SqlConnection(ODIS.Com.SQLConstr))
                {
                    Cn.Open();
                    using (Cmd = Cn.CreateCommand())
                    {
                        using (SqlBulkCopy bulkcopy = new SqlBulkCopy(Cn))
                        {
                            bulkcopy.BulkCopyTimeout = 660;
                            bulkcopy.DestinationTableName = "u売上と現場経費と管理経費";
                            bulkcopy.WriteToServer(dttemp);
                            bulkcopy.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
            string sql = "";
            sql = "select * from dbo.b部門別現場別損益_月別_経費分('" + yyadd + "','" + MMadd + "','" + ymd + "','" + yyyyMM + "') order by 組織CD, 現場CD";
            dt = Com.GetDB(sql);

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[2].Width = 40;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 70;
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[6].Width = 70;
            dataGridView1.Columns[7].Width = 70;
            dataGridView1.Columns[8].Width = 70;
            dataGridView1.Columns[9].Width = 70;
            dataGridView1.Columns[10].Width = 40;
            dataGridView1.Columns[11].Width = 100;

            dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "0\'%\'";

            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.Beige;
            //dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.Beige;

            //dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.Beige;
            //dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.Beige;

            //dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            //dataGridView1.Columns[14].HeaderCell.Style.BackColor = Color.AntiqueWhite;

            //dataGridView1.Columns[15].DefaultCellStyle.Format = "#,0";

            Com.InHistory("現場別損益", "", "");

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

            //string row = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            string col = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderCell.Value.ToString();

            string bumon = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            string genba = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();

            DataTable dturiage = new DataTable();
            DataTable dtzinkenhi = new DataTable();
            DataTable dtsyokeihi = new DataTable();

            int nRet;

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    string sql = "";

                    //売上
                    sql += "select 契約区分,  ";
                    sql += "uriagekingaku as 売上額, 作業区分, torihikisakiname as 取引先名, keiyakukoumoku as 契約名 ";
                    sql += " from kpcp01.\"CostomGetUriageDataDetails\" ";
                    sql += "where uriageym = '" + "" + yyyyMM + "" + "' and bumoncode = '" + bumon + "' ";
                    
                    //技術企画とエンジと事務所
                    if (bumon.Substring(1, 3) == "202" || bumon.Substring(1, 3) == "102" || bumon.Substring(0, 1) == "1" || bumon.Substring(1, 4) == "9900" || bumon.Substring(1, 4) == "0000")
                    {

                    }
                    else if (bumon == "24055" && genba != "10101") // 指定管理植栽
                    {
                        sql += " and koujicode <> '" + genba + "'";
                    }
                    else if (bumon == "24055" && genba == "10101") // 指定管理　国際センター
                    {
                        sql += " and koujicode = '" + genba + "'";
                    }
                    else
                    {
                        sql += " and koujicode = '" + genba + "'";
                    }

                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dturiage);

                    //諸経費
                    sql = "";
                    sql += "select kamokuname as 科目名, ";
                    sql += " case when taisyakukubunb = '1' then denpyoukingaku else denpyoukingaku * -1 end as 金額, ";
                    //sql += " case when taisyakukubunb = '1' then denpyoukingaku + syouhizeikingaku else (denpyoukingaku + syouhizeikingaku) * -1 end as 税込額, ";
                    sql += " tekiyou as 摘要";
                    sql += " , denpyounumber as 伝票番号, gyounumber as 行番";
                    //sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where (kamokucode between '8231' and '8338' or kamokucode between '8354' and '8600') and suitouymd like " + "" + yyyyMM + "" + " || '%'  and bumoncode = '" + bumon + "' ";
                    sql += " from kpcp01.\"CostomGetDenpyouDataDetails\" where (kamokucode between '8231' and '8338' or kamokucode between '8354' and '9900') and suitouymd like " + "" + yyyyMM + "" + " || '%'  and bumoncode = '" + bumon + "' ";

                    //sql += " and koujicode = '" + genba + "'";

                    //技術企画とエンジと事務所
                    if (bumon.Substring(1, 3) == "202" || bumon.Substring(1, 3) == "102" || bumon.Substring(0, 1) == "1" || bumon.Substring(1, 4) == "9900" || bumon.Substring(1, 4) == "0000")
                    {

                    }
                    else if (bumon == "24055" && genba != "10101") // 指定管理植栽
                    {
                        sql += " and koujicode <> '" + genba + "'";
                    }
                    else if (bumon == "24055" && genba == "10101") // 指定管理　国際センター
                    {
                        sql += " and koujicode = '" + genba + "'";
                    }
                    else if (genba == "19900" || genba == "29900" || genba == "39900") //各部事務所
                    {
                        sql += " and (koujicode = '" + genba + "' or koujicode = '')";
                    }
                    else
                    {
                        sql += " and koujicode = '" + genba + "'";
                    }

                    sql += " order by kamokucode, bumoncode, koujicode";

                    NpgsqlDataAdapter adaptersyokeihi = new NpgsqlDataAdapter(sql, conn);
                    nRet = adaptersyokeihi.Fill(dtsyokeihi);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //人件費
            string zinsql = "select * from dbo.人件費詳細取得('" + yyadd + "','" + MMadd + "','" + ymd + "','" + bumon + "','" + genba + "')";
            dtzinkenhi = Com.GetDB(zinsql);


            dataGridView2.DataSource = dturiage;
            dataGridView3.DataSource = dtzinkenhi;
            dataGridView4.DataSource = dtsyokeihi;

            //売上
            dataGridView2.Columns[0].Width = 75;
            dataGridView2.Columns[1].Width = 75;
            dataGridView2.Columns[2].Width = 75;
            dataGridView2.Columns[3].Width = 350;
            dataGridView2.Columns[4].Width = 300;

            dataGridView2.Columns[1].DefaultCellStyle.Format = "#,0";
            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            //人件費
            dataGridView3.Columns[0].Width = 75;

            for (int i = 1; i < 9; i++)
            {
                dataGridView3.Columns[i].Width = 50;
                dataGridView3.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            dataGridView3.Columns[9].Width = 75;
            dataGridView3.Columns[10].Width = 75;
            dataGridView3.Columns[11].Width = 75;

            for (int i = 12; i < 41; i++)
            {
                dataGridView3.Columns[i].Width = 45;
                dataGridView3.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }



            //諸経費
            dataGridView4.Columns[0].Width = 75;
            dataGridView4.Columns[1].Width = 75;
            dataGridView4.Columns[2].Width = 350;
            dataGridView4.Columns[3].Width = 75;
            dataGridView4.Columns[4].Width = 50;
            dataGridView4.Columns[1].DefaultCellStyle.Format = "#,0";
            dataGridView4.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;
            GetYM();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1) return;

            GetYM();
        }
    }
}
