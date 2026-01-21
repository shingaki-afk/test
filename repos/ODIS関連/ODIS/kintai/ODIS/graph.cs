using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Npgsql;
using System.Data.SqlClient;
using System.Windows.Forms.DataVisualization.Charting;

namespace ODIS.ODIS
{
    public partial class graph : Form
    {
        public graph()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            comboBox1.Items.Add("全地区");
            comboBox1.Items.Add("那覇");
            comboBox1.Items.Add("八重山");
            comboBox1.Items.Add("北部");
            comboBox1.Items.Add("広域");
            comboBox1.Items.Add("宮古島");
            comboBox1.Items.Add("久米島");
            comboBox1.SelectedIndex = 0;

            GetData();

            Com.InHistory("53_売上グラフ", "", "");
        }

        private void GetData()
        {
            //過去売上
            DataTable dtC = GetDataSql();

            //グラフ初期化
            chart1.Series.Clear();
            chart1.Legends.Clear();

            chart1.Palette = ChartColorPalette.Pastel;

            string[] months = { "04", "05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03" };

            foreach (string s in months)
            {
                chart1.Series.Add(s);

                if (s == "04")
                {
                    chart1.Legends.Add(s);
                }

                chart1.Series[s].ChartType = SeriesChartType.StackedColumn;　//グラフの種類を指定（Columnは棒グラフ）
                chart1.Series[s].LegendText = s + "月";  //凡例に表示するテキストを指定

                chart1.Series[s].IsValueShownAsLabel = true;//値表示
                chart1.Series[s].LabelFormat = "#,##0";//値表示
            }

            //X軸を全て表示
            chart1.ChartAreas[0].AxisX.Interval = 1;

            foreach (DataRow row in dtC.Rows)
            {
                chart1.Series[Convert.ToString(row["年月"]).Substring(4, 2)].Points.AddXY(Convert.ToString(row["年月"]).Substring(0, 4) + "年", row["売上"]);
            }

            //42期分
            foreach (DataRow row in GetDataMonths().Rows)
            {
                chart1.Series[Convert.ToString(row["uriageym"]).Substring(4, 2)].Points.AddXY(Convert.ToString(row["uriageym"]).Substring(0, 4) + "年", row["sum"]);
            }

            chart1.ChartAreas[0].AxisY.Minimum = 0;
            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "#,##0";
            chart1.Width = 1200;
            chart1.Height = 800;

            //副目盛線
            chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = true;
            chart1.ChartAreas[0].AxisY.MinorGrid.Interval = 100000000;
            chart1.ChartAreas[0].AxisY.MinorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightBlue;


        }


        private DataTable GetDataMonths()
        {
            int nRet;

            NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr);
            conn.Open();

            DataTable dt = new DataTable();
            string sql = "";
            if (comboBox1.Text == "全地区")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' group by uriageym order by uriageym";
            else if (comboBox1.Text == "八重山")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '3%' group by uriageym order by uriageym";
            else if (comboBox1.Text == "北部")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '4%' group by uriageym order by uriageym";
            else if (comboBox1.Text == "広域")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '5%' group by uriageym order by uriageym";
            else if (comboBox1.Text == "宮古島")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '6%' group by uriageym order by uriageym";
            else if (comboBox1.Text == "久米島")
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '7%' group by uriageym order by uriageym";

            else //那覇
                sql = "select uriageym, sum(uriagekingaku) from kpcp01.ctmgyoumukanritbl where uriageym between '201301' and '202403' and sakujyo = '0' and uriagecheck = '1' and bumoncode like '2%' group by uriageym order by uriageym";

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
            nRet = adapter.Fill(dt);

            conn.Close();

            //解放
            adapter.Dispose();
            conn.Dispose();

            return dt;
        }

        private DataTable GetDataSql()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter Adapter;
            DataTable dt = new DataTable();

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cmd = Cn.CreateCommand();
                //Cmd.CommandText = "select 年月, sum(売上) as 売上 from (select substring(convert(varchar(8), 伝票日付), 1, 6) as 年月, case when 貸借区分 = '2' then 金額 when 貸借区分 = '1' then 金額*-1 end as 売上 from dbo.旧伝票 where 科目コード like '801%') temp group by 年月 order by 年月";
                if (comboBox1.Text == "全地区")
                    Cmd.CommandText = "select 年月, sum(売上) as 売上 from (select 年月度 as 年月, case when 取引名 = '売上' then 金額 when 取引名 <> '1' then 金額*-1 end as 売上 from dbo.過去売上 where 年月度 between 200604 and 201212) temp group by 年月 order by 年月";
                else
                    Cmd.CommandText = "select 年月, sum(売上) as 売上 from (select 年月度 as 年月, case when 取引名 = '売上' then 金額 when 取引名 <> '1' then 金額*-1 end as 売上 from dbo.過去売上 where 年月度 between 200604 and 201212 and 地区名 like '" + comboBox1.Text + "' ) temp group by 年月 order by 年月";
                Adapter = new SqlDataAdapter(Cmd);
                Adapter.Fill(dt);
            }
            return dt;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
