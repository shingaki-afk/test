using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class YosanTotal : Form
    {
        private string yms = "";
        private string yme = "";

        private string zi = "";

        public YosanTotal()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dgvyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //TODO 毎年追加
            cbki.Items.Add("53期(2024)");
            cbki.Items.Add("54期(2025)");

            cbki.SelectedIndex = cbki.Items.Count - 1;

            IniSet();
            FirstSet();
            //GetData();

            cbsub.Items.Add("実績");
            cbsub.Items.Add("前期");

            cbsub.SelectedIndex = 0;

            Com.InHistory("531_予算集計", "", "");
        }

        private void IniSet()
        {
            yms = cbki.SelectedItem.ToString().Substring(4, 4);
            yme = (Convert.ToInt16(yms) + 1).ToString();

            DataTable mokutable = new DataTable();
            mokutable = Com.GetDB("select 次 from dbo.y予算マスタ where 始年 = '" + yms + "' order by 次 ");
            //mokutable = Com.GetDB("select 次 from dbo.y予算マスタ where 始年 = '" + yms.Substring(0, 4) + "04" + "' order by 次 ");

            cbzi.Items.Clear();

            foreach (DataRow row in mokutable.Rows)
            {
                cbzi.Items.Add("第" + row[0] + "次");
                zi = row[0].ToString();
            }

            cbzi.SelectedIndex = cbzi.Items.Count - 1;

        }

        private void FirstSet()
        {
            cbs.Items.Clear();

            cbs.Items.Add(yms + "04");
            cbs.Items.Add(yms + "05");
            cbs.Items.Add(yms + "06");
            cbs.Items.Add(yms + "07");
            cbs.Items.Add(yms + "08");
            cbs.Items.Add(yms + "09");
            cbs.Items.Add(yms + "10");
            cbs.Items.Add(yms + "11");
            cbs.Items.Add(yms + "12");
            cbs.Items.Add(yme + "01");
            cbs.Items.Add(yme + "02");
            cbs.Items.Add(yme + "03");

            cbe.Items.Clear();

            cbe.Items.Add(yms + "04");
            cbe.Items.Add(yms + "05");
            cbe.Items.Add(yms + "06");
            cbe.Items.Add(yms + "07");
            cbe.Items.Add(yms + "08");
            cbe.Items.Add(yms + "09");
            cbe.Items.Add(yms + "10");
            cbe.Items.Add(yms + "11");
            cbe.Items.Add(yms + "12");
            cbe.Items.Add(yme + "01");
            cbe.Items.Add(yme + "02");
            cbe.Items.Add(yme + "03");

            cbs.SelectedIndex = 0;
            cbe.SelectedIndex = 11;
        }

        private void GetData()
        {
            if (zi == "") return;

            DataTable dt = new DataTable();

            string hikiateflg = "";

            if (checkBox1.Checked)
            {
                hikiateflg = "1";
            }
            else
            {
                hikiateflg = "0";
            }

            string sql = " select * from dbo.y予算集計_按分追加_率追加('" + cbs.SelectedItem.ToString() + "','" + cbe.SelectedItem.ToString() + "','" + zi + "', '" + hikiateflg + "') order by 部門 ";

            dt = Com.GetDB(sql);
            dgvyosan.DataSource = dt;

            //表示固定
            dgvyosan.Columns[0].Frozen = true;

            //売上合計
            int ct = dt.Rows.Count - 1;
            //label3.Text = (Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) * 0.04).ToString("N0");

            //label6.Text = (Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) * 0.04 - Convert.ToDouble(dt.Rows[ct]["部門利益"].ToString())).ToString("N0");

            //label13.Text = (Convert.ToDouble(dt.Rows[ct]["賞与"].ToString()) + Convert.ToDouble(dt.Rows[ct]["管理賞与"].ToString())).ToString("N0");
            //label14.Text = (Convert.ToDouble(dt.Rows[ct]["賞与"].ToString()) + Convert.ToDouble(dt.Rows[ct]["管理賞与"].ToString()) - Convert.ToDouble(120000000)).ToString("N0");

            double par = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                par += Convert.ToDouble(dt.Rows[i]["全社按分"]);
            }

            label15.Text = (par / Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) *100 ).ToString();

            dgvyosan.Columns["部門"].Width = 80;
            dgvyosan.Columns["固定売上"].Width = 80;
            dgvyosan.Columns["臨時売上"].Width = 80;
            dgvyosan.Columns["売上"].Width = 80;
            dgvyosan.Columns["人件費"].Width = 80;
            dgvyosan.Columns["賞与"].Width = 80;
            dgvyosan.Columns["退職金"].Width = 80;
            dgvyosan.Columns["諸経費"].Width = 80;
            dgvyosan.Columns["経費"].Width = 80;
            dgvyosan.Columns["現場利益"].Width = 80;
            dgvyosan.Columns["現場計数"].Width = 60;
            dgvyosan.Columns["管理人件費"].Width = 80;
            dgvyosan.Columns["管理賞与"].Width = 80;
            dgvyosan.Columns["管理退職金"].Width = 80;
            dgvyosan.Columns["管理諸経費"].Width = 80;
            dgvyosan.Columns["管理経費"].Width = 80;
            dgvyosan.Columns["部門利益"].Width = 80;
            dgvyosan.Columns["部門計数"].Width = 60;
            dgvyosan.Columns["地区按分"].Width = 70;
            dgvyosan.Columns["全社按分"].Width = 70;
            dgvyosan.Columns["按分後利益"].Width = 70;
            dgvyosan.Columns["各部管理経費率_分母現場人件費"].Width = 100;
            dgvyosan.Columns["共通管理経費率_分母現場人件費"].Width = 100;
            dgvyosan.Columns["合計管理経費率_分母現場人件費"].Width = 100;
            dgvyosan.Columns["各部管理経費率_分母売上"].Width = 100;
            dgvyosan.Columns["共通管理経費率_分母売上"].Width = 100;
            dgvyosan.Columns["合計管理経費率_分母売上"].Width = 100;
            dgvyosan.Columns["部門利益率_分母売上"].Width = 70;
            dgvyosan.Columns["部門利益_月平均"].Width = 70;
            //dgvyosan.Columns["前期部門利益_月平均"].Width = 70;
            //dgvyosan.Columns["部門利益_月平均_前年差額"].Width = 100;

            //ヘッダーの中央表示
            for (int i = 0; i < dgvyosan.Columns.Count; i++)
            {
                dgvyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //
            //            dgvyosan.Columns[ 1].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["部門"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["現場利益"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["現場計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvyosan.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["管理経費"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["部門利益"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["部門計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dgvyosan.Columns["地区按分"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["全社按分"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["按分後利益"].DefaultCellStyle.Format = "#,0";
            dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["部門利益率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dgvyosan.Columns["部門利益_月平均"].DefaultCellStyle.Format = "#,0";
            //dgvyosan.Columns["前期部門利益_月平均"].DefaultCellStyle.Format = "#,0";
            //dgvyosan.Columns["部門利益_月平均_前年差額"].DefaultCellStyle.Format = "#,0";

            //表示位置
            dgvyosan.Columns["部門"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvyosan.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["現場利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["現場計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["部門利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["部門計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["地区按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["全社按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["按分後利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["部門利益率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvyosan.Columns["部門利益_月平均"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["前期部門利益_月平均"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["部門利益_月平均_前年差額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            ////色変更
            dgvyosan.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dgvyosan.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosan.Columns["現場利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            dgvyosan.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dgvyosan.Columns["部門利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;

            dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;



            //前期テスト
            DataTable dtzen = new DataTable();

            string sqlzen = "";

            if (cbsub.SelectedItem.ToString() == "実績")
            {
                //実績
                sqlzen = " select * from dbo.y予算集計実績_按分追加_率追加('" + (Convert.ToInt64(cbs.SelectedItem.ToString())).ToString() + "','" + Convert.ToInt64(cbe.SelectedItem.ToString()).ToString() + "','" + hikiateflg + "') order by 部門 ";
            }
            else
            {
                //前期
                sqlzen = " select * from dbo.y予算集計実績_按分追加_率追加('" + (Convert.ToInt64(cbs.SelectedItem.ToString()) - 100).ToString() + "','" + (Convert.ToInt64(cbe.SelectedItem.ToString()) - 100).ToString() + "','" + hikiateflg + "') order by 部門 ";
            }

            dtzen = Com.GetDB(sqlzen);
            dataGridView1.DataSource = dtzen;

            //表示固定
            dataGridView1.Columns[0].Frozen = true;

            dataGridView1.Columns["部門"].Width = 80;
            dataGridView1.Columns["固定売上"].Width = 80;
            dataGridView1.Columns["臨時売上"].Width = 80;
            dataGridView1.Columns["売上"].Width = 80;
            dataGridView1.Columns["人件費"].Width = 80;
            dataGridView1.Columns["賞与"].Width = 80;
            dataGridView1.Columns["退職金"].Width = 80;
            dataGridView1.Columns["諸経費"].Width = 80;
            dataGridView1.Columns["経費"].Width = 80;
            dataGridView1.Columns["現場利益"].Width = 80;
            dataGridView1.Columns["現場計数"].Width = 60;
            dataGridView1.Columns["管理人件費"].Width = 80;
            dataGridView1.Columns["管理賞与"].Width = 80;
            dataGridView1.Columns["管理退職金"].Width = 80;
            dataGridView1.Columns["管理諸経費"].Width = 80;
            dataGridView1.Columns["管理経費"].Width = 80;
            dataGridView1.Columns["部門利益"].Width = 80;
            dataGridView1.Columns["部門計数"].Width = 60;
            dataGridView1.Columns["地区按分"].Width = 70;
            dataGridView1.Columns["全社按分"].Width = 70;
            dataGridView1.Columns["按分後利益"].Width = 70;
            dataGridView1.Columns["各部管理経費率_分母現場人件費"].Width = 100;
            dataGridView1.Columns["共通管理経費率_分母現場人件費"].Width = 100;
            dataGridView1.Columns["合計管理経費率_分母現場人件費"].Width = 100;
            dataGridView1.Columns["各部管理経費率_分母売上"].Width = 100;
            dataGridView1.Columns["共通管理経費率_分母売上"].Width = 100;
            dataGridView1.Columns["合計管理経費率_分母売上"].Width = 100;
            dataGridView1.Columns["部門利益率_分母売上"].Width = 70;
            dataGridView1.Columns["部門利益_月平均"].Width = 70;

            //dataGridView1.Columns["前期部門利益_月平均"].Width = 70;
            //dataGridView1.Columns["部門利益_月平均_前年差額"].Width = 100;




            //ヘッダーの中央表示
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //三桁区切り表示
            dataGridView1.Columns["部門"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["固定売上"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["臨時売上"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["売上"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["人件費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["賞与"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["退職金"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["諸経費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["経費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["現場利益"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["現場計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dataGridView1.Columns["管理人件費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["管理賞与"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["管理退職金"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["管理諸経費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["管理経費"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["部門利益"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["部門計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数
            dataGridView1.Columns["地区按分"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["全社按分"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["按分後利益"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["部門利益率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            dataGridView1.Columns["部門利益_月平均"].DefaultCellStyle.Format = "#,0";


            //表示位置
            dataGridView1.Columns["部門"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns["固定売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["臨時売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["現場利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["現場計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["管理人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["管理賞与"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["管理退職金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["管理諸経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["管理経費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["部門利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["部門計数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["地区按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["全社按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["按分後利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["部門利益率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["部門利益_月平均"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            ////色変更
            dataGridView1.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dataGridView1.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView1.Columns["現場利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            dataGridView1.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView1.Columns["部門利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            dataGridView1.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dataGridView1.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dataGridView1.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;

            dataGridView1.Columns["各部管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            dataGridView1.Columns["共通管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            dataGridView1.Columns["合計管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;


            if (checkBox1.Checked)
            {
                dgvyosan.Columns["賞与"].Visible = false;
                dgvyosan.Columns["退職金"].Visible = false;
                dgvyosan.Columns["管理賞与"].Visible = false;
                dgvyosan.Columns["管理退職金"].Visible = false;

                dataGridView1.Columns["賞与"].Visible = false;
                dataGridView1.Columns["退職金"].Visible = false;
                dataGridView1.Columns["管理賞与"].Visible = false;
                dataGridView1.Columns["管理退職金"].Visible = false;
            }
            else
            {
                dgvyosan.Columns["賞与"].Visible = true;
                dgvyosan.Columns["退職金"].Visible = true;
                dgvyosan.Columns["管理賞与"].Visible = true;
                dgvyosan.Columns["管理退職金"].Visible = true;

                dataGridView1.Columns["賞与"].Visible = true;
                dataGridView1.Columns["退職金"].Visible = true;
                dataGridView1.Columns["管理賞与"].Visible = true;
                dataGridView1.Columns["管理退職金"].Visible = true;
            }

            if (!checkBox2.Checked)
            {
                dgvyosan.Columns["各部管理経費率_分母現場人件費"].Visible = false;
                dgvyosan.Columns["共通管理経費率_分母現場人件費"].Visible = false;
                dgvyosan.Columns["合計管理経費率_分母現場人件費"].Visible = false;

                dgvyosan.Columns["各部管理経費率_分母売上"].Visible = false;
                dgvyosan.Columns["共通管理経費率_分母売上"].Visible = false;
                dgvyosan.Columns["合計管理経費率_分母売上"].Visible = false;

                dgvyosan.Columns["部門利益率_分母売上"].Visible = false;
                dgvyosan.Columns["部門利益_月平均"].Visible = false;

                //dgvyosan.Columns["前期部門利益_月平均"].Visible = false;
                //dgvyosan.Columns["部門利益_月平均_前年差額"].Visible = false;


                dataGridView1.Columns["各部管理経費率_分母現場人件費"].Visible = false;
                dataGridView1.Columns["共通管理経費率_分母現場人件費"].Visible = false;
                dataGridView1.Columns["合計管理経費率_分母現場人件費"].Visible = false;

                dataGridView1.Columns["各部管理経費率_分母売上"].Visible = false;
                dataGridView1.Columns["共通管理経費率_分母売上"].Visible = false;
                dataGridView1.Columns["合計管理経費率_分母売上"].Visible = false;

                dataGridView1.Columns["部門利益率_分母売上"].Visible = false;
                dataGridView1.Columns["部門利益_月平均"].Visible = false;

                //dataGridView1.Columns["前期部門利益_月平均"].Visible = false;
                //dataGridView1.Columns["部門利益_月平均_前年差額"].Visible = false;
            }
            else
            {
                dgvyosan.Columns["各部管理経費率_分母現場人件費"].Visible = true;
                dgvyosan.Columns["共通管理経費率_分母現場人件費"].Visible = true;
                dgvyosan.Columns["合計管理経費率_分母現場人件費"].Visible = true;

                dgvyosan.Columns["各部管理経費率_分母売上"].Visible = true;
                dgvyosan.Columns["共通管理経費率_分母売上"].Visible = true;
                dgvyosan.Columns["合計管理経費率_分母売上"].Visible = true;

                dgvyosan.Columns["部門利益率_分母売上"].Visible = true;
                dgvyosan.Columns["部門利益_月平均"].Visible = true;

                //dgvyosan.Columns["前期部門利益_月平均"].Visible = true;
                //dgvyosan.Columns["部門利益_月平均_前年差額"].Visible = true;


                dataGridView1.Columns["各部管理経費率_分母現場人件費"].Visible = true;
                dataGridView1.Columns["共通管理経費率_分母現場人件費"].Visible = true;
                dataGridView1.Columns["合計管理経費率_分母現場人件費"].Visible = true;

                dataGridView1.Columns["各部管理経費率_分母売上"].Visible = true;
                dataGridView1.Columns["共通管理経費率_分母売上"].Visible = true;
                dataGridView1.Columns["合計管理経費率_分母売上"].Visible = true;

                dataGridView1.Columns["部門利益率_分母売上"].Visible = true;
                dataGridView1.Columns["部門利益_月平均"].Visible = true;

                //dataGridView1.Columns["前期部門利益_月平均"].Visible = true;
                //dataGridView1.Columns["部門利益_月平均_前年差額"].Visible = true;
            }
        }

        private void cbs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            { 
                GetData();
            }
        }

        private void cbe_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                GetData();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                GetData();
            }
        }

        private void YosanTotal_Load(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                GetData();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbzi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                zi = cbzi.SelectedItem.ToString().Substring(1, 1);

                GetData();
            }
        }

        private void cbki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsub.SelectedItem != null && cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null )
            { 
                yms = cbki.SelectedItem.ToString().Substring(4, 4);
                yme = (Convert.ToInt16(yms) + 1).ToString();

                IniSet();
                FirstSet();
                GetData();
            }
        }
    }
}
