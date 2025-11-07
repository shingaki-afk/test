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
    public partial class ZissekiHi : Form
    {

        //TODO 毎年変更
        private string yms = "202404";
        //private string yme = "202503";

        private string zi = "";

        public ZissekiHi()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dgvyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            FirstSet();
            GetData();

            Com.InHistory("55_実績×予算/前年", "", "");
        }

        private void FirstSet()
        {
            cbs.Items.Add("202404");
            cbs.Items.Add("202405");
            cbs.Items.Add("202406");
            cbs.Items.Add("202407");
            cbs.Items.Add("202408");
            cbs.Items.Add("202409");
            cbs.Items.Add("202410");
            cbs.Items.Add("202411");
            cbs.Items.Add("202412");
            cbs.Items.Add("202501");
            cbs.Items.Add("202502");
            cbs.Items.Add("202503");

            cbe.Items.Add("202404");
            cbe.Items.Add("202405");
            cbe.Items.Add("202406");
            cbe.Items.Add("202407");
            cbe.Items.Add("202408");
            cbe.Items.Add("202409");
            cbe.Items.Add("202410");
            cbe.Items.Add("202411");
            cbe.Items.Add("202412");
            cbe.Items.Add("202501");
            cbe.Items.Add("202502");
            cbe.Items.Add("202503");

            cbs.SelectedIndex = 0;
            cbe.SelectedIndex = 11;

            DataTable mokutable = new DataTable();
            mokutable = Com.GetDB("select 次 from dbo.y予算マスタ where 始年 = '" + yms.Substring(0, 4) + "' order by 次 ");

            foreach (DataRow row in mokutable.Rows)
            {
                comboBox1.Items.Add("第" + row[0] + "次");
                zi = row[0].ToString();
            }

            comboBox1.SelectedIndex = comboBox1.Items.Count - 1;
        }

        private void GetData()
        {
            if (zi == "") return;

            DataTable dt = new DataTable();
            string sql = " select * from dbo.y予算集計_実績比較_千円単位('" + cbs.SelectedItem.ToString() + "','" + cbe.SelectedItem.ToString() + "','" + zi + "', '0') order by 部門, 項目 ";
            dt = Com.GetDB(sql);
            dgvyosan.DataSource = dt;

            //表示固定
            dgvyosan.Columns[0].Frozen = true;

            ////売上合計
            //int ct = dt.Rows.Count - 1;
            //label3.Text = (Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) * 0.04).ToString("N0");

            //label6.Text = (Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) * 0.04 - Convert.ToDouble(dt.Rows[ct]["部門利益"].ToString())).ToString("N0");

            //label13.Text = (Convert.ToDouble(dt.Rows[ct]["賞与"].ToString()) + Convert.ToDouble(dt.Rows[ct]["管理賞与"].ToString())).ToString("N0");
            //label14.Text = (Convert.ToDouble(dt.Rows[ct]["賞与"].ToString()) + Convert.ToDouble(dt.Rows[ct]["管理賞与"].ToString()) - Convert.ToDouble(120000000)).ToString("N0");

            //double par = 0;
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    par += Convert.ToDouble(dt.Rows[i]["全社按分"]);
            //}

            //label15.Text = (par / Convert.ToDouble(dt.Rows[ct]["売上"].ToString()) *100 ).ToString();

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
            //dgvyosan.Columns["地区按分"].Width = 70;
            //dgvyosan.Columns["全社按分"].Width = 70;
            //dgvyosan.Columns["按分後利益"].Width = 70;
            //dgvyosan.Columns["各部管理経費率_分母現場人件費"].Width = 70;
            //dgvyosan.Columns["共通管理経費率_分母現場人件費"].Width = 70;
            //dgvyosan.Columns["合計管理経費率_分母現場人件費"].Width = 70;
            //dgvyosan.Columns["各部管理経費率_分母売上"].Width = 70;
            //dgvyosan.Columns["共通管理経費率_分母売上"].Width = 70;
            //dgvyosan.Columns["合計管理経費率_分母売上"].Width = 70;
            //dgvyosan.Columns["部門利益率_分母売上"].Width = 70;
            //dgvyosan.Columns["部門利益_月平均"].Width = 70;
            //dgvyosan.Columns["前期部門利益_月平均"].Width = 70;
            //dgvyosan.Columns["部門利益_月平均_前年差額"].Width = 70;

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
            //dgvyosan.Columns["地区按分"].DefaultCellStyle.Format = "#,0";
            //dgvyosan.Columns["全社按分"].DefaultCellStyle.Format = "#,0";
            //dgvyosan.Columns["按分後利益"].DefaultCellStyle.Format = "#,0";
            //dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["部門利益率_分母売上"].DefaultCellStyle.Format = "0.00\'%\'";
            //dgvyosan.Columns["部門利益_月平均"].DefaultCellStyle.Format = "#,0";
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
            //dgvyosan.Columns["地区按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["全社按分"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["按分後利益"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["部門利益率_分母売上"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["部門利益_月平均"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["前期部門利益_月平均"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgvyosan.Columns["部門利益_月平均_前年差額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            ////色変更
            //dgvyosan.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            //dgvyosan.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            //dgvyosan.Columns["現場利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            //dgvyosan.Columns["管理経費"].DefaultCellStyle.BackColor = Color.Khaki;
            //dgvyosan.Columns["部門利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            //dgvyosan.Columns["各部管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            //dgvyosan.Columns["共通管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;
            //dgvyosan.Columns["合計管理経費率_分母現場人件費"].DefaultCellStyle.BackColor = Color.LightSlateGray;

            //dgvyosan.Columns["各部管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            //dgvyosan.Columns["共通管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;
            //dgvyosan.Columns["合計管理経費率_分母売上"].DefaultCellStyle.BackColor = Color.LightSalmon;

            //dgvyosan.Rows[10].DefaultCellStyle.BackColor = Color.PaleGreen; //売上
            //dgvyosan.Rows[5].DefaultCellStyle.BackColor = Color.Khaki; //経費




        }

        private void cbs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbs.SelectedItem != null && cbe.SelectedItem != null)
            { 
                GetData();
            }
        }

        private void cbe_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbs.SelectedItem != null && cbe.SelectedItem != null)
            {
                GetData();
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            zi = comboBox1.SelectedItem.ToString().Substring(1, 1);

            GetData();
        }

        private void dgvyosan_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
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
            }
        }

        private void dgvyosan_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.Value != null && e.ColumnIndex == 1 && decimal.TryParse(e.Value.ToString(), out val))
            if (e.Value != null && e.ColumnIndex == 1 && e.Value.ToString() == "03_差額" && e.RowIndex % 2 == 0)
            {
                    dgvyosan.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PowderBlue;
            }
            else if (e.Value != null && e.ColumnIndex == 1 && e.Value.ToString() == "03_差額" && e.RowIndex % 2 != 0)
            {
                dgvyosan.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Thistle;
            }
        }
    }
}
