using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using ODIS.ODIS;

namespace ODIS.ODIS
{
    public partial class KeisuuYM : Form
    {
        public KeisuuYM()
        {
            InitializeComponent();
        }

        public KeisuuYM(string bumon, string genba, string bumoncd, string genbacd)
        {
            InitializeComponent();

            //フォームにスクロールバ
            this.AutoScroll = true;

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);
            dataGridView2.Font = new Font(dataGridView2.Font.Name, 12);
            dataGridView3.Font = new Font(dataGridView3.Font.Name, 12);

            //行ヘッダを非表示
            dataGridView1.RowHeadersVisible = false;
            dataGridView2.RowHeadersVisible = false;
            dataGridView3.RowHeadersVisible = false;

            label4.Text = bumon;
            label5.Text = genba;

            GetData(bumon, genba, bumoncd, genbacd);
        }

        private void GetData(string bumon, string genba, string bumoncd, string genbacd)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable dt = new DataTable();
            string sql = "";

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        sql += "select 年月, 固定売上, 臨時売上, 売上, 人件費, 諸経費, 経費, 利益, 計数, 評価 from dbo.kanrikeisuu ";
                        sql += "where 年月 between 202004 and 202303 and 部門ＣＤ = " + bumoncd + " and 現場ＣＤ = " + genbacd + " ";

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

            #region 今期
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

            string nendoA = "2022";
            string nendoB = "2023";

            foreach (DataRow row in wkdt.Rows)
            {
                decimal sum = 0;
                DataRow nr = gendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains(nendoA + "04")) { nr["４月"] = row[nendoA + "04"]; sum += Convert.ToDecimal(row[nendoA + "04"]); }
                if (row.Table.Columns.Contains(nendoA + "05")) { nr["５月"] = row[nendoA + "05"]; sum += Convert.ToDecimal(row[nendoA + "05"]); }
                if (row.Table.Columns.Contains(nendoA + "06")) { nr["６月"] = row[nendoA + "06"]; sum += Convert.ToDecimal(row[nendoA + "06"]); }
                if (row.Table.Columns.Contains(nendoA + "07")) { nr["７月"] = row[nendoA + "07"]; sum += Convert.ToDecimal(row[nendoA + "07"]); }
                if (row.Table.Columns.Contains(nendoA + "08")) { nr["８月"] = row[nendoA + "08"]; sum += Convert.ToDecimal(row[nendoA + "08"]); }
                if (row.Table.Columns.Contains(nendoA + "09")) { nr["９月"] = row[nendoA + "09"]; sum += Convert.ToDecimal(row[nendoA + "09"]); }
                if (row.Table.Columns.Contains(nendoA + "10")) { nr["１０月"] = row[nendoA + "10"]; sum += Convert.ToDecimal(row[nendoA + "10"]); }
                if (row.Table.Columns.Contains(nendoA + "11")) { nr["１１月"] = row[nendoA + "11"]; sum += Convert.ToDecimal(row[nendoA + "11"]); }
                if (row.Table.Columns.Contains(nendoA + "12")) { nr["１２月"] = row[nendoA + "12"]; sum += Convert.ToDecimal(row[nendoA + "12"]); }
                if (row.Table.Columns.Contains(nendoB + "01")) { nr["１月"] = row[nendoB + "01"]; sum += Convert.ToDecimal(row[nendoB + "01"]); }
                if (row.Table.Columns.Contains(nendoB + "02")) { nr["２月"] = row[nendoB + "02"]; sum += Convert.ToDecimal(row[nendoB + "02"]); }
                if (row.Table.Columns.Contains(nendoB + "03")) { nr["３月"] = row[nendoB + "03"]; sum += Convert.ToDecimal(row[nendoB + "03"]); }
                nr["年間"] = sum;
                gendt.Rows.Add(nr);
            }

            decimal SumK = 0;
            int level = 0;

            SumK = Convert.ToDecimal(gendt.Rows[3]["年間"]) == 0 ? 0 : Convert.ToDecimal(gendt.Rows[6]["年間"]) / Convert.ToDecimal(gendt.Rows[3]["年間"]) * 100;

            if (SumK >= 100)
            {
                level = 1;
            }
            else if (SumK >= 90 && SumK < 100)
            {
                level = 2;
            }
            else if (SumK >= 85 && SumK < 90)
            {
                level = 3;
            }
            else if (SumK >= 80 && SumK < 85)
            {
                level = 4;
            }
            else if (SumK >= 70 && SumK < 80)
            {
                level = 5;
            }
            else if (SumK >= 60 && SumK < 70)
            {
                level = 6;
            }
            else if (SumK < 60)
            {
                level = 7;
            }
            else if (SumK == 0)
            {
                level = 0;
            }
            else
            {
                level = 8;
            }

            gendt.Rows[8]["年間"] = SumK;
            gendt.Rows[9]["年間"] = level;

            dataGridView1.DataSource = gendt;


            #endregion

            #region 前期

            nendoA = "2021";
            nendoB = "2022";

            foreach (DataRow row in wkdt.Rows)
            {
                decimal sum = 0;
                DataRow nr = zendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains(nendoA + "04")) { nr["４月"] = row[nendoA + "04"]; sum += Convert.ToDecimal(row[nendoA + "04"]); }
                if (row.Table.Columns.Contains(nendoA + "05")) { nr["５月"] = row[nendoA + "05"]; sum += Convert.ToDecimal(row[nendoA + "05"]); }
                if (row.Table.Columns.Contains(nendoA + "06")) { nr["６月"] = row[nendoA + "06"]; sum += Convert.ToDecimal(row[nendoA + "06"]); }
                if (row.Table.Columns.Contains(nendoA + "07")) { nr["７月"] = row[nendoA + "07"]; sum += Convert.ToDecimal(row[nendoA + "07"]); }
                if (row.Table.Columns.Contains(nendoA + "08")) { nr["８月"] = row[nendoA + "08"]; sum += Convert.ToDecimal(row[nendoA + "08"]); }
                if (row.Table.Columns.Contains(nendoA + "09")) { nr["９月"] = row[nendoA + "09"]; sum += Convert.ToDecimal(row[nendoA + "09"]); }
                if (row.Table.Columns.Contains(nendoA + "10")) { nr["１０月"] = row[nendoA + "10"]; sum += Convert.ToDecimal(row[nendoA + "10"]); }
                if (row.Table.Columns.Contains(nendoA + "11")) { nr["１１月"] = row[nendoA + "11"]; sum += Convert.ToDecimal(row[nendoA + "11"]); }
                if (row.Table.Columns.Contains(nendoA + "12")) { nr["１２月"] = row[nendoA + "12"]; sum += Convert.ToDecimal(row[nendoA + "12"]); }
                if (row.Table.Columns.Contains(nendoB + "01")) { nr["１月"] = row[nendoB + "01"]; sum += Convert.ToDecimal(row[nendoB + "01"]); }
                if (row.Table.Columns.Contains(nendoB + "02")) { nr["２月"] = row[nendoB + "02"]; sum += Convert.ToDecimal(row[nendoB + "02"]); }
                if (row.Table.Columns.Contains(nendoB + "03")) { nr["３月"] = row[nendoB + "03"]; sum += Convert.ToDecimal(row[nendoB + "03"]); } 
                nr["年間"] = sum;
                zendt.Rows.Add(nr);
            }


            SumK = 0;
            level = 0;

            SumK = Convert.ToDecimal(zendt.Rows[3]["年間"]) == 0 ? 0 : Convert.ToDecimal(zendt.Rows[6]["年間"]) / Convert.ToDecimal(zendt.Rows[3]["年間"]) * 100;

            level = Com.GetLevel(SumK);

            zendt.Rows[8]["年間"] = SumK;
            zendt.Rows[9]["年間"] = level;

            dataGridView2.DataSource = zendt;


            #endregion

            #region 前々期

            nendoA = "2020";
            nendoB = "2021";

             foreach (DataRow row in wkdt.Rows)
            {
                decimal sum = 0;
                DataRow nr = zenzendt.NewRow();
                nr["項目"] = row["年月"];
                if (row.Table.Columns.Contains(nendoA + "04")) { nr["４月"] = row[nendoA + "04"]; sum += Convert.ToDecimal(row[nendoA + "04"]); }
                if (row.Table.Columns.Contains(nendoA + "05")) { nr["５月"] = row[nendoA + "05"]; sum += Convert.ToDecimal(row[nendoA + "05"]); }
                if (row.Table.Columns.Contains(nendoA + "06")) { nr["６月"] = row[nendoA + "06"]; sum += Convert.ToDecimal(row[nendoA + "06"]); }
                if (row.Table.Columns.Contains(nendoA + "07")) { nr["７月"] = row[nendoA + "07"]; sum += Convert.ToDecimal(row[nendoA + "07"]); }
                if (row.Table.Columns.Contains(nendoA + "08")) { nr["８月"] = row[nendoA + "08"]; sum += Convert.ToDecimal(row[nendoA + "08"]); }
                if (row.Table.Columns.Contains(nendoA + "09")) { nr["９月"] = row[nendoA + "09"]; sum += Convert.ToDecimal(row[nendoA + "09"]); }
                if (row.Table.Columns.Contains(nendoA + "10")) { nr["１０月"] = row[nendoA + "10"]; sum += Convert.ToDecimal(row[nendoA + "10"]); }
                if (row.Table.Columns.Contains(nendoA + "11")) { nr["１１月"] = row[nendoA + "11"]; sum += Convert.ToDecimal(row[nendoA + "11"]); }
                if (row.Table.Columns.Contains(nendoA + "12")) { nr["１２月"] = row[nendoA + "12"]; sum += Convert.ToDecimal(row[nendoA + "12"]); }
                if (row.Table.Columns.Contains(nendoB + "01")) { nr["１月"] = row[nendoB + "01"]; sum += Convert.ToDecimal(row[nendoB + "01"]); }
                if (row.Table.Columns.Contains(nendoB + "02")) { nr["２月"] = row[nendoB + "02"]; sum += Convert.ToDecimal(row[nendoB + "02"]); }
                if (row.Table.Columns.Contains(nendoB + "03")) { nr["３月"] = row[nendoB + "03"]; sum += Convert.ToDecimal(row[nendoB + "03"]); } 
                 nr["年間"] = sum;
                zenzendt.Rows.Add(nr);
            }

             SumK = 0;
             level = 0;

             SumK = Convert.ToDecimal(zenzendt.Rows[3]["年間"]) == 0 ? 0 : Convert.ToDecimal(zenzendt.Rows[6]["年間"]) / Convert.ToDecimal(zenzendt.Rows[3]["年間"]) * 100;

             level = Com.GetLevel(SumK);

             zenzendt.Rows[8]["年間"] = SumK;
             zenzendt.Rows[9]["年間"] = level;

             dataGridView3.DataSource = zenzendt;


            #endregion
        }

        private void KeisuuYM_Load(object sender, EventArgs e)
        {
            for (int i = 0; i <= 13; i++)
            {
                if (i == 0)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }

                if (i == 13)
                {
                    dataGridView1.Columns[i].Width = 100;
                    dataGridView2.Columns[i].Width = 100;
                    dataGridView3.Columns[i].Width = 100;
                }
                else
                {
                    dataGridView1.Columns[i].Width = 85;
                    dataGridView2.Columns[i].Width = 85;
                    dataGridView3.Columns[i].Width = 85;
                }


                dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                
                dataGridView2.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView2.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                
                dataGridView3.Columns[i].DefaultCellStyle.Format = "#,0";
                dataGridView3.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dataGridView1.Rows[3].DefaultCellStyle.BackColor = Color.PaleGreen;
            dataGridView1.Rows[6].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView1.Rows[8].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView1.Rows[8].DefaultCellStyle.Format = "0.00\'%\'";

            dataGridView2.Rows[3].DefaultCellStyle.BackColor = Color.PaleGreen;
            dataGridView2.Rows[6].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView2.Rows[8].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView2.Rows[8].DefaultCellStyle.Format = "0.00\'%\'";

            dataGridView3.Rows[3].DefaultCellStyle.BackColor = Color.PaleGreen;
            dataGridView3.Rows[6].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView3.Rows[8].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView3.Rows[8].DefaultCellStyle.Format = "0.00\'%\'";
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
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

                Com.GetLevelDisp(e, val);
            }
        }
    }
}
