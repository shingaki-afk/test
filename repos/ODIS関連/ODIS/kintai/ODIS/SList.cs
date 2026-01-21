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
    public partial class SList : Form
    {
        
        private string year = "";
        public SList()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 14);

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("2022");
            comboBox1.Items.Add("2023");
            comboBox1.Items.Add("2024");
            comboBox1.Items.Add("2025");

            //TODO とりあえず今はテストで-2にしている
            comboBox1.SelectedIndex = comboBox1.Items.Count - 1;

            GetData();

            Com.InHistory("67_障害者雇用状況", "", "");
        }

        private void GetData()
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            comboBox1.Enabled = false;

            label1.Text = "障害者一覧(" + year + "年4月1日～)";

            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            string y = year + "/";
            string yy = (Convert.ToInt16(year) + 1).ToString()  + "/";

            //string d = "01";
            //string[] stArray = new string[] { y + "04", y + "05", y + "06", y + "07", y + "08", y + "09", y + "10", y + "11", y + "12", yy + "01", yy + "02", yy + "03" };

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd = Cn.CreateCommand();
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "s障害雇用";

                    Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                    Cmd.Parameters.Add(new SqlParameter("yadd", SqlDbType.VarChar));

                    Cmd.Parameters["year"].Direction = ParameterDirection.Input;
                    Cmd.Parameters["yadd"].Direction = ParameterDirection.Input;

                    Cmd.Parameters["year"].Value = y;
                    Cmd.Parameters["yadd"].Value = yy;

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);

                    //foreach (string s in stArray)
                    //{
                    //    Cmd.Parameters["DateOfRecord"].Value = s + "/" + d;
                    //    da = new SqlDataAdapter(Cmd);
                    //    da.Fill(dt);
                    //}
                }
            }

            DataTable dtAll = new DataTable();
            dtAll = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                if (Convert.ToDateTime(dr[0]) > DateTime.Today.AddMonths(1))
                {
                    //未来月はnullを入れる
                    for (int i = 1; i <= 11; i++)
                    {
                        dr[i] = DBNull.Value;
                    }
                }
                else
                {
                    for (int i = 1; i <= 11; i++)
                    {
                        //小数点以下のゼロ対応
                        if (dr[i] != DBNull.Value)
                        {
                            dr[i] = Convert.ToDouble(dr[i]);
                        }
                    }
                }

                dr[0] = Convert.ToDateTime(dr[0]).Month + "月";

                dtAll.ImportRow(dr);
            }

            //DataTable wkdt = new DataTable();
            //wkdt = Com.replaceDataTable(dtAll);

            dataGridView1.DataSource = replaceDataTable(dtAll);

            for (int i = 0; i < 13; i++)
            {
                //項目名以外は右寄せ表示
                if (i == 0)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.Columns[i].Width = 250;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[i].Width = 100;

                    //三桁区切り表示
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0.#";
                }

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            DataTable dtlist = new DataTable();
            dtlist = Com.GetDB("select * from dbo.s障害一覧('" + year + "/04/01') order by 組織名, 現場名");

            dataGridView2.DataSource = dtlist;

            dataGridView2.Columns[0].Width = 70;
            dataGridView2.Columns[1].Width = 100;
            dataGridView2.Columns[2].Width = 80;
            dataGridView2.Columns[3].Width = 80;
            dataGridView2.Columns[4].Width = 80;
            dataGridView2.Columns[5].Width = 90;
            dataGridView2.Columns[6].Width = 90;
            dataGridView2.Columns[7].Width = 90;
            dataGridView2.Columns[8].Width = 200;
            dataGridView2.Columns[9].Width = 50;
            dataGridView2.Columns[10].Width = 50;
            dataGridView2.Columns[11].Width = 50;
            dataGridView2.Columns[12].Width = 50;
            dataGridView2.Columns[13].Width = 50;
            dataGridView2.Columns[14].Width = 80;
            dataGridView2.Columns[15].Width = 300;

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            comboBox1.Enabled = true;
        }

        private DataTable replaceDataTable(DataTable dt)
        {
            DataTable retDt = new DataTable();
            DataRow row = null;
            try
            {
                // 戻り値のDataTable作成
                //retDt.Columns.Add((string)dt.Columns[0].ColumnName, typeof(String));
                retDt.Columns.Add("区　分", typeof(String));

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    retDt.Columns.Add((string)dt.Rows[j].ItemArray[0], typeof(String));
                }

                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    if (dt.Columns[i].ColumnName != "地区コード")
                    {
                        row = retDt.NewRow();
                        //row[(string)dt.Columns[0].ColumnName] = dt.Columns[i].ColumnName;
                        row["区　分"] = dt.Columns[i].ColumnName;

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            row[(string)dt.Rows[j].ItemArray[0]] = dt.Rows[j].ItemArray[i];
                        }

                        retDt.Rows.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return retDt;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.NavajoWhite;
            dataGridView1.Rows[1].DefaultCellStyle.BackColor = Color.NavajoWhite;
            dataGridView1.Rows[2].DefaultCellStyle.BackColor = Color.NavajoWhite; 
            
            dataGridView1.Rows[3].DefaultCellStyle.BackColor = Color.Khaki;

            dataGridView1.Rows[4].DefaultCellStyle.BackColor = Color.MistyRose;
            dataGridView1.Rows[5].DefaultCellStyle.BackColor = Color.MistyRose;
            dataGridView1.Rows[6].DefaultCellStyle.BackColor = Color.MistyRose;
            dataGridView1.Rows[7].DefaultCellStyle.BackColor = Color.MistyRose;
            dataGridView1.Rows[8].DefaultCellStyle.BackColor = Color.MistyRose;

            dataGridView1.Rows[9].DefaultCellStyle.BackColor = Color.Khaki;

            dataGridView1.Rows[10].DefaultCellStyle.BackColor = Color.LightSteelBlue;
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            ////セルの列を確認
            //DateTime val = DateTime.Now;

            //if (year != "2025") return;

            //if (e.ColumnIndex == 6 && DateTime.TryParse(e.Value.ToString(), out val))
            //{
            //    //入社
            //    //セルの値により、背景色を変更する
            //    if (Convert.ToDateTime(e.Value) >= Convert.ToDateTime(year + "/04/01"))
            //    {
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //    }
            //}
            //else if (e.ColumnIndex == 7 && DateTime.TryParse(e.Value.ToString(), out val))
            //{
            //    //退社
            //    //セルの値により、背景色を変更する
            //    if (Convert.ToDateTime(e.Value) >= Convert.ToDateTime(year + "/04/01"))
            //    {
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.DarkGray;
            //    }
            //}
        }


        private void GetKeireki(string no)
        {
            //経歴情報
            DataTable dtkeireki = new DataTable();
            dtkeireki = Com.GetDB("select * from dbo.k雇用給与変更履歴表示('" + no + "') order by 適用開始日");

            dataGridView3.DataSource = dtkeireki;

            dataGridView3.Columns[6].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[7].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[8].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[9].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[10].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[11].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[12].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[13].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[14].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[15].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[16].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[17].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[18].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[19].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[20].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[21].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[22].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[23].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[24].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[25].DefaultCellStyle.Format = "#,0";
            dataGridView3.Columns[26].DefaultCellStyle.Format = "#,0";

            dataGridView3.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[21].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[24].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView3.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            GetKeireki(drv[0].ToString());
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0) return;　//最初行とばし
            if (e.ColumnIndex == 0) return; //日付とばし
            if (e.ColumnIndex == 1) return; //年齢とばし

            if (Convert.ToString(dataGridView3.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value) != Convert.ToString(e.Value))
            {
                //前がnullで後が0はとばし
                if (dataGridView3.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.Equals(DBNull.Value)) return;

                e.CellStyle.BackColor = Color.SpringGreen;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                year = comboBox1.SelectedItem.ToString();
                GetData();
            }
        }
    }
}
