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
    public partial class YosanTotalSyokusyu_Zenkihi : Form
    {
        private string yms = "";
        private string yme = "";

        private string zi = "";

        public YosanTotalSyokusyu_Zenkihi()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dgvyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //TODO 毎年追加
            cbki.Items.Add("53期(2024)");
            cbki.Items.Add("54期(2025)");

            //TODO とりあえず今はテストで-2にしている
            cbki.SelectedIndex = cbki.Items.Count - 1;

            IniSet();
            FirstSet();

            cbziku.Items.Add("実績");
            cbziku.Items.Add("予算");
            //cbziku.Items.Add("前年");
            cbziku.SelectedIndex = 0;

            cbhikaku.Items.Add("実績");
            cbhikaku.Items.Add("予算");
            cbhikaku.Items.Add("前年");
            cbhikaku.SelectedIndex = 1;

            cbtani.Items.Add("円");
            cbtani.Items.Add("千円");
            cbtani.SelectedIndex = 1;

            GetData();


            Com.InHistory("54_予算集計職種別_差額表示", "", "");
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

            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "y予算実績前期職種別_合計比較";

                        Cmd.Parameters.Add(new SqlParameter("yms", SqlDbType.VarChar));
                        Cmd.Parameters["yms"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("yme", SqlDbType.VarChar));
                        Cmd.Parameters["yme"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("zi", SqlDbType.VarChar));
                        Cmd.Parameters["zi"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("flg", SqlDbType.VarChar));
                        Cmd.Parameters["flg"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("main", SqlDbType.VarChar));
                        Cmd.Parameters["main"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("sub", SqlDbType.VarChar));
                        Cmd.Parameters["sub"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("tani", SqlDbType.VarChar));
                        Cmd.Parameters["tani"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["yms"].Value = cbs.SelectedItem.ToString();
                        Cmd.Parameters["yme"].Value = cbe.SelectedItem.ToString();

                        Cmd.Parameters["zi"].Value = zi;

                        if (cbhikiate.Checked)
                        {
                            Cmd.Parameters["flg"].Value = "1";
                        }
                        else
                        {
                            Cmd.Parameters["flg"].Value = "0";
                        }


                        Cmd.Parameters["main"].Value = cbziku.SelectedItem.ToString();
                        Cmd.Parameters["sub"].Value = cbhikaku.SelectedItem.ToString();

                        if (cbtani.SelectedItem.ToString() == "円")
                        {
                            Cmd.Parameters["tani"].Value = "1";
                        }
                        else
                        {
                            Cmd.Parameters["tani"].Value = "1000";
                        }


                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }

                dgvyosan.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            //dt = Com.GetDB(sql);
            dgvyosan.DataSource = dt;


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


            //ヘッダーの中央表示
            for (int i = 0; i < dgvyosan.Columns.Count; i++)
            {
                dgvyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }


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

            if (cbhikiate.Checked)
            {
                dgvyosan.Columns["賞与"].Visible = false;
                dgvyosan.Columns["退職金"].Visible = false;
                dgvyosan.Columns["管理賞与"].Visible = false;
                dgvyosan.Columns["管理退職金"].Visible = false;
            }
            else
            {
                dgvyosan.Columns["賞与"].Visible = true;
                dgvyosan.Columns["退職金"].Visible = true;
                dgvyosan.Columns["管理賞与"].Visible = true;
                dgvyosan.Columns["管理退職金"].Visible = true;
            }
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
            if (e.Value != null && e.ColumnIndex == 2 && e.Value.ToString() == "03_差額" && e.RowIndex % 2 == 0)
            {
                dgvyosan.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PowderBlue;
            }
            else if (e.Value != null && e.ColumnIndex == 2 && e.Value.ToString() == "03_差額" && e.RowIndex % 2 != 0)
            {
                dgvyosan.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Thistle;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbziku_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbhikaku_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbtani_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                yms = cbki.SelectedItem.ToString().Substring(4, 4);
                yme = (Convert.ToInt16(yms) + 1).ToString();

                IniSet();
                FirstSet();
                GetData();
            }
        }

        private void cbzi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                zi = cbzi.SelectedItem.ToString().Substring(1, 1);

                GetData();
            }
        }

        private void cbs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void cbe_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbki.SelectedItem != null && cbzi.SelectedItem != null && cbs.SelectedItem != null && cbe.SelectedItem != null && cbziku.SelectedItem != null && cbhikaku.SelectedItem != null && cbtani.SelectedItem != null)
            {
                GetData();
            }
        }

        private void YosanTotalSyokusyu_Zenkihi_Load(object sender, EventArgs e)
        {

        }
    }
}
