using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ODIS
{
    public partial class ExData : Form
    {

        private TargetDays td = new TargetDays();
        private string[] sArr;

        public ExData()
        {
            InitializeComponent();
        }

        public ExData(string[] sArray)
        {

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            sArr = sArray;

            name.Text = sArr[0];
            number.Text = sArr[1];
            syokusyu.Text = sArr[2];
            genba.Text = sArr[3];
            kubun.Text = sArr[4];
            nyuusya.Text = sArr[5];
            taisyoku.Text = sArr[6];

            //string st = "";

            //for (int i = 0; i > -12; i--)
            //{
            //    st = td.StartYMD.AddMonths(i).ToString("yyyy年MM月給与 ") +
            //                td.StartYMD.AddMonths(i - 1).ToString("　(M/d～") +
            //                td.EndYMD.AddMonths(i - 1).ToString("M/d)");
            //}

            //GetData(td.StartYMD);

            GetExData();

        }

        private void GetExData()
        {
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
                        Cmd.CommandText = "勤怠情報取得";

                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                        Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["社員番号"].Value = sArr[1];

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

            dataGridView1.DataSource = dt;

            ////列幅設定

            if (dt.Rows.Count == 0) return;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i < 3)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //dataGridView1.Columns[i].Width = 50;
                }
                else if (i < 16)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //dataGridView1.Columns[i].Width = 50;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //dataGridView1.Columns[i].Width = 40;
                }

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        public void GetData(DateTime datet)
        {
            SqlConnection Cn;
            SqlCommand Cmd;

            DataTable dt = new DataTable();
            DataSet ds = new DataSet();

            DataTable dt0 = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
            DataTable dt5 = new DataTable();
            DataTable dt6 = new DataTable();
            DataTable dt7 = new DataTable();
            DataTable dt8 = new DataTable();
            DataTable dt9 = new DataTable();
            DataTable dt10 = new DataTable();
            DataTable dt11 = new DataTable();
            //DataTable dt12 = new DataTable();

            ds.Tables.Add(dt0); dt0.TableName = "dt0";
            ds.Tables.Add(dt1); dt1.TableName = "dt1";
            ds.Tables.Add(dt2); dt2.TableName = "dt2";
            ds.Tables.Add(dt3); dt3.TableName = "dt3";
            ds.Tables.Add(dt4); dt4.TableName = "dt4";
            ds.Tables.Add(dt5); dt5.TableName = "dt5";
            ds.Tables.Add(dt6); dt6.TableName = "dt6";
            ds.Tables.Add(dt7); dt7.TableName = "dt7";
            ds.Tables.Add(dt8); dt8.TableName = "dt8";
            ds.Tables.Add(dt9); dt9.TableName = "dt9";
            ds.Tables.Add(dt10); dt10.TableName = "dt10";
            ds.Tables.Add(dt11); dt11.TableName = "dt11";
            //ds.Tables.Add(dt12); dt12.TableName = "dt12";

            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "過去勤怠";

                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                        Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("処理年", SqlDbType.VarChar));
                        Cmd.Parameters["処理年"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("処理月", SqlDbType.VarChar));
                        Cmd.Parameters["処理月"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("基準日", SqlDbType.VarChar));
                        Cmd.Parameters["基準日"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("給与支給区分", SqlDbType.VarChar));
                        Cmd.Parameters["給与支給区分"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("勤怠対象開始年月日", SqlDbType.VarChar));
                        Cmd.Parameters["勤怠対象開始年月日"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["社員番号"].Value = sArr[1];
                        Cmd.Parameters["給与支給区分"].Value = sArr[7];

                        //for (int i = 11; i >= 0; i--)
                        for (int i = 0; i < 12; i++)
                        {
                            Cmd.Parameters["処理年"].Value = datet.AddMonths(i*-1).Year.ToString();
                            Cmd.Parameters["処理月"].Value = String.Format("{0:00}", datet.AddMonths(i * -1).Month);
                            Cmd.Parameters["基準日"].Value = datet.AddMonths(i * -1).ToString("yyyy/MM/dd");
                            Cmd.Parameters["勤怠対象開始年月日"].Value = datet.AddMonths(i * -1).ToString("yyyy/MM/dd");

                            da = new SqlDataAdapter(Cmd);
                            da.Fill(ds.Tables["dt" + i.ToString()]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            string dtcolName="";

            dt = ds.Tables["dt0"].Copy();

            dtcolName = dt.Columns[3].ColumnName;
            dt.Columns.Remove(dtcolName);

            dtcolName = dt.Columns[2].ColumnName;
            dt.Columns.Remove(dtcolName);

            dtcolName = dt.Columns[0].ColumnName;
            dt.Columns.Remove(dtcolName);

            dt.Columns[0].ColumnName = "項目名";

            //列追加
            for (int i = 11; i >= 0; i--)
            {
                string colN = datet.AddMonths(i * -1).ToString("yy年MM月");
                dt.Columns.Add(colN);

                for (int i2 = 0; i2 < ds.Tables["dt"+i.ToString()].Rows.Count; i2++)
                {
                    dt.Rows[i2][colN] = Convert.ToDecimal(ds.Tables["dt" + i.ToString()].Rows[i2][3]).ToString("#,##0.##;-#,##0.##;#");
                }
            }

            dataGridView1.DataSource = dt;

            //列幅設定
            
            //項目名
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter; 
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; 

            //各月
            for (int i = 1; i <= 12; i++)
            {
                dataGridView1.Columns[i].Width = 60;
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter; 
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; 
            }

            //データグリッドビューの背景色変更
            this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.White;
            this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);
            
            //ソート不可対応
            //foreach (DataGridViewColumn c in dataGridView1.Columns)
            //    c.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private DataTable replaceDataTable(DataTable dt)
        {
            DataTable retDt = new DataTable();
            DataRow row = null;
            try
            {
                // 戻り値のDataTable作成
                retDt.Columns.Add((string)dt.Columns[0].ColumnName, typeof(String));
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    retDt.Columns.Add((string)dt.Rows[j].ItemArray[0], typeof(String));
                }

                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    row = retDt.NewRow();
                    row[(string)dt.Columns[0].ColumnName] = dt.Columns[i].ColumnName;
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        row[(string)dt.Rows[j].ItemArray[0]] = dt.Rows[j].ItemArray[i];
                    }

                    retDt.Rows.Add(row);
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
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val == 0)
                {
                    //e.CellStyle.ForeColor = Color.Gray;
                    e.Value = null;
                }
            }
        }
    }
}
