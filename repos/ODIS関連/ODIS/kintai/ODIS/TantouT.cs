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
    public partial class TantouT : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public TantouT()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            GetData();

            Com.InHistory("95_担当テーブル設定", "", "");

        }

        private void GetData()
        {
            //dataGridView1.DataSource = null;
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "";
            if (checkBox1.Checked)
            {
                sql = "SELECT * FROM [dbo].[担当テーブル] ";
            }
            else
            {
                sql = "SELECT * FROM [dbo].[担当テーブル] where isnull(契約終了日,'') = '' or isnull(契約終了日,'') > " + DateTime.Now.ToString("yyyyMM");
            }


            da = new SqlDataAdapter(sql, Cn);

            cb = new SqlCommandBuilder(da);

            da.Fill(dt);

            dataGridView1.DataSource = dt;

            label1.Text = dt.Rows.Count.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.Value == null) return;

            //セルの列を確認
            int val = 0;

            if (e.ColumnIndex == 20 && int.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < Convert.ToInt32(DateTime.Now.ToString("yyyyMM")))
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Gray;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
