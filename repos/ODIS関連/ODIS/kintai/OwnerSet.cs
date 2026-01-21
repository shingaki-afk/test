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

namespace ODIS
{
    public partial class OwnerSet : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public OwnerSet()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            GetNonSetData();
            GetData();
        }

        private void GetNonSetData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable nondt = new DataTable();
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "select * from dbo.担当未設定";
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(nondt);
                }
            }

            dataGridView2.DataSource = nondt;
        }

        private void GetData()
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //テーブルクリア
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * from dbo.担当テーブル";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }

            GetData();
        }
    }
}
