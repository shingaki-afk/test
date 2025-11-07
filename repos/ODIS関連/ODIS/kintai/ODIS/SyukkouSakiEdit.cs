using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SyukkouSakiEdit : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public SyukkouSakiEdit()
        {
            InitializeComponent();

            GetData();

            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
        }

        private void GetData()
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //テーブルクリア
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select 組織CD, 組織名, 現場CD, 現場名, 八重山 from dbo.担当テーブル";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
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
