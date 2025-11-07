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
    public partial class RiyuuKanriZin : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public RiyuuKanriZin()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            GetData();
        }

        private void GetData()
        {
            dataGridView1.DataSource = "";
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * , (select 退職年月日 from s社員基本情報 b where a.社員番号 = b.社員番号) as 退職年月日 from dbo.z人事情報エラーチェック認識済 a";
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
