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
    public partial class TermList : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public TermList()
        {
            InitializeComponent();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "SELECT * FROM [dbo].[TermList]";

            da = new SqlDataAdapter(sql, Cn);

            cb = new SqlCommandBuilder(da);

            da.Fill(dt);

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Visible = false;
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

        public DataTable GetTermList()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();

            SqlDataAdapter da;

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "SELECT * FROM [dbo].[TermList] order by 期限日";

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            DataTable Disp = new DataTable();
            Disp.Columns.Add("残日数", typeof(string));
            Disp.Columns.Add("対象日", typeof(DateTime));
            Disp.Columns.Add("対象者", typeof(string));
            Disp.Columns.Add("区分", typeof(string));
            Disp.Columns.Add("内容", typeof(string));
            Disp.Columns.Add("登録者", typeof(string));
            Disp.Columns.Add("状況", typeof(string));
            Disp.Columns.Add("備考", typeof(string));

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = Disp.NewRow();

                DateTime d;
                TimeSpan ts;
                if (DateTime.TryParse(row["期限日"].ToString(), out d))
                {
                    nr["対象日"] = d;
                    ts = d.Subtract(DateTime.Now);
                    nr["残日数"] = ts.Days.ToString() + "日";
                }
                else
                {
                    nr["対象日"] = new DateTime(0);
                    nr["残日数"] = DBNull.Value;
                }

                nr["対象者"] = row["対象者"];
                nr["区分"] = row["区分"];
                nr["内容"] = row["内容"];
                nr["登録者"] = row["登録者"];
                nr["状況"] = row["状況"];
                nr["備考"] = row["備考"];
                Disp.Rows.Add(nr);
            }

            return Disp;
        }
    }
}
