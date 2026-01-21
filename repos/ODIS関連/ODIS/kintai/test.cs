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

namespace ODIS
{
    public partial class test : Form
    {
        private SqlConnection Cn;
        //private SqlCommand Cmd;
        private SqlDataAdapter da;
        //private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();
        public test()
        {
            InitializeComponent();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * from dbo.事業登録";

            da = new SqlDataAdapter(sql, Cn);

            da.Fill(dt);

            //try
            //{
            //    using (Cn = new SqlConnection(Com.SQLConstr))
            //    {
            //        using (Cmd = Cn.CreateCommand())
            //        {
            //            string sql = "select * from dbo.事業登録";
            //            Cmd.CommandText = sql;
            //            da = new SqlDataAdapter(Cmd);

            //            cb = new SqlCommandBuilder(da);

            //            da.Fill(dt);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("エラー" + ex.ToString());
            //    throw;
            //}

            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //データ更新
            da.Update(dt);

            //データ更新終了をDataTableに伝える
            dt.AcceptChanges(); 
        }
    }
}
