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
    public partial class Kourei : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public Kourei()
        {
            InitializeComponent();
            GetData();

            //TODO
            //更新切れor切れ予定の方の背景色変更
            //バリデート設定

            Com.InHistory("91_単年契約情報管理", "", "");
        }

        private void GetData()
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";
            //テーブルクリア
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * ";
            sql += ", (select 氏名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 氏名 ";
            sql += ", (select 地区名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 地区名 ";
            sql += ", (select 組織名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 組織名 ";
            sql += ", (select 現場名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 現場名 ";
            sql += ", (select 退職年月日 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 退職年月日 ";
            sql += "from dbo.t単年契約社員情報 a ";

            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);

            dataGridView1.DataSource = dt;
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[4].ReadOnly = true;
            dataGridView1.Columns[5].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.White;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.White;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.PowderBlue;
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.PowderBlue; //LightGray
            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.PowderBlue;
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.PowderBlue;
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.PowderBlue;

            dataGridView1[0, dataGridView1.Rows.Count - 1].ReadOnly = false;
            dataGridView1[0, dataGridView1.Rows.Count - 1].Style.BackColor = Color.White;

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
        }


    }
}
