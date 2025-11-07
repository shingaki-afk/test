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
    public partial class BillKan : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public BillKan()
        {
            InitializeComponent();
            //フォームを最大化
            this.WindowState = FormWindowState.Maximized; 

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //
            GetNewData();

            GetData();


            Com.InHistory("632_ビル管登録現場_更新", "", "");
        }

        private void GetNewData()
        {
            //削除
            DataTable dtdel = new DataTable();
            dtdel = Com.GetDB("delete from dbo.契約固定");

            DataTable dt = new DataTable();
            dt = Com.GetPosDB("select * from kpcp01.売上固定データ取得");

            using (var bulkCopy = new SqlBulkCopy(ODIS.Com.SQLConstr))
            {
                bulkCopy.DestinationTableName = "契約固定"; //dt.TableName; // テーブル名をSqlBulkCopyに教える
                bulkCopy.WriteToServer(dt); // bulkCopy実行
            }
        }

        private void GetData()
        {
            dataGridView1.DataSource = "";
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select ビル管登録現場, 社員番号, (select 氏名 from dbo.社員基本情報 b where a.社員番号 = b.社員番号) as 氏名, 工事コード, 工事枝コード, 連番, (select 契約名 from dbo.契約固定 c where a.工事コード = c.工事コード and a.工事枝コード = c.工事枝コード and a.連番 = c.連番) as 契約名, 適用開始日 from dbo.ビル管登録現場 a";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Width = 300; //ビル管登録現場
            dataGridView1.Columns[1].Width = 100; //社員番号
            dataGridView1.Columns[2].Width = 100; //氏名
            dataGridView1.Columns[3].Width = 100; //工事コード
            dataGridView1.Columns[4].Width = 100; //工事枝コード
            dataGridView1.Columns[5].Width = 100; //連番
            dataGridView1.Columns[6].Width = 300; //契約名
            dataGridView1.Columns[7].Width = 100; //契約開始日

            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;

            //インデックス0の列のセルの背景色を水色にする
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.Beige;
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

            MessageBox.Show("更新しましたー。");
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value != null && e.Value.ToString() == "")
            {
                e.CellStyle.BackColor = Color.Red;
            }
        }

        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("test");
        }
    }
}
