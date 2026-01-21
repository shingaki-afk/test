using Npgsql;
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
    public partial class UriageKeiyaku : Form
    {
        public UriageKeiyaku()
        {
            InitializeComponent();
            
            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;


            GetNewData();

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
            //コンボボックス無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            DataTable dt = new DataTable();
            int nRet;



            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter("select * from kpcp01.\"売上固定データ取得\" where 工事名 like '%" + textBox1.Text + "%' or 契約名 like '%" + textBox1.Text + "%' or 部門名 like '%" + textBox1.Text + "%'" , conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            dataGridView1.DataSource = dt;

            //全て入力した後に列幅を自動調節する
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //dataGridView1.Columns[0].Width = 50;
            //dataGridView1.Columns[1].Width = 100;
            //dataGridView1.Columns[2].Width = 50;
            //dataGridView1.Columns[3].Width = 200;
            //dataGridView1.Columns[4].Width = 200;
            //dataGridView1.Columns[5].Width = 60;
            //dataGridView1.Columns[6].Width = 50;

            //月別タブの表示設定
            for (int i = 7; i < dt.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //列幅指定
                if (i == 19)
                {
                    dataGridView1.Columns[i].Width = 100;
                }
                else
                {
                    dataGridView1.Columns[i].Width = 50;
                    //三桁区切り表示
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                }

                //ヘッダーの中央表示
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;

            Com.InHistory("17_売上契約", textBox1.Text, "");
        }

        private void button1_Click(object sender, EventArgs e)
        {


            GetData();


        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) GetData();
        }
    }
}
