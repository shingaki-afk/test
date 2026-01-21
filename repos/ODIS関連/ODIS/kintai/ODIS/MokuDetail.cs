using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class MokuDetail : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();
        private DataTable wkdt = new DataTable();

        public MokuDetail(string bumon, string genba)
        {
            InitializeComponent();

            GetData(bumon, genba);
        }

        private void GetData(string bumon, string genba)
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";
            //テーブルクリア
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * from dbo.目論見更新データ取得 where 部門ＣＤ like '%" + bumon + "%' and 現場ＣＤ like '%" + genba + "%'";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[2].Width = 90;
            dataGridView1.Columns[3].Width = 90;
            dataGridView1.Columns[4].Width = 90;
            dataGridView1.Columns[5].Width = 90;
            dataGridView1.Columns[6].Width = 90;
            dataGridView1.Columns[7].Width = 90;
            dataGridView1.Columns[8].Width = 90;
            dataGridView1.Columns[9].Width = 60;
            dataGridView1.Columns[10].Width = 60;

            //ヘッダーの中央表示
            dataGridView1.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[9].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[10].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //三桁区切り表示
            dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["計数"].DefaultCellStyle.Format = "0.00\'%\'";//計数

            //表示位置
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //読み取り専用
            this.dataGridView1.Columns["年月"].ReadOnly = true;
            this.dataGridView1.Columns["売上"].ReadOnly = true;
            this.dataGridView1.Columns["経費"].ReadOnly = true;
            this.dataGridView1.Columns["利益"].ReadOnly = true;
            this.dataGridView1.Columns["計数"].ReadOnly = true;
            this.dataGridView1.Columns["評価"].ReadOnly = true;

            //非表示
            dataGridView1.Columns["部門ＣＤ"].Visible = false;
            dataGridView1.Columns["現場ＣＤ"].Visible = false;
            dataGridView1.Columns["部門名"].Visible = false;
            dataGridView1.Columns["現場名"].Visible = false;
            dataGridView1.Columns["最終更新日"].Visible = false;

            dataGridView1.Columns["年月"].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns["売上"].DefaultCellStyle.BackColor = Color.PaleGreen;
            dataGridView1.Columns["経費"].DefaultCellStyle.BackColor = Color.Khaki;
            dataGridView1.Columns["利益"].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView1.Columns["計数"].DefaultCellStyle.BackColor = Color.LightGray;

            this.label1.Text = dt.Rows[1]["部門名"].ToString();
            this.label2.Text = dt.Rows[1]["現場名"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                MessageBox.Show("更新しました。画面再表示で反映されます");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }
        }

        public static void GetLevelDispMoku(DataGridViewCellFormattingEventArgs e, decimal val)
        {

            if (e.ColumnIndex == 10)
            {
                if (val == 0)
                {
                    e.Value = "-";
                }
                else if (val == 1)
                {
                    e.Value = "Ｅ";
                    e.CellStyle.BackColor = Color.Black;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (val == 2)
                {
                    e.Value = "Ｄ";
                    e.CellStyle.BackColor = Color.Gray;
                }
                else if (val == 3)
                {
                    e.Value = "Ｃ";
                    e.CellStyle.BackColor = Color.Crimson;
                }
                else if (val == 4)
                {
                    e.Value = "Ｂ";
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (val == 5)
                {
                    e.Value = "Ａ";
                    e.CellStyle.BackColor = Color.CornflowerBlue;
                }
                else if (val == 6)
                {
                    e.Value = "Ｓ";
                    e.CellStyle.BackColor = Color.LawnGreen;
                }
                else if (val == 7)
                {
                    e.Value = "ＳＳ";
                    e.CellStyle.BackColor = Color.Green;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (val == 8)
                {
                    e.Value = "ＳＳＳ";
                    e.CellStyle.BackColor = Color.Indigo;
                    e.CellStyle.ForeColor = Color.White;
                }
                else
                {
                    e.Value = "Error";
                }

                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;


            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 0)
                {
                    e.CellStyle.ForeColor = Color.Red;
                }

                //評価
                GetLevelDispMoku(e, val);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //固定、臨時、物品売上が変更
            decimal kotei = Convert.ToDecimal(dataGridView1[1, e.RowIndex].Value);
            decimal ringi = Convert.ToDecimal(dataGridView1[2, e.RowIndex].Value);
            decimal butsu = Convert.ToDecimal(dataGridView1[3, e.RowIndex].Value);
            decimal uri = kotei + ringi + butsu;

            decimal syokeihi = Convert.ToDecimal(dataGridView1[5, e.RowIndex].Value);
            decimal zinkenhi = Convert.ToDecimal(dataGridView1[6, e.RowIndex].Value);
            decimal keihi = syokeihi + zinkenhi;

            decimal rieki = uri - keihi;
            decimal keisu = uri == 0 ? 0 : keihi / uri * 100;
            int hyouka = 0;
            //売上
            dataGridView1[4, e.RowIndex].Value = uri;
            //経費
            dataGridView1[7, e.RowIndex].Value = keihi;
            //利益
            dataGridView1[8, e.RowIndex].Value = rieki;
            //計数
            dataGridView1[9, e.RowIndex].Value = keisu;

            //評価
            if (keisu == 0) hyouka = 0;
            else if (keisu > 100) hyouka = 1;
            else if (keisu > 90) hyouka = 2;
            else if (keisu > 85) hyouka = 3;
            else if (keisu > 80) hyouka = 4;
            else if (keisu > 70) hyouka = 5;
            else if (keisu > 60) hyouka = 6;
            else if (keisu > 50) hyouka = 7;
            else hyouka = 8;

            dataGridView1[10, e.RowIndex].Value = hyouka;
        }



        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //全角は半角に
            String result = Strings.StrConv(e.FormattedValue.ToString(), VbStrConv.Narrow).Trim();
            result = result.Replace(",", "");
            ////空白チェック
            if (result == "")
            {
                MessageBox.Show("空白はだめです");
                e.Cancel = true;
                return;
            }

            //数値以外除去 
            Regex regex = new Regex(@"^[0-9]+$");
            if (!regex.IsMatch(result.Replace(",","")))
            {
                MessageBox.Show("数字以外は入力不可です");
                //this.errorProvider1.SetError(cont, "数字以外は入力しないでください");
                e.Cancel = true;
            }


            //全角入力対応
            //dataGridView1.CurrentCell.Value = result;
            //dataGridView1[e.RowIndex][e.ColumnIndex] = result;


            //DataGridView dgv = (DataGridView)sender;

            ////新しい行のセルでなく、セルの内容が変更されている時だけ検証する
            //if (e.RowIndex == dgv.NewRowIndex || !dgv.IsCurrentCellDirty)
            //{
            //    return;
            //}

            //if (dgv.Columns[e.ColumnIndex].Name == "有効期限" && e.FormattedValue.ToString() == "")
            //{
            //    //行にエラーテキストを設定
            //    dgv.Rows[e.RowIndex].ErrorText = "値が入力されていません。";
            //    //入力した値をキャンセルして元に戻すには、次のようにする
            //    //dgv.CancelEdit();
            //    //キャンセルする
            //    e.Cancel = true;
            //}








        }
    }
}
