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
    public partial class ZigyouUp : Form
    {
        private SqlConnection Cn;
        //private SqlCommand Cmd;


        //サンエー控除対象者一覧
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();


        //浦添駐車場対象者一覧
        private SqlDataAdapter urada;
        private SqlCommandBuilder uracb;
        private DataTable uradt = new DataTable();

        //通信手当
        //private SqlDataAdapter tuuda;
        //private SqlCommandBuilder tuucb;
        private DataTable tuudt = new DataTable();

        public ZigyouUp()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //左側メインの更新
            GetMainData();

            //右側詳細の更新
            GetData();


            //サンエー
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            //dataGridView1.Columns[5].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView1.Columns[7].ReadOnly = true;
            dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.AntiqueWhite;

            dataGridView5.Columns[0].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView5.Columns[1].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView5.Columns[2].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView5.Columns[3].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView5.Columns[4].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            //dataGridView5.Columns[5].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView5.Columns[6].ReadOnly = true;
            dataGridView5.Columns[6].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView5.Columns[7].ReadOnly = true;
            dataGridView5.Columns[7].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView5.Columns[8].ReadOnly = true;
            dataGridView5.Columns[8].DefaultCellStyle.BackColor = Color.WhiteSmoke;

            Com.InHistory("46_駐車場控除入力", "", "");

        }

        private void GetMainData()
        {
            //リセット!
            uradt.Clear();
            dt.Clear();

            //dataGridView5.DataSource = null;
            //dataGridView1.DataSource = null;

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            //浦添駐車場
            string urasql = "select 項目, 内容, 社員番号, 氏名, 金額, 備考, (select 組織名 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 部門, (select 退職年月日 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 退職年月日, a.管理No from dbo.固定控除 a where 内容 = '固01_浦添駐車場代' order by No";
            urada = new SqlDataAdapter(urasql, Cn);
            uracb = new SqlCommandBuilder(urada);
            urada.Fill(uradt);
            dataGridView5.DataSource = uradt;
            GetDataUra();

            //サンエー駐車場
            string sql = "select a.社員番号, a.氏名, (select b.現場名 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 現場名, CEILING(金額/1.1) as [請求額(税抜)],　金額 as [控除額(税込)], 備考, (select b.退職年月日 from dbo.社員基本情報 b where b.社員番号 = a.社員番号) as 退職年月日, a.管理No, a.項目, a.内容, No from dbo.固定控除 a where 内容 = '固04_サンエー駐車場代' order by No";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //サンエー更新ボタン

            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                GetMainData();
                GetData();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            //浦添駐車場更新ボタン
            try
            {
                //データ更新
                urada.Update(uradt);

                //データ更新終了をDataTableに伝える
                uradt.AcceptChanges();

                GetMainData();
                GetDataUra();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }


        }


        private void GetData()
        {
            //サンエー駐車場情報取得

            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.sサンエー請求書額とZeeMの比較";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt2);

                        Cmd.CommandText = "select a.社員番号, a.氏名, a.金額 as [控除額(税込)], b.固定他１ from dbo.固定控除 a left join dbo.k固定給一覧 b on a.社員番号 = b.社員番号 where a.金額 <> b.固定他１ and a.内容 = '固04_サンエー駐車場代'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt3);

                        Cmd.CommandText = "select a.社員番号, a.氏名, b.退職年月日 from dbo.固定控除 a left join dbo.社員基本情報 b on a.社員番号 = b.社員番号 where 退職年月日 is not null and a.内容 = '固04_サンエー駐車場代'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt4);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView2.DataSource = dt2;
            dataGridView3.DataSource = dt3;
            dataGridView4.DataSource = dt4;
        }



        private void GetDataUra()
        {
            //浦添駐車場詳細更新

            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        //Cmd.CommandText = "select sum(金額) as ODIS側, sum(b.固定他１) as ZeeM側, sum(金額)-sum(b.固定他１) as 差額 from dbo.固定控除 a left join dbo.k固定給一覧 b on a.社員番号 = b.社員番号 where a.内容 = '固01_浦添駐車場代'";
                        Cmd.CommandText = " select * from dbo.u浦添駐車場チェック";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt2);

                        Cmd.CommandText = "select a.社員番号, a.氏名, a.金額, b.固定他１ from dbo.固定控除 a left join dbo.k固定給一覧 b on a.社員番号 = b.社員番号 where a.金額 <> b.固定他１ and a.内容 = '固01_浦添駐車場代'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt3);

                        Cmd.CommandText = "select  a.社員番号, a.氏名, b.退職年月日 from dbo.固定控除 a left join dbo.社員基本情報 b on a.社員番号 = b.社員番号 where b.退職年月日 is not null and a.内容 = '固01_浦添駐車場代'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt4);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView6.DataSource = dt2;
            dataGridView7.DataSource = dt3;
            dataGridView8.DataSource = dt4;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //decimalに変換できるか確かめる
            double d;
            if (double.TryParse(textBox1.Text, out d))
            {
                label8.Text = Math.Floor(d * 1.1).ToString();
            }
            else
            {
                MessageBox.Show("数値を入力ください。");
            }
        }
    }
}
