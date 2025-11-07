using ODIS;
using ODIS.ODIS;
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
using System.IO;
using System.Data.OleDb;

namespace ODIS.ODIS
{
    public partial class Furikae : Form
    {
        /// <summary>
        /// 応援対象者データ
        /// </summary>
        private DataTable dt = new DataTable();

        /// <summary>
        /// 応援コンボボックスデータ
        /// </summary>
        private DataTable ComboDt = new DataTable();

        /// <summary>
        /// 対象期間インスタンス
        /// </summary>
        private TargetDays td = new TargetDays();

        //private bool checkflg = false;



        private DataTable soshikidt = new DataTable();

        private DataTable soshikisakidt = new DataTable();


        private DataTable genbadt = new DataTable();

        private Boolean dataflg = true;

        public Furikae()
        {
            InitializeComponent();
            
            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            // 選択モードを行単位での選択のみにする
            //dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //社員表示
            GetData();

            //地区
            tiku_moto.Items.Add("1　本社");
            tiku_moto.Items.Add("2　那覇");
            tiku_moto.Items.Add("3　八重山");
            tiku_moto.Items.Add("4　北部");
            tiku_moto.Items.Add("5　広域");
            tiku_moto.Items.Add("6　宮古島");
            tiku_moto.Items.Add("7　久米島");

            tiku_saki.Items.Add("1　本社");
            tiku_saki.Items.Add("2　那覇");
            tiku_saki.Items.Add("3　八重山");
            tiku_saki.Items.Add("4　北部");
            tiku_saki.Items.Add("5　広域");
            tiku_saki.Items.Add("6　宮古島");
            tiku_saki.Items.Add("7　久米島");

            //組織元一覧　
            soshikidt = Com.GetDB("select distinct 組織CD, 組織名 from dbo.担当テーブル where isnull(定員数, 0) > 0 ");

            //組織先一覧
            soshikisakidt = Com.GetDB("select distinct 組織CD, 組織名 from dbo.担当テーブル where isnull(定員数, 0) > 0 or isnull(出向flg, 0) > 0 ");
            //select distinct 現場CD, 現場名 from dbo.担当テーブル where (isnull(定員数, 0) > 0 or isnull(出向flg, 0) > 0)

            //TODO たぶんもうつかわない
            //ComboDt = Com.GetDB("select * from dbo.担当テーブル");

            GetSyukoData();
            GetSyuukei();

            Com.InHistory("34_出向応援入力", "", "");

            change.Enabled = false;
            delete.Enabled = false;
        }

        private void GetSyuukei()
        {
            //応援集計
            dataGridView2.DataSource = Com.GetDB("select * from o応援集計_部門指定('" + Program.loginbusyo + "', '" + td.StartYMD.ToString("yyyyMM") + "')");
            // td.EndYMD.ToString("yyyyMM")
        }


        private void GetData()
        {
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";

            //TODO
            if (checkBox1.Checked)
            {
                //全て表示
            }
            else
            {
                result = "where  担当区分  like '%" + Program.loginbusyo + "%' ";
            }

            //result = "where  担当区分  like '%" + Program.loginbusyo + "%' ";


            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                   result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //先頭が「and」の場合、「where」にする
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
                result = " where " + result;
            }

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    string sql = "select * from dbo.o応援情報取得func('" + td.EndYMD.ToString("yyyy") + "','" + td.EndYMD.ToString("MM") + "','" + td.EndYMD.ToString("yyyy/MM/dd") + "') " + result + " order by 地区, 組織, 現場";
                    Cmd.CommandText = sql;
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;

            dataGridView1.Columns["時給"].Visible = false;
            dataGridView1.Columns["担当区分"].Visible = false;
            dataGridView1.Columns["担当事務"].Visible = false;
            dataGridView1.Columns["reskey"].Visible = false;

            dataGridView1.Columns[0].Width = 60; //社員番号
            dataGridView1.Columns[1].Width = 120; //氏名
            dataGridView1.Columns[2].Width = 60; //地区名
            dataGridView1.Columns[3].Width = 120; //組織名
            dataGridView1.Columns[4].Width = 140; //現場名
            dataGridView1.Columns[5].Width = 40; //出向単価
            dataGridView1.Columns[6].Width = 40; //勤務時間

            dataGridView1.Columns[11].Width = 40; //0.25
            dataGridView1.Columns[12].Width = 40; //1.25
            dataGridView1.Columns[13].Width = 40; //1.35

            dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[11].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[12].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[13].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dt.Clear();
            GetData();
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            no.Text = drv[0].ToString();
            name.Text = drv[1].ToString();
            tiku_moto.SelectedItem = drv[2].ToString();
            soshiki_moto.SelectedItem = drv[3].ToString();
            genba_moto.SelectedItem = drv[4].ToString();


            tanka.Value = Convert.ToDecimal(drv[5].ToString());
            nikkin.Value = Convert.ToDecimal(drv[6].ToString());

            id.Text = "";
            change.Enabled = false;
            delete.Enabled = false;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dt.Clear();
                GetData();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //TODO
            if (genba_moto.SelectedIndex == -1)
            {
                MessageBox.Show("応援元が選択されていません。");
            }
            else if (genba_saki.SelectedIndex == -1)
            {
                MessageBox.Show("応援先が選択されていません。");
            }
            else
            {
                dataflg = false;
                InsertSK("追加");
                GetSyukoData();
                GetSyuukei();
                dataflg = true;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //TODO
            if (genba_moto.SelectedIndex == -1)
            {
                MessageBox.Show("応援元が選択されていません。");
            }
            else if (genba_saki.SelectedIndex == -1)
            {
                MessageBox.Show("応援先が選択されていません。");
            }
            else
            {
                dataflg = false;
                InsertSK("変更");
                GetSyukoData();
                GetSyuukei();
                dataflg = true;
            }
        }

        /// <summary>
        /// 応援データ登録処理
        /// </summary>
        private void InsertSK(string flg)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[o応援データ更新]";

                    Cmd.Parameters.Add(new SqlParameter("処理年月", SqlDbType.VarChar));
                    Cmd.Parameters["処理年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名", SqlDbType.VarChar));
                    Cmd.Parameters["氏名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("元組織CD", SqlDbType.Char));
                    Cmd.Parameters["元組織CD"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("元組織名", SqlDbType.VarChar));
                    Cmd.Parameters["元組織名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("元現場CD", SqlDbType.Char));
                    Cmd.Parameters["元現場CD"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("元現場名", SqlDbType.VarChar));
                    Cmd.Parameters["元現場名"].Direction = ParameterDirection.Input;


                    Cmd.Parameters.Add(new SqlParameter("先組織CD", SqlDbType.Char));
                    Cmd.Parameters["先組織CD"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("先組織名", SqlDbType.VarChar));
                    Cmd.Parameters["先組織名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("先現場CD", SqlDbType.Char));
                    Cmd.Parameters["先現場CD"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("先現場名", SqlDbType.VarChar));
                    Cmd.Parameters["先現場名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("単価", SqlDbType.Decimal));
                    Cmd.Parameters["単価"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("労働時間", SqlDbType.Decimal));
                    Cmd.Parameters["労働時間"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("振替額", SqlDbType.Decimal));
                    Cmd.Parameters["振替額"].Direction = ParameterDirection.Input;

                    //Cmd.Parameters.Add(new SqlParameter("残業時間", SqlDbType.Decimal));
                    //Cmd.Parameters["残業時間"].Direction = ParameterDirection.Input;

                    //Cmd.Parameters.Add(new SqlParameter("深夜時間", SqlDbType.Decimal));
                    //Cmd.Parameters["深夜時間"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                    Cmd.Parameters["備考"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("入力者部門", SqlDbType.VarChar));
                    Cmd.Parameters["入力者部門"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("入力者氏名", SqlDbType.VarChar));
                    Cmd.Parameters["入力者氏名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.DateTime));
                    Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("flg", SqlDbType.VarChar));
                    Cmd.Parameters["flg"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("id", SqlDbType.VarChar));
                    Cmd.Parameters["id"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["処理年月"].Value = td.StartYMD.ToString("yyyyMM");
                    Cmd.Parameters["社員番号"].Value = no.Text;
                    Cmd.Parameters["氏名"].Value = name.Text;

                    Cmd.Parameters["元組織CD"].Value = soshiki_moto.SelectedItem.ToString().Split('　')[0];
                    Cmd.Parameters["元組織名"].Value = soshiki_moto.SelectedItem.ToString().Split('　')[1];
                    Cmd.Parameters["元現場CD"].Value = genba_moto.SelectedItem.ToString().Split('　')[0];
                    Cmd.Parameters["元現場名"].Value = genba_moto.SelectedItem.ToString().Split('　')[1];

                    Cmd.Parameters["先組織CD"].Value = soshiki_saki.SelectedItem.ToString().Split('　')[0];
                    Cmd.Parameters["先組織名"].Value = soshiki_saki.SelectedItem.ToString().Split('　')[1];
                    Cmd.Parameters["先現場CD"].Value = genba_saki.SelectedItem.ToString().Split('　')[0];
                    Cmd.Parameters["先現場名"].Value = genba_saki.SelectedItem.ToString().Split('　')[1];

                    Cmd.Parameters["単価"].Value = tanka.Value;
                    Cmd.Parameters["労働時間"].Value = nikkin.Value;
                    Cmd.Parameters["振替額"].Value = hurikaegaku.Value;

                    Cmd.Parameters["備考"].Value = bikou.Text;

                    Cmd.Parameters["入力者部門"].Value = Program.loginbusyo;
                    Cmd.Parameters["入力者氏名"].Value = Program.loginname;

                    Cmd.Parameters["更新日時"].Value = DateTime.Now;
                    Cmd.Parameters["flg"].Value = flg;
                    Cmd.Parameters["id"].Value = id.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }


        }

        private void GetSyukoData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            //TODO
            //出向データ表示は　入力担当部署、入力担当名　シス管はスルー
            //シス管がinsert  
            //シス菅がupdate

            string sql = "";
            if (checkBox2.Checked)
            {
                sql = "select* from dbo.o応援データ where 処理年月 like '" + td.EndYMD.ToString("yyyyMM") + "%' order by 更新日時 desc ";
            }
            else
            {
                sql = "select* from dbo.o応援データ where 処理年月 like '" + td.EndYMD.ToString("yyyyMM") + "%' and (isnull(入力者部門,'') like '%" + Program.loginbusyo + "%' or 入力者氏名 = '" + Program.loginname + "') order by 更新日時 desc ";
            }




            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = sql;
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            dataGridView3.DataSource = dt;

            dataGridView3.Columns["元組織CD"].Visible = false;
            dataGridView3.Columns["元現場CD"].Visible = false;
            dataGridView3.Columns["先組織CD"].Visible = false;
            dataGridView3.Columns["先現場CD"].Visible = false;

            //dataGridView3.Columns["入力者部門"].Visible = false;
            //dataGridView3.Columns["入力者氏名"].Visible = false;
            //dataGridView3.Columns["更新日時"].Visible = false;
        }



        private void DeleteSyukoData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "delete from o応援データ where id = " + id.Text;
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (!dataflg) return; 

            DataGridViewRow dgr = dataGridView3.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            no.Text = drv["社員番号"].ToString();
            name.Text = drv["氏名"].ToString();

            //string tiku = drv["元組織CD"].ToString().Substring(0, 1);

            tiku_moto.SelectedIndex = tiku_moto.FindString(drv["元組織CD"].ToString().Substring(0, 1));
            soshiki_moto.SelectedIndex = soshiki_moto.FindString(drv["元組織CD"].ToString().Substring(0, 5)); 
            genba_moto.SelectedIndex = genba_moto.FindString(drv["元現場CD"].ToString().Substring(0, 5)); 

            //soshiki_moto.SelectedItem = drv["元組織CD"].ToString() + "　" + drv["元組織名"].ToString();
            //genba_moto.SelectedItem = drv["元現場CD"].ToString() + "　" + drv["元現場名"].ToString();

            tiku_saki.SelectedIndex = tiku_saki.FindString(drv["先組織CD"].ToString().Substring(0, 1));
            soshiki_saki.SelectedIndex = soshiki_saki.FindString(drv["先組織CD"].ToString().Substring(0, 5));
            genba_saki.SelectedIndex = genba_saki.FindString(drv["先現場CD"].ToString().Substring(0, 5));

            //soshiki_saki.SelectedItem = drv["先組織CD"].ToString() + "　" + drv["先組織名"].ToString();
            //genba_saki.SelectedItem = drv["先現場CD"].ToString() + "　" + drv["先現場名"].ToString();

            tanka.Text = drv["単価"].ToString();
            nikkin.Text = drv["労働時間"].ToString();
            hurikaegaku.Text = drv["振替額"].ToString();
            bikou.Text = drv["備考"].ToString();

            nyuuryokusya.Text = drv["入力者氏名"].ToString();

            id.Text = drv["id"].ToString();

            if (id.Text == "")
            {
                change.Enabled = false;
            }
            else
            { 
                change.Enabled = true;
            }

            delete.Enabled = true;
        }



        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result;
            if (nyuuryokusya.Text == Program.loginname)
            {
                result = MessageBox.Show("削除して大丈夫ですか？", "削除前確認", MessageBoxButtons.OKCancel);
            }
            else
            {
                result = MessageBox.Show("入力者ではないようですが、削除して大丈夫ですか？", "削除前確認", MessageBoxButtons.OKCancel);
            }

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                dataflg = false;
                DeleteSyukoData();
                GetSyukoData();
                GetSyuukei();
                dataflg = true;
            }



        }





        private void editsaki_Click(object sender, EventArgs e)
        {
            SyukkouSakiEdit see = new SyukkouSakiEdit();
            see.ShowDialog();
        }



        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.Equals(DBNull.Value) && e.ColumnIndex == 0)
            {
                ////セルの値により、背景色を変更する
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Lavender;
            }

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void tiku_moto_SelectedIndexChanged(object sender, EventArgs e)
        {
            soshiki_moto.Items.Clear();

            DataRow[] dr = soshikidt.Select("組織CD like '" + (tiku_moto.SelectedIndex + 1).ToString() + "%'");

            foreach (DataRow drw in dr)
            {
                soshiki_moto.Items.Add(drw["組織CD"].ToString() + "　" + drw["組織名"].ToString());
            }

            //TODO 2020/04/02 変更
            //soshiki_moto.SelectedIndex = -1;
        }

        private void soshiki_moto_SelectedIndexChanged(object sender, EventArgs e)
        {
            genba_moto.Items.Clear();
            genbadt.Clear();

            genbadt = Com.GetDB("select distinct 現場CD, 現場名 from dbo.担当テーブル where isnull(定員数,0) > 0 and 組織CD = '" + soshiki_moto.SelectedItem.ToString().Substring(0, 5) + "'");
            foreach (DataRow drw in genbadt.Rows)
            {
                genba_moto.Items.Add(drw["現場CD"].ToString() + "　" + drw["現場名"].ToString());
            }

            //genba_moto.SelectedIndex = 0;
        }

        private void Furikae_Load(object sender, EventArgs e)
        {

        }

        private void bikou_TextChanged(object sender, EventArgs e)
        {

        }

        private void tiku_saki_SelectedIndexChanged(object sender, EventArgs e)
        {
            soshiki_saki.Items.Clear();

            DataRow[] dr = soshikisakidt.Select("組織CD like '" + (tiku_saki.SelectedIndex + 1).ToString() + "%'");

            foreach (DataRow drw in dr)
            {
                soshiki_saki.Items.Add(drw["組織CD"].ToString() + "　" + drw["組織名"].ToString());
            }
        }

        private void soshiki_saki_SelectedIndexChanged(object sender, EventArgs e)
        {
            genba_saki.Items.Clear();
            genbadt.Clear();

            genbadt = Com.GetDB("select distinct 現場CD, 現場名 from dbo.担当テーブル where (isnull(定員数, 0) > 0 or isnull(出向flg, 0) > 0) and 組織CD = '" + soshiki_saki.SelectedItem.ToString().Substring(0, 5) + "'");
            foreach (DataRow drw in genbadt.Rows)
            {
                genba_saki.Items.Add(drw["現場CD"].ToString() + "　" + drw["現場名"].ToString());
            }
        }

        private void tanka_ValueChanged(object sender, EventArgs e)
        {
            hurikaegaku.Value = tanka.Value * nikkin.Value;
        }

        private void nikkin_ValueChanged(object sender, EventArgs e)
        {
            hurikaegaku.Value = tanka.Value * nikkin.Value;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            no.Text = "";
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            name.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tanka.Value = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            nikkin.Value = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            hurikaegaku.Value = 0;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            bikou.Text = "";
        }

        private void editsaki_Click_1(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            checkBox1.Enabled = false;

            GetData();

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            checkBox1.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("時給に、①社保②雇保③労災④子供拠出⑤退職引当金⑥賞与引当金⑦全友協350円⑧通勤手当を加算した値です。⑦有給分⑧福利厚生⑨研修⑩被服⑪その他　は未対応です。※全て前月値となります。");
        }

        private void dataGridView3_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //項目行はスキップ
            if (e.RowIndex < 0) return;

            if (e.ColumnIndex == 1) //2列目
            {
                //TODO
                if (e.Value.ToString() != td.StartYMD.ToString("yyyyMM"))
                {
                    //TODO
                   dataGridView3.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Gray;
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            GetSyukoData();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = extractAllCsv(@"C:\temp\出向.csv");
        }

        /// <summary>
        /// csvファイルから全データを取得してdatatableへ
        /// </summary>
        /// <param name="filePath">抽出元CSV</param>
        /// <returns></returns>
        public DataTable extractAllCsv(string filePath)
        {
            DataTable dt = new DataTable();                     //取得データを格納
            string csvDir = Path.GetDirectoryName(filePath);           //CSVファイルのあるフォルダ
            string csvFileName = Path.GetFileName(filePath);           //CSVファイルの名前

            //接続文字列
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
          + csvDir + ";Extended Properties=\"text;HDR=Yes;FMT=Delimited\"";

            OleDbConnection con = new OleDbConnection(connectionString);

            //csvファイルから取得
            string commText = "SELECT * FROM [" + csvFileName + "]";
            OleDbDataAdapter da = new OleDbDataAdapter(commText, con);

            da.Fill(dt);

            return dt;
        }

        
    }
}
