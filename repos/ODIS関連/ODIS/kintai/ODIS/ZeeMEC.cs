using C1.C1Excel;
using ODIS;
using ODIS.ODIS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class ZeeMEC : Form
    {
        /// <summary>
        /// 共通クラスのインスタンス
        /// </summary>
        private Common co = new Common();

        /// <summary>
        /// 対象期間インスタンス
        /// </summary>
        private TargetDays td = new TargetDays();

        /// <summary>
        /// 改行コード
        /// </summary>
        public string nl = Environment.NewLine;

        public ZeeMEC()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //勤怠入力の未提出数出力
            GetKintaiInfo();

            //給与未計算対象者取得
            GetMikeisan();

            //給与明細備考
            //GetMsg();


            //if (Program.loginname == "喜屋武　大祐")
            //{
            //    button28.Enabled = true;
            //}

            Com.InHistory("44_給与計算画面", "", "");
        }

        private void GetMikeisan()
        {
            DataTable dt = new DataTable();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[給与計算状況]";

                        Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                        Cmd.Parameters["year"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("month", SqlDbType.VarChar));
                        Cmd.Parameters["month"].Direction = ParameterDirection.Input;

                        TargetDays td = new TargetDays();
                        Cmd.Parameters["year"].Value = td.StartYMD.AddMonths(1).Year.ToString();
                        Cmd.Parameters["month"].Value = td.StartYMD.AddMonths(1).ToString("MM");

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

        }

        //勤怠入力の未提出数出力
        private void GetKintaiInfo()
        {
            DataTable dt = new DataTable();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.勤怠入力状況取得 ";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            int naha = 0;
            int yae = 0;
            int hoku = 0;

            foreach (DataRow row in dt.Rows)
            {
                if(row["入力有無"].Equals(DBNull.Value))
                {
                    if (row["地区名"].ToString() == "八重山")
                    {
                        yae++;
                    }
                    else if (row["地区名"].ToString() == "北部")
                    {
                        hoku++;
                    }
                    else
                    {
                        naha++;
                    }
                }
            }
            label20.Text = naha.ToString();
            label21.Text = hoku.ToString();
            label22.Text = yae.ToString();
            label23.Text = (naha + hoku + yae).ToString();
        }
        
        //退職処理チェック
        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.t退職チェックデータ取得 order by 退職日_ODIS";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }


            DataTable Disp = new DataTable();
            Disp.Columns.Add("状況", typeof(string));
            Disp.Columns.Add("退職日", typeof(string));
            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("漢字氏名", typeof(string));
            Disp.Columns.Add("退職理由", typeof(string));

            Disp.Columns.Add("エラー", typeof(string));
            Disp.Columns.Add("支払方法", typeof(string));

            Disp.Columns.Add("地区名", typeof(string));
            Disp.Columns.Add("組織名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));
            Disp.Columns.Add("特記事項", typeof(string));
            Disp.Columns.Add("在籍年月", typeof(string));
            Disp.Columns.Add("住民", typeof(string));
            Disp.Columns.Add("社保", typeof(string));
            Disp.Columns.Add("住民税額有無", typeof(string));

            int machi = 0;
            int zumi = 0;
            int hmae = 0;
            int kmae = 0;
            int kaku = 0;
            int ecount = 0;

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = Disp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["漢字氏名"] = row["漢字氏名"];
                nr["地区名"] = row["地区名"];
                nr["組織名"] = row["組織名"];
                nr["現場名"] = row["現場名"];
                nr["在籍年月"] = row["在籍年月"];
                nr["退職日"] = Convert.ToDateTime(row["退職日_ODIS"]).ToString("yyyy/MM/dd");
                nr["支払方法"] = row["支払方法_ODIS"];
                nr["退職理由"] = row["退職理由_ODIS"];
                nr["特記事項"] = row["特記事項_ODIS"];
                nr["住民"] = row["住民税徴収方法"];
                nr["社保"] = row["社会保険徴収方法"];
                nr["住民税額有無"] = row["最高金額"];
                //状況
                string zyoukyou = "";
                if (row["マスタ更新区分"].Equals(DBNull.Value))
                {
                    if (row["勤怠入力日時"].Equals(DBNull.Value))
                    {
                        zyoukyou = "1_勤怠待";
                        machi++;
                    }
                    else
                    {
                        zyoukyou = "2_勤怠入力済";
                        zumi++;
                    }
                }
                else if (row["マスタ更新区分"].ToString() == "0" && row["処理区分"].ToString() == "0")
                {
                    zyoukyou = "3_発令前";
                    hmae++;
                }
                else if (row["マスタ更新区分"].ToString() == "1" && row["処理区分"].ToString() == "0")
                {
                    zyoukyou = "4_確定前";
                    kmae++;

                }
                else if (row["マスタ更新区分"].ToString() == "1" && row["処理区分"].ToString() == "1")
                {
                    zyoukyou = "5_確定";
                    kaku++;
                }
                else
                {
                    zyoukyou = "例外発生！";
                }

                //エラー
                string error = "";

                //入力済かチェック
                if (!row["マスタ更新区分"].Equals(DBNull.Value))
                {
                    //発令前かチェック　発令前なら、発令日と比較する必要がある

                    //退職マスタ作成伝票のさくでぎょう
                    if (Convert.ToDateTime(row["退職日_ODIS"]).ToString() != Convert.ToDateTime(row["発令日"]).ToString())
                    {
                        error += "退職日不一致    ";
                    }
                    
                    if (row["退職理由_ODIS"].ToString() != row["退職理由_ZEEM"].ToString())
                    {
                        error += "退職理由不一致   ";
                    }

                    //住民税徴収方法
                    //退職日が1/1～4/30の場合、2[給与より一括徴収]でなければならない。
                    //他は、1[徴収無し]
                    if (row["住民税徴収方法"].ToString() != "1")
                    {
                        //TODO
                    }

                }

                if (row["支払方法_ODIS"].ToString() != row["支払方法_ZEEM"].ToString())
                {
                    //前月末退職者はスルー
                    if (row["退職日_ODIS"].ToString() != td.StartYMD.AddDays(-1).ToString())
                    { 
                        error += "支払方法不一致   ";
                    }
                }

                if (error.Length > 0)
                {
                    ecount++;
                }

                nr["状況"] = zyoukyou;
                nr["エラー"] = error;

                Disp.Rows.Add(nr);
            }

            label7.Text = machi.ToString();
            label8.Text = zumi.ToString();
            label9.Text = hmae.ToString();
            label10.Text = kmae.ToString();
            label11.Text = kaku.ToString();
            label12.Text = (machi + zumi + hmae + kmae + kaku).ToString() + "(内 エラーは" + ecount + "名)";


            dataGridView1.DataSource = Disp;


            DataTable ctdt = new DataTable();
            ctdt = Com.GetDB("select count(*) from dbo.web明細お知らせ where 処理年 = '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "' and 処理月 = '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");

            if (ctdt.Rows[0][0].ToString() != "0")
            {
                label25.Text = "明細備考 " + ctdt.Rows[0][0].ToString() + "件登録処理済";
            }
            else
            {
                if (kaku > 0 && machi + zumi + hmae + kmae == 0)
                {

                    label25.Text = "明細備考　登録がまだ完了していません！！";
                }
                else
                {
                    label25.Text = "明細備考　退職処理完了待ち";
                }
            }
        }


        //口座未登録者一覧
        private void button3_Click(object sender, EventArgs e)
        {

            DataTable dt = new DataTable();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.口座未登録者情報 ";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView1.DataSource = dt;
        }

        //当月琉銀口座登録一覧
        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.当月登録琉銀口座一覧 order by 支店名カナ";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView1.DataSource = dt;
        }



        private void button6_Click(object sender, EventArgs e)
        {
            //全対象者データ取得
            DataTable dt = co.GetKintaiKihon(1, "");

            //担当別データ
            DataTable list = co.GetKintaiKihon(9, "");

            //エラー数と警告数を取得しリストに表示
            foreach (DataRow dr in list.Rows)
            {
                string filtStr = "担当管理 = '" + dr["担当"].ToString() + "' and 登録フラグ = '1'";
                DataRow[] drYae = dt.Select(filtStr, "");


                int errorCt = 0;  //エラー件数
                int emergCt = 0;　//警告件数

                foreach (DataRow row in drYae)
                {
                    string[] st = co.ErrorCheck(row, "");

                    if (st[0].Length > 0)
                    {
                        errorCt++;
                    }

                    if (st[4].Length > 0)
                    {
                        emergCt++;
                    }
                }

                foreach (DataRow d in list.Rows)
                {
                    if (d[1].ToString() == dr["担当"].ToString())
                    {
                        d["エラー"] = errorCt.ToString();
                        d["警告"] = emergCt.ToString();
                    }
                }
            }

            dataGridView1.DataSource = list;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //全対象者データ取得
            DataTable dt = co.GetKintaiKihon(1, "");

            DataRow[] drYae = dt.Select("登録フラグ = '1'", "現場CD");

            DataTable Disp = new DataTable();
            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("漢字氏名", typeof(string));
            Disp.Columns.Add("組織名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));
            Disp.Columns.Add("担当管理", typeof(string));
            Disp.Columns.Add("状況", typeof(string));

            foreach (DataRow row in drYae)
            {
                string[] st = co.ErrorCheck(row, "");
                if (st[4].Length > 0)
                {
                    DataRow nr = Disp.NewRow();
                    nr["社員番号"] = row["社員番号"];
                    nr["漢字氏名"] = row["漢字氏名"];
                    nr["組織名"] = row["組織名"];
                    nr["現場名"] = row["現場名"];
                    nr["担当管理"] = row["担当管理"];
                    nr["状況"] = st[4].ToString();
                    Disp.Rows.Add(nr);
                }
            }

            dataGridView1.DataSource = Disp;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetData("select 社員番号 from dbo.勤怠データ group by 社員番号 having count(社員番号) > 1");
        }

        //private void button11_Click(object sender, EventArgs e)
        //{
        //    dataGridView1.DataSource = GetData("select * from dbo.再入社一覧_年調用('" + td.StartYMD.Year.ToString() + "/01/01')");
        //}

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetData("select * from 勤怠OIDSZeeM差分チェック");
        }

        private SqlConnection Cn_tate;
        private SqlDataAdapter da_tate;
        private SqlCommandBuilder cb_tate;
        private DataTable dt_tate = new DataTable();


        private void button15_Click(object sender, EventArgs e)
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";
            dataGridView1.Columns.Clear();

            //テーブルクリア
            dt_tate.Clear();
            dt_tate.Columns.Clear();

            Cn_tate = new SqlConnection(Com.SQLConstr);
            Cn_tate.Open();

            string sql = "select * from dbo.給与立替者 order by etc";
            da_tate = new SqlDataAdapter(sql, Cn_tate);
            cb_tate = new SqlCommandBuilder(da_tate);
            da_tate.Fill(dt_tate);
            dataGridView1.DataSource = dt_tate;

            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da_tate.Update(dt_tate);

                //データ更新終了をDataTableに伝える
                dt_tate.AcceptChanges();
                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }
        }

 


        private DataTable GetData(string sql)
        {
            DataTable dt = new DataTable();
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = sql;
                        Cmd.CommandTimeout = 12000;
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            return dt;
        }


        private void button23_Click(object sender, EventArgs e)
        {
            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //テーブルクリア
            dt_tate.Clear();
            dt_tate.Columns.Clear();

            Cn_tate = new SqlConnection(Com.SQLConstr);
            Cn_tate.Open();

            string sql = "select * from dbo.勤怠状況管理";
            da_tate = new SqlDataAdapter(sql, Cn_tate);
            cb_tate = new SqlCommandBuilder(da_tate);
            da_tate.Fill(dt_tate);
            dataGridView1.DataSource = dt_tate;

            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetData("select * from dbo.m明細備考欄未設定チェック");
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetData("select * from dbo.飛越退職");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Program.loginname != "喜屋武　大祐")
            {
                if (Program.loginname != "太田　朋宏")
                {
                    MessageBox.Show("だめー。");
                    return;
                }
            }

            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //TODO 毎処理でデータソースに入れる必要はない
            dataGridView1.DataSource = GetData("exec k給与縦横テーブルコピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec [dbo].[904_給与明細情報コピー] '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec k給与明細コピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec 有給管理コピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec  控除臨時コピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("SET ANSI_DEFAULTS OFF exec  web明細お知らせ設定 '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "', '" + td.StartYMD.ToString("yyyy") + "', '" + td.StartYMD.ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec  k給与預り金コピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec  k勤怠コピー '" + td.StartYMD.ToString("yyyy") + "', '" + td.StartYMD.ToString("MM") + "'");
            dataGridView1.DataSource = GetData("exec  g月給者で退職記念品対象者コピー");
            dataGridView1.DataSource = GetData("exec c管理計数給与データコピー '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "', '" + td.EndYMD.ToString("yyyy/MM/dd") + "'");
            dataGridView1.DataSource = GetData("exec c管理計数給与データ_賞与更新 '" + td.StartYMD.AddMonths(1).ToString("yyyyMM") + "'");
            Com.InHistory("計算後一括処理", "", "");


            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;

            MessageBox.Show("おわりー");

        }

        private void 出力ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string quote = "";
            string separator = ",";
            string replace = "";

            DataTable dt = new DataTable();

            dt.Columns.Add("会社コード");
            dt.Columns.Add("処理年");
            dt.Columns.Add("処理月");
            dt.Columns.Add("処理回数");
            dt.Columns.Add("社員番号");
            dt.Columns.Add("勤怠枚数通番");
            dt.Columns.Add("F0500");
            dt.Columns.Add("F0300");
            dt.Columns.Add("F0400");
            dt.Columns.Add("F0100");
            dt.Columns.Add("F2600");
            dt.Columns.Add("F0200");
            dt.Columns.Add("F0600");
            dt.Columns.Add("F0700");
            dt.Columns.Add("F0800");
            dt.Columns.Add("F0900");
            dt.Columns.Add("F1000");
            dt.Columns.Add("F1100");
            dt.Columns.Add("F1200");
            dt.Columns.Add("F1300");
            dt.Columns.Add("F1400");
            dt.Columns.Add("F1800");
            dt.Columns.Add("F1500");
            dt.Columns.Add("F1600");
            dt.Columns.Add("F1700");
            dt.Columns.Add("F1900");
            dt.Columns.Add("F2000");
            dt.Columns.Add("期間FROM");
            dt.Columns.Add("期間TO");
            dt.Columns.Add("備考");
            dt.Columns.Add("基準額１");
            dt.Columns.Add("基準額２");
            dt.Columns.Add("基準額３");

            TargetDays td = new TargetDays();

            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt2 = new DataTable();
            SqlDataAdapter da;
            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[k勤怠データ取得]";

                    Cmd.Parameters.Add(new SqlParameter("担当地区", SqlDbType.VarChar));
                    Cmd.Parameters["担当地区"].Direction = ParameterDirection.Input;
                    Cmd.Parameters["担当地区"].Value = "";

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt2);
                }
            }

            foreach (DataRow dr in dt2.Rows)
            {
                DataRow wr;
                wr = dt.NewRow();
                wr["会社コード"] = "E0";
                wr["処理年"] = td.StartYMD.AddMonths(1).Year.ToString();
                wr["処理月"] = td.StartYMD.AddMonths(1).Month.ToString();　//TODO 
                wr["処理回数"] = "0";
                wr["社員番号"] = dr["社員番号"];
                wr["勤怠枚数通番"] = "1";
                wr["F0500"] = dr["延長h"].ToString().Replace(".0", "");
                wr["F0300"] = dr["法休h"].ToString().Replace(".0", "");
                wr["F0400"] = dr["所休h"].ToString().Replace(".0", "");
                wr["F0100"] = dr["総残h"].ToString().Replace(".0", "");
                wr["F2600"] = dr["六十超h"].ToString().Replace(".0", "");
                wr["F0200"] = dr["深夜h"].ToString().Replace(".0", "");
                wr["F0600"] = dr["遅刻回"];
                wr["F0700"] = dr["遅刻h"].ToString().Replace(".0", "");
                wr["F0800"] = dr["所定"];
                wr["F0900"] = dr["法休"];
                wr["F1000"] = dr["所休"];
                wr["F1100"] = dr["有給"];
                wr["F1200"] = dr["特休"];
                wr["F1300"] = dr["無特"];
                wr["F1400"] = dr["振休"];
                wr["F1800"] = dr["公休"];
                //wr["F1500"] = dr["調休"];
                wr["F1500"] = "";
                wr["F1600"] = Convert.ToDecimal(dr["届欠"]) + Convert.ToDecimal(dr["途欠"]);
                wr["F1700"] = dr["無届"];
                wr["F1900"] = dr["回数1"];
                wr["F2000"] = dr["回数2"];
                wr["期間FROM"] = td.StartYMD.ToString("yyyy/MM/dd");
                wr["期間TO"] = td.EndYMD.ToString("yyyy/MM/dd");
                wr["備考"] = "";
                wr["基準額１"] = "0";
                wr["基準額２"] = "0";
                wr["基準額３"] = "0";
                dt.Rows.Add(wr);
            }

            SaveToCSV(dt, true, separator, quote, replace, "");
        }

        #region CSV出力処理
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt">ZeemCSV用データ</param>
        /// <param name="hasHeader">ヘッダーフラグ</param>
        /// <param name="separator"></param>
        /// <param name="quote"></param>
        /// <param name="replace"></param>
        private void SaveToCSV(DataTable dt, bool hasHeader, string separator, string quote, string replace, string tiku)
        {
            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;
            string text;

            //TODO:場所とファイル名要検討
            //string hozonFileName = @"\\Daikensrv01\電算室\勤怠チェック\ToZeem\" + tiku + ".csv";
            string hozonFileName = @"\\daikensrv03\17_総務部\04_給与\毎月給与計算業務\勤怠データ\" + tiku + td.StartYMD.ToString("yyyMM") + ".csv";

            //保存用のファイルを開く。上書きモードで。
            StreamWriter writer = new StreamWriter(hozonFileName, false, Encoding.GetEncoding("shift_jis"));
            //カラム名を保存するか
            if (hasHeader)
            {
                //カラム名を保存する場合
                for (int i = 0; i < cols; i++)
                {
                    //カラム名を取得
                    if (quote != "")
                    {
                        text = dt.Columns[i].ColumnName.Replace(quote, replace);
                    }
                    else
                    {
                        text = dt.Columns[i].ColumnName;
                    }

                    if (i != cols - 1)
                    {
                        writer.Write(quote + text + quote + separator);
                    }
                    else
                    {
                        writer.WriteLine(quote + text + quote);
                    }
                }
            }

            //データの保存処理
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (quote != "")
                    {
                        text = dt.Rows[i][j].ToString().Replace(quote, replace);
                    }
                    else
                    {
                        text = dt.Rows[i][j].ToString();
                    }

                    if (j != cols - 1)
                    {
                        writer.Write(quote + text + quote + separator);
                    }
                    else
                    {
                        writer.WriteLine(quote + text + quote);
                    }
                }
            }
            //ストリームを閉じる
            writer.Close();

            MessageBox.Show(rows.ToString() + "件出力しました。" + nl + hozonFileName);

        }
        #endregion

        private void 消去ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.loginname != "喜屋武　大祐")
            {
                MessageBox.Show("意図しないタイミングでデータが全部きえました。管理者に連絡願います。");
                return;
            }

                DialogResult result = MessageBox.Show("入力されてる勤怠データを消去していいですか？",
                        "警告",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) return;

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "truncate table dbo.[勤怠データ]";
                    using (dr = Cmd.ExecuteReader())
                    {
                        //TODO
                    }
                }
            }
        }

        private void 担当設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OwnerSet f = new OwnerSet();
            f.ShowDialog();
        }

        private void 期間設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DaysSet f = new DaysSet();
            f.ShowDialog();
        }


        //メッセージデータ取得
        public void GetMsg()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        //Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "select 処理年 + '年' + 処理月 + '月' as 処理年月, COUNT(*) as 処理数, max(お知らせ) as メッセージ from dbo.web明細お知らせ group by 処理年, 処理月 order by 処理年 desc, 処理月 desc";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        //TODO 期間
                        Cmd.CommandText = "select * from dbo.現金支給一覧表 order by 地区名 desc, 組織名 desc, 現場名 desc";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView1.DataSource = dt;

            //return;


                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("対象者無です。");
                    return;
                }

                //ボタン無効化・カーソル変更
                button15.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                string fileName = "";
                fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\05_給与受領書.xlsx";


                //エクセルオブジェクト
                Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();

                //エクセルを非表示
                m_MyExcel.Visible = false;
                //アラート非表示
                m_MyExcel.DisplayAlerts = false;

                //ブックオブジェクト
                Microsoft.Office.Interop.Excel.Workbook m_MyBook;

                //シートオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet m_MySheet;

                //ブックを開く
                m_MyBook = m_MyExcel.Workbooks.Open(Filename: fileName);

                //シート取得
                m_MySheet = m_MyBook.Worksheets[1];
                m_MySheet.Select();

                //m_MySheet.Cells[3, 1] = "【" + "test" + "】" + "test";
                //m_MySheet.Cells[3, 7] = DateTime.Today.ToString("作成日: yyyy/M/d");

                DataRow[] nenkyuudr = dt.Select("組織名 <> '-' and 現場名 <> '-'", "");

                int rows = nenkyuudr.Length;
                //int cols = dt.Columns.Count;



                for (int i = 0; i < rows; i++)
                {
                    if (i % 2 == 0)
                    {
                        m_MySheet.Cells[3, 2] = "私は、" + td.StartYMD.AddMonths(1).ToString("yyyy") + "年" + td.StartYMD.AddMonths(1).ToString("MM") + "月度給与を";
                        m_MySheet.Cells[9, 3] = nenkyuudr[i][1].ToString(); //組織
                        m_MySheet.Cells[9, 5] = nenkyuudr[i][2].ToString(); //現場
                        m_MySheet.Cells[10, 3] = nenkyuudr[i][3].ToString();　//社員番号
                        m_MySheet.Cells[10, 5] = nenkyuudr[i][4].ToString();　//氏名
                    }
                    else
                    {
                        m_MySheet.Cells[15, 2] = "私は、" + td.StartYMD.AddMonths(1).ToString("yyyy") + "年" + td.StartYMD.AddMonths(1).ToString("MM") + "月度給与を";
                        m_MySheet.Cells[21, 3] = nenkyuudr[i][1].ToString();
                        m_MySheet.Cells[21, 5] = nenkyuudr[i][2].ToString();
                        m_MySheet.Cells[22, 3] = nenkyuudr[i][3].ToString();
                        m_MySheet.Cells[22, 5] = nenkyuudr[i][4].ToString();

                        m_MySheet = m_MyBook.Worksheets[i / 2 + 2];
                    }
                }

                string localPass = @"C:\ODIS\CHECKLIST\";
                string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_") + "現金受領書";

                //フォルダがなければ作成する
                if (!System.IO.File.Exists(localPass))
                {
                    System.IO.Directory.CreateDirectory(localPass);
                }

                //excel保存 ローカルへ
                m_MyBook.SaveAs(exlName + ".xlsx");

                m_MyBook.Close(false);
                m_MyExcel.Quit();

                //excel出力
                System.Diagnostics.Process.Start(exlName + ".xlsx");

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button15.Enabled = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = GetData("select * from k給与計算後エラー一覧取得2('" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "') order by 立替 ");

            dataGridView1.Columns[12].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[13].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[14].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[15].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[16].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[17].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[18].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[19].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[20].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[21].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[22].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[23].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[24].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[25].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[26].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[27].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[28].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[29].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns[30].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns[12].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[21].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[24].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[27].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[28].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[29].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns[30].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        //明細備考確認
        private void button9_Click(object sender, EventArgs e)
        {
            GetMsg();
        }

        //private void button17_Click(object sender, EventArgs e)
        //{
        //    dataGridView1.DataSource = Com.GetDB("select * from dbo.k休業手当ODISZeeM比較('" + td.StartYMD.AddMonths(1).ToString("yyyy") + "','" + td.StartYMD.AddMonths(1).ToString("MM") + "')");
        //}

        private void button18_Click(object sender, EventArgs e)
        {
            ////マウスカーソルを砂時計にする
            //Cursor.Current = Cursors.WaitCursor;
            //button18.Enabled = false;

            //dataGridView1.DataSource = GetData("exec k休業手当削除登録 '" + td.StartYMD.AddMonths(1).ToString("yyyy") + "', '" + td.StartYMD.AddMonths(1).ToString("MM") + "'");

            ////マウスカーソルをデフォルトにする
            //Cursor.Current = Cursors.Default;
            //Application.DoEvents();
            //button18.Enabled = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Com.GetDB("select * from dbo.固定控除と変動控除と臨時のOZ誤差チェック('" + td.StartYMD.AddMonths(1).ToString("yyyy") + "','" + td.StartYMD.AddMonths(1).ToString("MM") + "')");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt = Com.GetDB("select * from dbo.t当月登録おきぎん口座一覧 order by 支店名カナ");

            dataGridView1.DataSource = dt;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            string x = td.EndYMD.ToString("yyyy/MM/dd");
            dataGridView1.DataSource = Com.GetDB("select * from [dbo].[t通勤手当①比較]('" + x + "')");
        }

        private void button21_Click(object sender, EventArgs e)
        {
            string x = td.EndYMD.ToString("yyyy/MM/dd");
            string y = td.StartYMD.AddDays(-1).ToString("yyyy/MM/dd");
            string z = td.StartYMD.ToString("yyyy/MM") + "%";
            dataGridView1.DataSource = Com.GetDB("select * from dbo.i異動入力差分_ODIS軸('" + x + "', '" + y + "', '" + z + "')");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            string x = td.EndYMD.ToString("yyyy/MM/dd");
            string y = td.StartYMD.AddDays(-1).ToString("yyyy/MM/dd");
            string z = td.StartYMD.ToString("yyyy/MM") + "%";
            dataGridView1.DataSource = Com.GetDB("select * from dbo.i異動入力差分_ZeeM軸('" + x + "', '" + y + "', '" + z + "')");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Com.GetDB("select * from [dbo].[t通勤手当②比較]");

        }

        private void button26_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Com.GetDB("select * from [dbo].[t通勤手当③比較]");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button27.Enabled = false;

            //出力処理
            DataTable EntryDisp = new DataTable();
            EntryDisp = Com.GetDB("select * from [dbo].[t通勤手当②比較_出力]");

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\70_固定金額外部取込.xlsx");

            //リストシート
            XLSheet ls = c1XLBook1.Sheets["70_固定金額外部取込"];

            int rows = EntryDisp.Rows.Count;
            int cols = EntryDisp.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[i + 1, j + 0].Value = EntryDisp.Rows[i][j].ToString();
                }
            }


            string localPass = @"\\daikensrv03\17_総務部\04_給与\毎月給与計算業務\固定金額変更取込_通勤手当他\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            c1XLBook1.Save(exlName + ".csv");
            System.Diagnostics.Process.Start(exlName + ".csv");



            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button27.Enabled = true;
        }

        //private void button28_Click(object sender, EventArgs e)
        //{

        //    //マウスカーソルを砂時計にする
        //    Cursor.Current = Cursors.WaitCursor;
        //    button28.Enabled = false;

        //    //新しいワークブックを作成します。
        //    C1XLBook c1XLBook1 = new C1XLBook();
        //    c1XLBook1.KeepFormulas = true;

        //    //ブックをロードします
        //    c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\90_給与明細.xlsx");


        //    //リストシート
        //    XLSheet ls = c1XLBook1.Sheets["List"];

        //    DataTable EntryDisp = new DataTable();

        //    string y = "2021";
        //    string m = "09";
        //    string ym = "2021/08";

        //    EntryDisp = Com.GetDB("select * from dbo.k給与明細出力('" + y + "','" + m + "')");

        //    int rows = EntryDisp.Rows.Count;
        //    int cols = EntryDisp.Columns.Count;
        //    double d;

        //    //XLStyle st1 = new XLStyle(c1XLBook1);


        //    for (int i = 0; i < rows; i++)
        //    {
        //        for (int j = 0; j < cols; j++)
        //        { 
        //            if(j == 0)
        //            {
        //                //社員番号
        //                ls[i + 2, j + 1].Value = EntryDisp.Rows[i][j];
        //            }
        //            //else if (j == 69)
        //            //{
        //            //    //時給
        //            //    ls[i + 2, j + 1].Value = EntryDisp.Rows[i][j].ToString().TrimEnd('0').TrimEnd('.');
        //            //}
        //            else if (double.TryParse(EntryDisp.Rows[i][j].ToString(), out d))
        //            {
        //                //変換出来たら、dにその数値が入る
        //                //ls[i + 2, j + 1].Style = st1;
        //                //ls[i + 2, j + 1].Value = String.Format("{0:#,0}", d);
        //                ls[i + 2, j + 1].Value = d.ToString("N2").TrimEnd('0').TrimEnd('.');

        //            }
        //            else
        //            {
        //                ls[i + 2, j + 1].Value = EntryDisp.Rows[i][j];
        //            }
        //        }
        //    }

        //    XLSheet ws = c1XLBook1.Sheets["様式"];
        //    ws[1, 1].Value = "2021年09月支給";

        //    for (int i = 1; i <= rows; i++)
        //    {
        //        XLSheet newSheet = ws.Clone();
        //        newSheet.Name = i.ToString();   // クローンをリネーム
        //        newSheet[0, 8].Value = i;      // 値の変更
        //        c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
        //        //ct2 = ct2 + 10;
        //    }

        //    c1XLBook1.Sheets.Remove("様式");

        //    string localPass = @"C:\ODIS\MEISAI\";
        //    string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

        //    //フォルダがなければ作成する
        //    if (!System.IO.File.Exists(localPass))
        //    {
        //        System.IO.Directory.CreateDirectory(localPass);
        //    }

        //    c1XLBook1.Save(exlName + ".xlsx");

        //    //Excel Change PDF           
        //    Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
        //    m_MyExcel.Visible = false; //エクセルを非表示
        //    m_MyExcel.DisplayAlerts = false; //アラート非表示
        //    Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
        //    //Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト

        //    //ブックを開く
        //    m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");

        //    //PDF保存
        //    m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");

        //    m_MyBook.Close(false);
        //    m_MyExcel.Quit();

        //    //FileInfo file = new FileInfo(exlName + ".xlsx");
        //    //file.Delete();

        //    //PDF出力
        //    //System.Diagnostics.Process.Start(exlName + ".xlsx");
        //    System.Diagnostics.Process.Start(@"C:\ODIS\MEISAI\");
        //    System.Diagnostics.Process.Start(exlName + ".pdf");

        //    //マウスカーソルをデフォルトにする
        //    Cursor.Current = Cursors.Default;
        //    Application.DoEvents();
        //    button28.Enabled = true;
        //}

        //private void button11_Click(object sender, EventArgs e)
        //{

        //    //マウスカーソルを砂時計にする
        //    Cursor.Current = Cursors.WaitCursor;
        //    button28.Enabled = false;

        //    //新しいワークブックを作成します。
        //    C1XLBook c1XLBook1 = new C1XLBook();
        //    c1XLBook1.KeepFormulas = true;

        //    //ブックをロードします
        //    c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\91_登録番号_案内とお願い.xlsx");


        //    //リストシート
        //    XLSheet ls = c1XLBook1.Sheets["List"];

        //    //DataTable EntryDisp = new DataTable();

        //    //EntryDisp = Com.GetDB("select top 10 * from dbo.t取引先");

        //    //int rows = EntryDisp.Rows.Count;
        //    //int cols = EntryDisp.Columns.Count;
        //    //double d;

        //    //for (int i = 0; i < rows; i++)
        //    //{
        //    //    for (int j = 0; j < cols; j++)
        //    //    {
        //    //        ls[i+1, j+1].Value = EntryDisp.Rows[i][j];
        //    //    }
        //    //}

        //    XLSheet ws = c1XLBook1.Sheets["様式"];
        //    //ws[1, 1].Value = "2021年09月支給";

        //    for (int i = 1; i <= 213; i++)
        //    {
        //        XLSheet newSheet = ws.Clone();
        //        newSheet.Name = i.ToString();   // クローンをリネーム
        //        newSheet[0,6].Value = i;      // 値の変更
        //        //newSheet.PageSetup.PrintArea = @"$A$1:$K$43";
        //        c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
        //        //ct2 = ct2 + 10;
        //    }
            

        //    c1XLBook1.Sheets.Remove("様式");

        //    string localPass = @"C:\ODIS\ANNAI\";
        //    string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

        //    //フォルダがなければ作成する
        //    if (!System.IO.File.Exists(localPass))
        //    {
        //        System.IO.Directory.CreateDirectory(localPass);
        //    }

        //    c1XLBook1.Save(exlName + ".xlsx");


        //    //TODO コメントアウト202309

        //    ////Excel Change PDF           
        //    //Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
        //    //m_MyExcel.Visible = false; //エクセルを非表示
        //    //m_MyExcel.DisplayAlerts = false; //アラート非表示
        //    //Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
        //    ////Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト

        //    ////ブックを開く
        //    //m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");

        //    ////PDF保存
        //    //m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");

        //    //m_MyBook.Close(false);
        //    //m_MyExcel.Quit();

        //    ////FileInfo file = new FileInfo(exlName + ".xlsx");
        //    ////file.Delete();

        //    ////PDF出力
        //    ////System.Diagnostics.Process.Start(exlName + ".xlsx");
        //    //System.Diagnostics.Process.Start(@"C:\ODIS\ANNAI\");
        //    //System.Diagnostics.Process.Start(exlName + ".pdf");

        //    //TODO ここまで

        //    //マウスカーソルをデフォルトにする
        //    Cursor.Current = Cursors.Default;
        //    Application.DoEvents();
        //    button28.Enabled = true;
        //}
    }
}
