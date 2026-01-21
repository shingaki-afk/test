using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using ODIS.ODIS;

namespace ODIS
{
    public partial class Administrator : Form
    {
        ///メインデータ

        /// <summary>
        /// 改行コード
        /// </summary>
        public string nl = Environment.NewLine;

        /// <summary>
        /// 社員情報+勤怠データ
        /// </summary>
        private DataTable dt;

        /// <summary>
        /// 対象期間インスタンス
        /// </summary>
        private TargetDays td = new TargetDays();

        /// <summary>
        /// 共通クラスのインスタンス
        /// </summary>
        private Common co = new Common();

        //TODO:パスをそとだしに
        /// <summary>
        /// OCR出力ファイルの出力パス
        /// </summary>
        public string ocrFolderPass = @"\\Daikensrv01\電算室\勤怠チェック\OCR\";

        //入力制限設定
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dtlimit = new DataTable();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Administrator()
        {
            //コントロール初期設定
            InitializeComponent();
            dt = co.GetKintaiKihon(1, "");

            GetMsg();


            //入力制限設定
            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();
            string sql = "select * from dbo.地区別制限フラグ";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dtlimit);
            dataGridView2.DataSource = dtlimit;
            //GetData();
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

        #region CSV出力処理

        private void 那覇ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutCsv("那覇");
        }

        private void 八重山ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutCsv("八重山");
        }

        private void 北部ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutCsv("北部");
        }

        private void OutCsv(string tiku)
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
                    Cmd.CommandText = "[dbo].[勤怠データ取得]";

                    Cmd.Parameters.Add(new SqlParameter("担当地区", SqlDbType.VarChar));
                    Cmd.Parameters["担当地区"].Direction = ParameterDirection.Input;
                    Cmd.Parameters["担当地区"].Value = tiku;

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
                wr["F1600"] = dr["届欠"];
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

            SaveToCSV(dt, true, separator, quote, replace, tiku);
        }

        #endregion


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

        private void 消去ToolStripMenuItem_Click(object sender, EventArgs e)
        {
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text == "" || this.comboBox2.Text == "" || this.textBox1.Text == "")
            {
                MessageBox.Show("年月とメッセージを入れてください");
                return;
            }



            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataReader dr;

            try
            {
                 using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[給与明細メッセージ設定]";

                            Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                            Cmd.Parameters["year"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("month", SqlDbType.VarChar));
                            Cmd.Parameters["month"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("start", SqlDbType.VarChar));
                            Cmd.Parameters["start"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("end", SqlDbType.VarChar));
                            Cmd.Parameters["end"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("msg", SqlDbType.VarChar));
                            Cmd.Parameters["msg"].Direction = ParameterDirection.Input;

                            //Cmd.Parameters.Add(new SqlParameter("ct", SqlDbType.Int));
                            //Cmd.Parameters["ct"].Direction = ParameterDirection.Output;

                            TargetDays td = new TargetDays();
                            //Cmd.Parameters["year"].Value = td.StartYMD.AddMonths(1).Year.ToString();
                            //Cmd.Parameters["month"].Value = td.StartYMD.AddMonths(1).ToString("MM");
                            //Cmd.Parameters["start"].Value = td.StartYMD.ToString("yyyyMMdd");

                            Cmd.Parameters["year"].Value = this.comboBox1.SelectedItem.ToString();
                            Cmd.Parameters["month"].Value = this.comboBox2.SelectedItem.ToString();

                            DateTime ym = Convert.ToDateTime(this.comboBox1.SelectedItem.ToString() + "/" + this.comboBox2.SelectedItem.ToString() + "/01");
                            Cmd.Parameters["start"].Value = ym.AddMonths(-1).ToString("yyyy/MM/dd");
                            Cmd.Parameters["end"].Value = ym.AddDays(-1).ToString("yyyy/MM/dd");
                            Cmd.Parameters["msg"].Value = this.textBox1.Text;

                            using (dr = Cmd.ExecuteReader())
                            {
                                GetMsg();
                                //this.label1.Text = Cmd.Parameters["ct"].Value.ToString();
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
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
                        Cmd.CommandText = "select 年 + '年' + 月 + '月' as 処理年月, COUNT(*) as 処理数, MIN(メッセージ１) as メッセージ from QUATRO.dbo.QCTTMSG where 会社コード = 'E0' group by 年, 月 order by 年 desc, 月 desc";
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

        private void 全地区ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutCsv("");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dtlimit);

                //データ更新終了をDataTableに伝える
                dtlimit.AcceptChanges();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }
        }

        private void 更新ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
