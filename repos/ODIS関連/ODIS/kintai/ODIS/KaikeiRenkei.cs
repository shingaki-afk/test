using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KaikeiRenkei : Form
    {

        private TargetDays td = new TargetDays();


        public KaikeiRenkei()
        {
            InitializeComponent();

            if (Program.loginname != "喜屋武　大祐")
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("09_会計連携", "", "");
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string yyyyMM = td.StartYMD.ToString("yyyyMM");
            //yyyyMM = "202406";

            //削除
            DataTable dtdel = new DataTable();
            dtdel = Com.GetDB("delete from dbo.pプロステージ売上データ where uriageym = '" + yyyyMM + "'");
            //dtdel = Com.GetDB("delete from dbo.pプロステージ売上データ ");

            DataTable dt = new DataTable();
            dt = Com.GetPosDB("select * from kpcp01.\"CostomGetUriageRenkeiData\" where uriageym = '" + yyyyMM + "'");
            //dt = Com.GetPosDB("select * from kpcp01.\"CostomGetUriageRenkeiData\" ");

            using (var bulkCopy = new SqlBulkCopy(ODIS.Com.SQLConstr))
            {
                bulkCopy.DestinationTableName = "pプロステージ売上データ"; //dt.TableName; // テーブル名をSqlBulkCopyに教える
                bulkCopy.WriteToServer(dt); // bulkCopy実行
            }

            MessageBox.Show("プロステージ⇒Sqlserverへ売上データをコピーしました。");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sday = td.StartYMD.ToString("yyyyMMdd");
            string eday = td.EndYMD.ToString("yyyyMMdd");

            //削除
            DataTable dtdel = new DataTable();
            dtdel = Com.GetDB("delete from dbo.s新旧仕訳元データ");

            DataTable dt = new DataTable();
            dt = Com.GetPosDB("select * FROM kpcp01.\"CostomGetFurikaeDataAll\" where suitouymd between '" + sday + "' and '" + eday + "'");

            using (var bulkCopy = new SqlBulkCopy(ODIS.Com.SQLConstr))
            {
                bulkCopy.DestinationTableName = "s新旧仕訳元データ"; //dt.TableName; // テーブル名をSqlBulkCopyに教える
                bulkCopy.WriteToServer(dt); // bulkCopy実行
            }

            MessageBox.Show("プロステージ⇒Sqlserverへ入金データをコピーしました。");
        }

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
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[c管理計数更新]";

                        Cmd.Parameters.Add(new SqlParameter("ym_mae", SqlDbType.Char));
                        Cmd.Parameters["ym_mae"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("ym", SqlDbType.Char));
                        Cmd.Parameters["ym"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("syouyo", SqlDbType.VarChar));
                        Cmd.Parameters["syouyo"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("taisyoku", SqlDbType.VarChar));
                        Cmd.Parameters["taisyoku"].Direction = ParameterDirection.Input;

                        TargetDays td = new TargetDays();
                        Cmd.Parameters["ym_mae"].Value = td.StartYMD.ToString("yyyyMM");
                        Cmd.Parameters["ym"].Value = td.StartYMD.AddMonths(1).ToString("yyyyMM");

                        Cmd.Parameters["syouyo"].Value = "1.15"; //TODO 現在使用しておりません。
                        Cmd.Parameters["taisyoku"].Value = "0.05";　//TODO 現在使用しておりません。

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

            MessageBox.Show("管理計数テーブルにデリート/インサートしました。");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select 伝票番号, count(*) as 伝票行数 from dbo.PCA会計仕訳データ where 伝票日付 like '" + td.StartYMD.ToString("yyyyMM") + "%' and 摘要文 = '入金自動振替' group by 伝票番号");
            dataGridView1.DataSource = dt;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.n入金連携データ取得");
            //dataGridView1.DataSource = dt;

            //CSVファイルが既にあれば削除
            string FilePath = @"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\12_入金\入金_" + td.StartYMD.ToString("yyyyMM") + ".csv";
            if (File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }

            //様式コピー
            File.Copy(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\12_入金\入金_様式.csv", FilePath);

            //CSVファイルに書き込むときに使うEncoding
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string strOutputPath = FilePath;

            //書き込むファイルを開く
            using (StreamWriter sr = new StreamWriter(strOutputPath, true, enc))
            {
                //dt.Columns[i].ColumnName

                //レコードを書き込む
                foreach (DataRow row in dt.Rows)
                {
                    sr.WriteLine(string.Join(",", row.ItemArray));
                }
            }

            MessageBox.Show("CSV出力しました。" + Environment.NewLine + FilePath);
            System.Diagnostics.Process.Start(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\12_入金\");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DataTable dtp = new DataTable();
            DataTable dtm = new DataTable();
            dtp = Com.GetDB("select * from dbo.r連携_売上CSV出力_金額プラス('" + td.EndYMD.ToString("yyyyMMdd") + "','" + td.EndYMD.ToString("MM") + "','" + td.EndYMD.ToString("yyyyMM") + "') order by 貸方部門コード, 貸方工事コード ");
            dtm = Com.GetDB("select * from dbo.r連携_売上CSV出力_金額マイナス('" + td.EndYMD.ToString("yyyyMMdd") + "','" + td.EndYMD.ToString("MM") + "','" + td.EndYMD.ToString("yyyyMM") + "') order by 借方部門コード, 借方工事コード ");

            //dtp = Com.GetDB("select * from dbo.u売上CSV出力_金額プラス('" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("yyyyMMdd") + "','" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("MM") + "','" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("yyyyMM") + "') order by 貸方部門コード, 貸方工事コード ");
            //dtm = Com.GetDB("select * from dbo.u売上CSV出力_金額マイナス('" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("yyyyMMdd") + "','" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("MM") + "','" + td.EndYMD.AddYears(-1).AddMonths(0).ToString("yyyyMM") + "') order by 借方部門コード, 借方工事コード ");

            dtp.Merge(dtm);

            //CSVファイルが既にあれば削除
            string FilePath = @"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\11_売上\売上_" + td.StartYMD.ToString("yyyyMM") + ".csv";
            //string FilePath = @"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\11_売上\売上_" + td.StartYMD.AddYears(-1).AddMonths(0).ToString("yyyyMM") + ".csv";
            if (File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }

            //様式コピー
            File.Copy(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\11_売上\売上_様式.csv", FilePath);

            //CSVファイルに書き込むときに使うEncoding
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string strOutputPath = FilePath;

            //書き込むファイルを開く
            using (StreamWriter sr = new StreamWriter(strOutputPath, true, enc))
            {
                //dt.Columns[i].ColumnName

                //レコードを書き込む
                foreach (DataRow row in dtp.Rows)
                {
                    sr.WriteLine(string.Join(",", row.ItemArray));
                }
            }

            MessageBox.Show("CSV出力しました。" + Environment.NewLine + FilePath);
            System.Diagnostics.Process.Start(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\11_売上\");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataTable dtp = new DataTable();
            DataTable dtm = new DataTable();
            dtp = Com.GetDB("select * from dbo.zZeeM会計連携_00貸方('" + td.EndYMD.ToString("yyyy/MM/dd") + "','" + td.EndYMD.AddDays(1).ToString("yyyy") + "','" + td.EndYMD.AddDays(1).ToString("MM") + "') where 貸方金額 <> 0 ");
            dtm = Com.GetDB("select * from dbo.r連携_人給社保以外CSV出力_金額マイナス('" + td.EndYMD.AddDays(1).ToString("yyyy") + "','" + td.EndYMD.AddDays(1).ToString("MM") + "','" + td.EndYMD.ToString("yyyy/MM/dd") + "') order by 借方部門コード, 借方工事コード ");

            dtp.Merge(dtm);

            //CSVファイルが既にあれば削除
            string FilePath = @"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\13_人給\01社保以外_" + td.StartYMD.ToString("yyyyMM") + ".csv";
            if (File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }

            //様式コピー
            File.Copy(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\13_人給\01_社保以外_様式.csv", FilePath);

            //CSVファイルに書き込むときに使うEncoding
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string strOutputPath = FilePath;

            //書き込むファイルを開く
            using (StreamWriter sr = new StreamWriter(strOutputPath, true, enc))
            {
                //レコードを書き込む
                foreach (DataRow row in dtp.Rows)
                {
                    sr.WriteLine(string.Join(",", row.ItemArray));
                }
            }

            MessageBox.Show("CSV出力しました。" + Environment.NewLine + FilePath);
            System.Diagnostics.Process.Start(@"\\daikensrv03\17_総務部\07_システム\10_次期システム等資料\会計システム\11_売上\");
        }
    }
}
