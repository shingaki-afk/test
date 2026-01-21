using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Timers;
using System.Net.Mail;
using System.Net;

namespace ODIS.ODIS
{
    public partial class Kintone : Form
    {
        public string batchStdOut;
        public string batchStdErr;
        public int batchExitCode = 0;

        public string info;
        public string nl = Environment.NewLine;

        //タイマー
        System.Timers.Timer timer;

        //インサートダミーテーブル
        DataTable dummy;
        public Kintone()
        {
            InitializeComponent();

            GetData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.exkin差分チェック ");
            dataGridView1.DataSource = dt;
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

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            DateTime dt = DateTime.Now;
            string ymdhms = dt.ToString("yyyyMMdd_HHmmss");
            string junban = "社員番号,氏名,カナ名,地区CD,地区名,組織CD,組織名,現場CD,現場名,役職CD,役職名,給与支給区分,給与支給区分名,生年月日,入社年月日,退職年月日,性別区分,契約社員,休暇付与区分,職種,最終学歴,社外経験,メール,内線番号,携帯区分,携帯番号,短縮番号,郵便番号,住所,雇保,健保,障区分,障内容,障備考";

            //csvをキントーンに同期
            string batPathup = "cli-kintone.exe --export -a 71 -d oki-daiken -t NHr8IUXco2lT1HltFh0QIJjtIioivjQ7OOTV4w5X -e sjis -c " + junban + @" > \\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\export_" + ymdhms + ".csv";

            //コマンドプロンプトで実行
            Com.ExecBatProcess(batPathup, out batchStdOut, out batchStdErr, out batchExitCode);

            DataTable dtable = new DataTable();
            string path = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\export_" + ymdhms + ".csv";
            //string path = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\export_20220421_111331.csv";

            dtable = extractAllCsv(path);

            dataGridView1.DataSource = dtable;

            //削除
            DataTable dtdel = new DataTable();
            dtdel = Com.GetDB("delete from dbo.exkin社員基本情報");

            using (var bulkCopy = new SqlBulkCopy(ODIS.Com.SQLConstr))
            {
                bulkCopy.DestinationTableName = "exkin社員基本情報"; //dt.TableName; // テーブル名をSqlBulkCopyに教える
                bulkCopy.WriteToServer(dtable); // bulkCopy実行
            }

            MessageBox.Show("exkin社員基本情報へ同期、おわりー");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (numericUpDown1.Value <= 0)
            {
                MessageBox.Show("1分以上にして。。");
                return;
            }

            label5.Text = "間隔" + Convert.ToInt32(numericUpDown1.Value).ToString() + "分で同期開始中。";

            int hh = Convert.ToInt32(numericUpDown1.Value * 60000);


            //イベント間隔1000ミリ秒でタイマーを初期化
            timer = new System.Timers.Timer(hh); //60000で1分

            //タイマーにイベントを登録
            timer.Elapsed += OnTimedEvent;

            //タイマーを開始する
            timer.Start();

            // バルーンヒントを表示する
            //notifyIcon.BalloonTipTitle = "KintoneSync";
            //notifyIcon.BalloonTipText = "タイマー同期が開始されました。";
            //notifyIcon.ShowBalloonTip(30); //ミリ秒

            btnstart.Enabled = false;
            btnstop.Enabled = true;

            GetData();

        }

        private void GetData()
        {
            dataGridView1.DataSource = null;
            DataTable getdb = new DataTable();
            getdb = Com.GetDB("select top 100 * from dbo.KintoneSync order by 処理日時 desc");
            dataGridView1.DataSource = getdb;
        }


        //タイマーに呼び出されるメソッドの定義
        private void OnTimedEvent(Object sender, ElapsedEventArgs e)
        {
            Sync();
        }

        private void SendMail(string sub, string body)
        {
            //メールとばす
            using (var client = new SmtpClient("smtp.gmail.com", 587))
            {
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential("admin@oki-daiken.co.jp", "admin0110");
                client.EnableSsl = true;

                // MailMessageクラスを使って送信するメールを作成する
                var message = new MailMessage();

                // 差出人アドレス
                message.From = new MailAddress("webmaster@oki-daiken.co.jp", "KintoneSyncエラー");
                message.To.Add(new MailAddress("kyan@oki-daiken.co.jp"));
                message.To.Add(new MailAddress("ota-tomo@oki-daiken.co.jp"));

                // メールの件名
                //message.Subject = "内線or携帯の重複エラー";
                message.Subject = sub;

                // メールの本文
                //message.Body = no;
                message.Body = body;

                try
                {
                    // 作成したメールを送信する
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("例外が発生しました" + ex);
                    //メール送信
                }
            }
        }

        private void Sync()
        {
            //最初に、内線・携帯に重複がないかチェック
            DataTable dtck = new DataTable();
            dtck = Com.GetDB("select * from dbo.j重複チェック_内線携帯 ");

            if (dtck.Rows.Count > 0)
            {
                string no = "";
                //重複あれば終了
                for (int i = 0; i < dtck.Rows.Count; i++)
                {
                    no += dtck.Rows[i][0].ToString() + " ";
                }

                dummy = Com.GetDB("insert dbo.KintoneSync (処理日時, 内容) VALUES('" + DateTime.Now + ':' + DateTime.Now.Millisecond + "','" + no + "')");

                //メール送信
                SendMail("内線or携帯の重複エラー", no);

                return;
            }

            info = "";
            batchStdOut = "";
            batchStdErr = "";
            batchExitCode = 0;

            string quote = "\"";
            string separator = ",";
            string replace = "";
            DateTime dt = DateTime.Now;
            string ymdhms = dt.ToString("yyyyMMdd_HHmmss");

            //①INSERT処理

            //InsertCSV作成
            string passin = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\insert_" + ymdhms + ".csv";
            DataTable dtin = new DataTable();

            //差分データ出力
            dtin = Com.GetDB("select * from dbo.ex社員基本情報インサートCSV");

            if (dtin.Rows.Count > 0)
            {
                info += "Insert " + dtin.Rows.Count + "件    ";

                SqlConnection Cn;
                SqlCommand Cmd;
                SqlDataAdapter da;

                Cn = new SqlConnection(Com.SQLConstr);
                Cn.Open();

                //トランザクション開始
                SqlTransaction tran = Cn.BeginTransaction();


                try
                {
                    //データがあればCSV出力
                    Com.OutPutCSV(dtin, true, separator, quote, replace, passin);

                    //csvをキントーンに同期
                    string batPathin = @"cli-kintone.exe --import -a 71 -d oki-daiken -t NHr8IUXco2lT1HltFh0QIJjtIioivjQ7OOTV4w5X -e sjis -f \\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\insert_" + ymdhms + ".csv";
                    Com.ExecBatProcess(batPathin, out batchStdOut, out batchStdErr, out batchExitCode);

                    //if (batchStdOut.ToString() != "")
                    if (batchStdErr.ToString() != "")
                    {
                        //メール送信
                        SendMail("Insertエラー", batchStdOut.ToString());
                        info += "Insertエラー " + batchStdOut.ToString();
                        return;
                    }
                    //ex社員基本情報へインサート
                    DataTable dtin2 = new DataTable();

                    string sql = "insert ex社員基本情報 select * from dbo.ex社員基本情報インサートCSV";
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = sql;
                        Cmd.CommandTimeout = 12000;
                        Cmd.Transaction = tran;
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dtin2);

                        tran.Commit();
                    }

                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Insert Error" + ex.ToString());
                    tran.Rollback();

                    //メール送信
                    SendMail("Insertエラー", "Insertに失敗したようです。" + ex.ToString());

                    throw;
                }
            }
            else
            {
                info += "Insert 0件   ";
            }

            //TODO
            //if (batchStdOut != "") info += "batchStdOut: " + batchStdOut + nl;
            //if (batchStdErr != "") info += "batchStdErr: " + batchStdErr + nl;
            //if (batchExitCode.ToString() != "0") info += "batchExitCode: " + batchExitCode.ToString() + nl;

            batchStdOut = "";
            batchStdErr = "";
            batchExitCode = 0;

            //②UPDATE処理
            string passup = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\update_" + ymdhms + ".csv";
            DataTable dtup = new DataTable();

            dtup = Com.GetDB("select * from dbo.ex社員基本情報アップデートCSV");

            if (dtup.Rows.Count > 0)
            {
                info += "Update  " + dtup.Rows.Count + "件" + nl;

                SqlConnection Cn;
                SqlCommand Cmd;
                SqlDataAdapter da;

                //using (Cn = new SqlConnection(ODIS.Com.SQLConstr))
                //{ 
                //トランザクション開始

                Cn = new SqlConnection(Com.SQLConstr);
                Cn.Open();

                SqlTransaction tran = Cn.BeginTransaction();

                try
                {
                    //データがあればCSV出力
                    Com.OutPutCSV(dtup, true, separator, quote, replace, passup);

                    //csvをキントーンに同期
                    string batPathup = @"cli-kintone.exe --import -a 71 -d oki-daiken -t NHr8IUXco2lT1HltFh0QIJjtIioivjQ7OOTV4w5X -e sjis -f \\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\update_" + ymdhms + ".csv";
                    //string batPathup = @"cli-kintone.exe --import -a 71 -d oki-daiken -t NHr8IUXco2lT1HltFh0QIJjtIioivjQ7OOTV4w5X -e sjis -f \\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\csv\insert_20230704_093000.csv";
                    Com.ExecBatProcess(batPathup, out batchStdOut, out batchStdErr, out batchExitCode);

                    //if (batchStdOut.ToString() != "")
                    if (batchStdErr.ToString() != "")
                    {
                        //メール送信
                        SendMail("Upadteエラー", batchStdOut.ToString());
                        info += "Upadteエラー " + batchStdOut.ToString();
                        return;
                    }

                    //ex社員基本情報を更新
                    DataTable dtin2 = new DataTable();

                    string sql = "exec dbo.ex社員基本情報アップデート更新";
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = sql;
                        Cmd.CommandTimeout = 12000;
                        Cmd.Transaction = tran;
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dtin2);

                        tran.Commit();
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("UPDATE Error" + ex.ToString());
                    tran.Rollback();
                    //メール送信

                    SendMail("UPDATEエラー", "UPDATEに失敗したようです。" + ex.ToString());

                    throw;
                }
            }
            else
            {
                info += "Update  0件" + nl;
            }

            //TODO
            //if (batchStdOut != "") info += "batchStdOut: " + batchStdOut + nl;
            //if (batchStdErr != "") info += "batchStdErr: " + batchStdErr + nl;
            //if (batchExitCode.ToString() != "0") info += "batchExitCode: " + batchExitCode.ToString() + nl;

            dummy = Com.GetDB(@"insert dbo.KintoneSync (処理日時, 内容) VALUES('" + DateTime.Now + "','" + info + "')");
        }

        private void btnstop_Click(object sender, EventArgs e)
        {
            timer.Stop();

            btnstart.Enabled = true;
            btnstop.Enabled = false;

            GetData();

            label5.Text = "同期停止中。";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Sync();
            GetData();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
