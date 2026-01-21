using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Windows.Forms;
using ODIS.ODIS;

namespace ODIS
{
    static class Program
    {
        public static string machine;
        public static string account;
        public static string loginID;
        public static string access;
        public static string loginbusyo;
        public static string loginname;
        public static int dispZinzi;
        public static string loginml;
        public static string tiku;
        public static string soshiki;
        public static string yakusyokucd;

        public static DataTable acdt = new DataTable();

        [STAThread]
        static void Main()
        {
            // グローバル例外ハンドラ設定
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += GlobalThreadExceptionHandler;
            AppDomain.CurrentDomain.UnhandledException += GlobalUnhandledExceptionHandler;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            machine = Environment.MachineName;
            account = Environment.UserName;

            if (account == "goya")
                account = "14819802";

            acdt = Com.GetDB("select a.*, b.メール, b.地区名, b.組織名, b.役職CD from dbo.アカウント管理 a left join dbo.社員基本情報 b on a.ID = b.社員番号 ");
            bool flg = false;
            foreach (DataRow row in acdt.Rows)
            {
                if (row["ID"].ToString() == account)
                {
                    Program.loginID = row["ID"].ToString();
                    Program.access = row["権限"].ToString();
                    Program.loginbusyo = row["部署"].ToString();
                    Program.loginname = row["名前"].ToString();
                    Program.dispZinzi = row["人事検索権限"].Equals(DBNull.Value) ? 0 : (int)row["人事検索権限"];
                    Program.loginml = row["メール"].ToString();
                    Program.tiku = row["地区名"].ToString();
                    Program.soshiki = row["組織名"].ToString();
                    Program.yakusyokucd = row["役職CD"].ToString();
                    flg = true;
                }
            }

            var dtnow = DateTime.Now;
            var dts = new DateTime(2024, 6, 1, 22, 0, 0);
            var dte = new DateTime(2024, 6, 2, 05, 0, 0);
            if (dtnow > dts && dtnow < dte)
            {
                MessageBox.Show("メンテ中です。　6/1 20:00～ 6/2 5:00");
                Com.InHistory("使用不可" + account + "_" + machine, "", "");
                return;
            }

            if (flg)
            {
                Com.InHistory("00_ログイン", "", "");
                Application.Run(new Main());
                return;
            }

            // りんご例外対応
            if (Program.loginname == "呉屋　武" || account == "呉屋　武" || machine == "INTELMBA")
            {
                Com.InHistory("ログイン_りんご対応", "", "");
                Application.Run(new Main());
                return;
            }

            MessageBox.Show("ODISカウントがありません。");
            Com.InHistory("ODISカウント無_" + account + "_" + machine, "", "");
        }

        // ========= 例外ハンドラ =========

        static void GlobalThreadExceptionHandler(object sender, System.Threading.ThreadExceptionEventArgs e)
            => HandleException(e.Exception, "Application.ThreadException");

        static void GlobalUnhandledExceptionHandler(object sender, UnhandledExceptionEventArgs e)
            => HandleException(e.ExceptionObject as Exception, "AppDomain.UnhandledException");

        static void HandleException(Exception ex, string sourceTag)
        {
            if (ex == null) return;

            try
            {
                // 画面通知
                MessageBox.Show(
                    $"エラーが発生しました。\n{ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // エラー情報抽出
                var loc = FindExceptionLocation(ex);

                // メール送信
                SendErrorEmail(ex, loc, sourceTag);
            }
            catch (Exception emailEx)
            {
                // メール送信自体の失敗は落とさない
                Debug.WriteLine("Error mail send failed: " + emailEx);
                Debug.WriteLine("Original exception: " + ex);
            }
        }

        // ★ 例外からファイル/行/メソッドを特定
        private static (string File, int Line, string Member) FindExceptionLocation(Exception ex)
        {
            var st = new StackTrace(ex, true); // true: ソース情報
            var frames = st.GetFrames();

            string file = "(unknown)";
            int line = 0;
            string member = ex.TargetSite != null
                ? $"{ex.TargetSite.DeclaringType?.FullName}.{ex.TargetSite.Name}"
                : "(unknown)";

            if (frames != null && frames.Length > 0)
            {
                var f = Array.Find(frames, fr => fr.GetFileLineNumber() > 0) ?? frames[0];
                var path = f.GetFileName();
                file = string.IsNullOrWhiteSpace(path) ? "(unknown)" : Path.GetFileName(path);
                line = f.GetFileLineNumber();
                var mb = f.GetMethod();
                if (mb != null)
                {
                    var typeName = (mb.DeclaringType != null) ? mb.DeclaringType.FullName : "(global)";
                    member = $"{typeName}.{mb.Name}";
                }
            }
            return (file, line, member);
        }

        // ★ メール送信（場所情報を本文に含める）
        private static void SendErrorEmail(Exception ex, (string File, int Line, string Member) loc, string sourceTag)
        {
            string appVer = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            string subject = $"【アプリケーションエラー通知】{Program.loginname}";
            string body =
$@"日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}
発生元: {sourceTag}
メッセージ: {ex.Message}

--- 例外発生箇所 ---
ファイル: {loc.File}
行番号: {loc.Line}
メソッド: {loc.Member}

--- 実行環境 ---
ユーザー: {Program.account}
マシン: {Program.machine}
アプリ版: {appVer}
プロセス: {Process.GetCurrentProcess().ProcessName} (PID {Process.GetCurrentProcess().Id})
--- スタックトレース ---
{ex}";

            using (var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential("admin@oki-daiken.co.jp", "auyj glfh umla akat"),
                EnableSsl = true,
            })
            using (var mailMessage = new MailMessage
            {
                From = new MailAddress("webmaster@oki-daiken.co.jp"),
                Subject = subject,
                Body = body,
                IsBodyHtml = false,
            })
            {
                mailMessage.To.Add("kyan@oki-daiken.co.jp");
                smtpClient.Send(mailMessage);
            }
        }
    }
}
