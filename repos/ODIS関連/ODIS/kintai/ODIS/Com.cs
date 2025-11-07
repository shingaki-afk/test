using System;
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using Npgsql;
using C1.C1Excel;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ODIS.ODIS
{
    class Com
    {
        public static string SQLConstr = "Password=Pa$$w0rd;User ID=developer;Initial Catalog=dev;Data Source=192.168.100.10";
        public static string POSConstr = "Server=192.168.100.2;Port=5432;User Id=prostage;Password=prostage;Database=prostage;CommandTimeout=400";

        /// <summary>
        /// 改行コード
        /// </summary>
        public static string nl = Environment.NewLine;

        //--------------------------------------------------------------------------------
        //1バイト文字で構成された文字列であるか判定
        //
        //1バイト文字のみで構成された文字列 : 全角変換
        //2バイト文字が含まれている文字列   : 半角変換

        //半角があれば全角
        //全角があれば半角
        //--------------------------------------------------------------------------------
        public static string isOneByteChar(string str)
        {
            byte[] byte_data = System.Text.Encoding.GetEncoding(932).GetBytes(str);
            //if (byte_data.Length == str.Length)
            if (byte_data.Length < str.Length * 2)
            {
                //全角変換
                return Microsoft.VisualBasic.Strings.StrConv(str, VbStrConv.Wide, 0);
            }
            else
            {
                //半角変換
                return Microsoft.VisualBasic.Strings.StrConv(str, VbStrConv.Narrow, 0);
            }
        }



        public static void InHistory(string type, string str, string count)
        {
            if (Program.machine == "DENSAN004-PC") return;
            SqlConnection Cn;
            SqlCommand Cmd;

            using (Cn = new SqlConnection(SQLConstr))
            {
                Cn.Open();
                Cmd = Cn.CreateCommand();
                Cmd.CommandType = CommandType.StoredProcedure;
                Cmd.CommandText = "[dbo].[InsertHistory]";

                Cmd.Parameters.Add(new SqlParameter("type", SqlDbType.VarChar));
                Cmd.Parameters["type"].Direction = ParameterDirection.Input;

                Cmd.Parameters.Add(new SqlParameter("datetime", SqlDbType.DateTime));
                Cmd.Parameters["datetime"].Direction = ParameterDirection.Input;

                Cmd.Parameters.Add(new SqlParameter("ipadress", SqlDbType.VarChar));
                Cmd.Parameters["ipadress"].Direction = ParameterDirection.Input;

                Cmd.Parameters.Add(new SqlParameter("hostname", SqlDbType.VarChar));
                Cmd.Parameters["hostname"].Direction = ParameterDirection.Input;

                Cmd.Parameters.Add(new SqlParameter("searchcondit", SqlDbType.VarChar));
                Cmd.Parameters["searchcondit"].Direction = ParameterDirection.Input;

                Cmd.Parameters.Add(new SqlParameter("searchresult", SqlDbType.VarChar));
                Cmd.Parameters["searchresult"].Direction = ParameterDirection.Input;

                Cmd.Parameters["type"].Value = type;
                Cmd.Parameters["datetime"].Value = DateTime.Now;
                Cmd.Parameters["ipadress"].Value = Program.loginname;
                Cmd.Parameters["hostname"].Value = Program.machine;
                Cmd.Parameters["searchcondit"].Value = str;
                Cmd.Parameters["searchresult"].Value = count;

                SqlDataReader dr = Cmd.ExecuteReader();
            }

            //TODO
            //Application.Exit();
        }

        public static DataTable replaceDataTable(DataTable dt)
        {
            //変更前
            DataTable retDt = new DataTable();
            DataRow row = null;
            try
            {
                // 戻り値のDataTable作成
                retDt.Columns.Add((string)dt.Columns[0].ColumnName, typeof(string));
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    retDt.Columns.Add(Convert.ToString(dt.Rows[j].ItemArray[0]), typeof(decimal));
                }

                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    row = retDt.NewRow();
                    row[Convert.ToString(dt.Columns[0].ColumnName)] = dt.Columns[i].ColumnName;
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        //row[Convert.ToString(dt.Rows[j].ItemArray[0])] = Convert.ToDecimal(dt.Rows[j].ItemArray[i]);
                        var obj = dt.Rows[j].ItemArray[i];
                        row[Convert.ToString(dt.Rows[j].ItemArray[0])] =
                        (obj == DBNull.Value || obj == null || string.IsNullOrWhiteSpace(obj.ToString()))
                                                        ? 0m
                                                        : Convert.ToDecimal(obj);
                    }

                    retDt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return retDt;
        }

        public static DataTable GetDB(string sql)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(ODIS.Com.SQLConstr))
                {
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

        public static DataTable SetBulkCopy(DataTable dt, string tname)
        {
            SqlConnection Cn;

            try
            {
                using (Cn = new SqlConnection(ODIS.Com.SQLConstr))
                {
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(Cn))
                    {
                        bulkcopy.BulkCopyTimeout = 660;
                        bulkcopy.DestinationTableName = tname;
                        bulkcopy.WriteToServer(dt);
                        bulkcopy.Close();
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

        public static DataTable GetPosDB(string sql)
        {
            DataTable dt = new DataTable();
            int nRet;
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr))
                {
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
                    nRet = adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return dt;
        }

        public static int GetLevel(decimal kei)
        {
            if (kei == 0)
            {
                return 0;
            }
            else if (kei >= 100)
            {
                return 1;
            }
            else if (kei >= 90 && kei < 100)
            {
                return 2;
            }
            else if (kei >= 85 && kei < 90)
            {
                return 3;
            }
            else if (kei >= 80 && kei < 85)
            {
                return 4;
            }
            else if (kei >= 70 && kei < 80)
            {
                return 5;
            }
            else if (kei >= 60 && kei < 70)
            {
                return 6;
            }
            else if (kei >= 50 && kei < 60)
            {
                return 7;
            }
            else if (kei >= 50 && kei < 60)
            {
                return 7;
            }
            else if (kei > 0 && kei < 50)
            {
                return 8;
            }
            else 
            {
                return 9;
            }
        }

        public static void GetLevelDisp(DataGridViewCellFormattingEventArgs e, decimal val)
        {

            if (e.RowIndex == 9 || e.RowIndex == 20 || e.RowIndex == 32)
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

        public static void GetLevelDispPCA(DataGridViewCellFormattingEventArgs e, decimal val)
        {

            if (e.RowIndex == 8 || e.RowIndex == 31)
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

        /// <summary>
        /// officeバージョン取得
        /// </summary>
        /// <returns></returns>
        public static bool OfficeCK()
        {
            // Excelアプリケーションに接続
            Type classType = Type.GetTypeFromProgID("Excel.Application");
            object app = Activator.CreateInstance(classType);

            if (app == null)
            {
                // 未インストールの場合
                MessageBox.Show("office未インストール");
                return false;
            }

            // バージョンを取得
            object versionObj = app.GetType().InvokeMember(
                "Version", BindingFlags.GetProperty, null, app, null);
            string version = versionObj.ToString();

            if ("11.0".Equals(version))
            {
                // Office2003インストール済みの場合
                return false;
            }
            else if ("12.0".Equals(version))
            {
                // Office2007インストール済みの場合
                return false;
            }
            else if ("14.0".Equals(version))
            {
                // Office2010インストール済みの場合
                return true;
            }
            else if ("15.0".Equals(version))
            {
                // Office2013インストール済みの場合
                return true;
            }
            else
            {
                MessageBox.Show("offceVer例外 ODIS管理者へ連絡願います。");
                return true;
            }

        }

        public static void CalcTuukin(string howto, decimal kyori, decimal kinmu, ref decimal tanka, ref decimal hi, ref decimal ka)
        {
            decimal flg = 0;
            decimal kotei = 0;

            switch (howto)
            {
                case "1 車": flg = 1; kotei = 150; break;
                case "2 バイク": flg = 1; kotei = 150; break;
                //case "3 徒歩・自転車": flg = 1; kotei = 100; break;
                case "4 バス・モノレール": flg = 1; kotei = 300; break;
                case "5 送迎(会社)": flg = 0; kotei = 0; break;
                case "6 送迎(知人・親族)": flg = 1; kotei = 150; break;
                case "7 業務車両": flg = 0; kotei = 0; break;
                case "8 徒歩": flg = 1; kotei = 100; break;
                case "9 自転車": flg = 1; kotei = 100; break;
                default: break;
            }

            //通勤1日単価
            tanka = kyori * 30 * flg + kotei;

            //40overの場合の単価
            if (kyori > 40) tanka = 40 * 30 * flg + kotei;

            //概算通勤手当総額
            decimal dec = tanka * kinmu;


            //通勤1日単価
            tanka = kyori < 1 ? 0 : tanka;

            if (kyori < 1)
            {
                hi = 0;
                ka = 0;
            }
            else if (kyori < 2) //0
            {
                hi = 0;
                ka = dec;
            }
            else if (kyori < 10) //4200
            {
                hi = dec > 4200 ? 4200 : dec;
                ka = dec > 4200 ? dec - 4200 : 0;
            }
            else if (kyori < 15)
            {
                hi = dec > 7100 ? 7100 : dec;
                ka = dec > 7100 ? dec - 7100 : 0;
            }
            else if (kyori < 25)
            {
                hi = dec > 12900 ? 12900 : dec;
                ka = dec > 12900 ? dec - 12900 : 0;
            }
            else if (kyori < 35)
            {
                hi = dec > 18700 ? 18700 : dec;
                ka = dec > 18700 ? dec - 18700 : 0;
            }
            else if (kyori < 45)
            {
                hi = dec > 24400 ? 24400 : dec;
                ka = dec > 24400 ? dec - 24400 : 0;
            }
            else if (kyori < 55)
            {
                hi = dec > 28000 ? 28000 : dec;
                ka = dec > 28000 ? dec - 28000 : 0;
            }
            else
            {
                hi = dec > 31600 ? 31600 : dec;
                ka = dec > 31600 ? dec - 31600 : 0;
            }

            //バス、モノレールは全額非課税に
            if (howto == "4 バス・モノレール")
            {
                hi = kyori < 1 ? 0 : dec;
                ka = 0;
            }
        }

        public static void GetSyukkinbo(DataTable dt, DateTime ym, bool flg, bool excel)
        {

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //施設:false
            if (flg)
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\30_出勤簿.xlsx");
            }
            else
            {
                //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\30_出勤簿_施設警備.xlsx");

            }


            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];

            //label6.Text = "まだぐるぐるします。";

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                }
            }

            string sheetname = ym.ToString("yyyyMM");

            XLSheet ws = c1XLBook1.Sheets[sheetname];

            //出勤簿テンプレートシート
            //XLSheet ws = c1XLBook1.Sheets["出勤簿"];

            int MM = ym.Month;
            //YYYY年M月分 (M月給与)
            ws[0, 1].Value = ym.ToString("yyyy年M月 ") + ym.AddMonths(1).ToString(" (M月給与分)");

            for (int i = 1; i <= rows; i++)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = i.ToString();   // クローンをリネーム
                newSheet[0, 15].Value = i;      // 値の変更
                c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
            }

            //空
            if (rows == 0)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = "空";   // クローンをリネーム
                newSheet[0, 15].Value = "";      // 値の変更
                c1XLBook1.Sheets.Add(newSheet); // クローンをブックに追加
            }

            // テンプレートシートを削除
            //TODO 毎年変更
            c1XLBook1.Sheets.Remove("202501");
            c1XLBook1.Sheets.Remove("202502");
            c1XLBook1.Sheets.Remove("202503");
            c1XLBook1.Sheets.Remove("202504");
            c1XLBook1.Sheets.Remove("202505");
            c1XLBook1.Sheets.Remove("202506");
            c1XLBook1.Sheets.Remove("202507");
            c1XLBook1.Sheets.Remove("202508");
            c1XLBook1.Sheets.Remove("202509");
            c1XLBook1.Sheets.Remove("202510");
            c1XLBook1.Sheets.Remove("202511");
            c1XLBook1.Sheets.Remove("202512");
            c1XLBook1.Sheets.Remove("202601");
            c1XLBook1.Sheets.Remove("202602");
            c1XLBook1.Sheets.Remove("202603");


            string localPass = @"C:\ODIS\SyukkinBo\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");

            //label6.Text = "まだぐるぐるします。";

            if (excel)
            {
                System.Diagnostics.Process.Start(exlName + ".xlsx");
            }
            else
            { 
            //Excel Change PDF           
            Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
            m_MyExcel.Visible = false; //エクセルを非表示
            m_MyExcel.DisplayAlerts = false; //アラート非表示
            Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
            //Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト

            //ブックを開く
            m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");

            //label6.Text = "もう少しですが、こっから長いです。。";

            //PDF保存
            m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");


            m_MyBook.Close(false);
            m_MyExcel.Quit();

            //excel出力
            //System.Diagnostics.Process.Start(@"c:\temp\test2.xlsx");
            //PDF出力
            System.Diagnostics.Process.Start(exlName + ".pdf");
            }
        }

        public static string RitouCalc(string bumoncd, string syokusyuname)
        {
            string ritou = "0";
            if (bumoncd == "" || syokusyuname == "")
            {
                return ritou.ToString();
            }

            DataTable dt = Com.GetDB(" select (select 離島_職種 from dbo.K_職務給_職種 where 備考 = '" + syokusyuname + "') + (select 地区 from dbo.r離島手当_地区 where code = '" + bumoncd.Substring(0, 1) + "') as 離島手当 ");
            ritou = dt.Rows[0][0].ToString();

            if (ritou == "") ritou = "0";
            return ritou;
        }


        //public static string RitouCalc(string bumoncd, string genbacd, string syokusyuname)
        //{
        //    int ritou = 0;

        //    if (bumoncd == "" || genbacd == "" || syokusyuname == "" )
        //    {
        //        return ritou.ToString();
        //    }

        //    if (bumoncd.Substring(0, 1) == "3") ritou += 5000; //八重山
        //    if (bumoncd.Substring(0, 1) == "6") ritou += 10000; //宮古
        //    if (bumoncd.Substring(0, 1) == "7") ritou += 5000; //宮古

        //    //switch (genbacd)
        //    //{
        //    //    case "10207": ritou += 30000; break; //渡嘉敷
        //    //    case "10172": ritou += 5000; break; //久米島病院
        //    //    default: break;
        //    //}

        //    //switch (bumoncd)
        //    //{
        //    //    case "22027": ritou += 5000; break; //久米島エンジ
        //    //    case "22028": ritou += 5000; break; //久米島エンジ
        //    //    default: break;
        //    //}

        //    if (ritou == 0) return ritou.ToString();

        //    switch (syokusyuname)
        //    {
        //        case "現業": ritou += 5000; break;
        //        case "客室": ritou += 10000; break;
        //        case "施設": ritou += 20000; break;
        //        case "警備": ritou += 10000; break;
        //        case "飲食": ritou += 0; break;
        //        case "サービス": ritou += 0; break;
        //        case "管理事務": ritou += 0; break;
        //        case "エンジニア": ritou += 20000; break;
        //        case "電気主任C": ritou += 20000; break;
        //        case "電気主任B": ritou += 20000; break;
        //        case "電気主任A": ritou += 20000; break;
        //        case "植栽": ritou += 20000; break;
        //        default: break;
        //    }



        //    return ritou.ToString();
        //}


        /// <summary>
        /// バッチコマンドを実行する
        /// </summary>
        /// <param name="executeCommand">コマンド文字列</param>
        /// <param name="stdOut">標準出力</param>
        /// <param name="stdErr">標準エラー出力</param>
        /// <param name="exitCode">終了コード</param>
        /// <returns>リターンコード</returns>
        public static int ExecBatProcess(string executeCommand, out string stdOut, out string stdErr, out int exitCode)
        {
            stdOut = "";
            stdErr = "";
            exitCode = 0;

            try
            {
                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd.exe", "/c " + executeCommand);
                processStartInfo.CreateNoWindow = true;
                processStartInfo.UseShellExecute = false;

                processStartInfo.RedirectStandardOutput = true;
                processStartInfo.RedirectStandardError = true;

                process = Process.Start(processStartInfo);
                process.WaitForExit();

                //エラーメッセージがでる
                stdOut = process.StandardOutput.ReadToEnd();

                stdErr = process.StandardError.ReadToEnd();
                exitCode = process.ExitCode;

                process.Close();

                return 0;
            }
            catch
            {
                stdErr = executeCommand + " catchエラー";
                return 16;
            }
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
        public static void OutPutCSV(DataTable dt, bool hasHeader, string separator, string quote, string replace, string pass)
        {
            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;
            string text;

            //保存用のファイルを開く。上書きモードで。
            StreamWriter writer = new StreamWriter(pass, false, Encoding.GetEncoding("shift_jis"));

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

            //MessageBox.Show(rows.ToString() + "件出力しました。" + nl + pass);

        }
        #endregion


    }
}
