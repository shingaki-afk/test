using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using ODIS.ODIS;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Text;
using System.Net.Mail;
using System.Net;

namespace ODIS
{
    public partial class Client : Form
    {
        /// <summary>
        /// 改行コード
        /// </summary>
        public string nl = Environment.NewLine;

        /// <summary>
        /// 社員情報+勤怠データ(登録フラグ有　提出担当有)
        /// </summary>
        private DataTable dt;

        /// <summary>
        /// 対象期間インスタンス
        /// </summary>
        private TargetDays td = new TargetDays();

        /// <summary>
        /// 年間最低休日インスタンス
        /// </summary>
        private SyoteiDays sd = new SyoteiDays();

        /// <summary>
        /// 年間最低休日インスタンス
        /// </summary>
        private SyoteiDays_dns sd_dns = new SyoteiDays_dns();


        /// <summary>
        /// 共通クラスのインスタンス
        /// </summary>
        private Common co = new Common();

        private int dgvListCol = 0;
        private int dgvListRow = 0;
        private int dgvCol = 0;
        private int dgvRow = 0;


        private string yakusyoku = "";
        private string kubuncode = "";
        
        
        //登録フラグ
        private string hidden = "";
        private string gengoukei = "";
        private string hituyouDay = "";

        //勤怠状況テーブルデータ
        private DataTable kintaiinfodt = new DataTable();

        /// <summary>
        /// 年休データテーブル
        /// </summary>
        private DataTable nenkyuudt;

        //チェックデータテーブル
        private DataTable ckdt = new DataTable();


        decimal syoteisuu = 0; //所定
        decimal kyuuzisuu = 0; //休日
        decimal houteisuu = 0; //法定休日
        decimal roudzikan = 0; //労働時間


        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Client(string loginID)
        {
            //コントロール初期設定
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);
            dataGridView2.Font = new Font(dataGridView1.Font.Name, 9);
            dataGridView3.Font = new Font(dataGridView1.Font.Name, 11);

            // 選択モードを行単位での選択のみにする
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //現在の対象年月を表示
            TargetDays td = new TargetDays();

            //エラーメッセージは赤に
            ErrorMsg.ForeColor = Color.Red;

            GetData();

            //年休データを取得
            nenkyuudt = co.nenkyuudt;

            combotoku.Items.Add("");
            combotoku.Items.Add("本人結婚(4日)");
            combotoku.Items.Add("一親等死亡(3日)");
            combotoku.Items.Add("一親等結婚(1日)");
            //combotoku.Items.Add("学校臨時休校に伴う休");
            combotoku.Items.Add("遠隔地赴任/帰任(4日)");

            combomutoku.Items.Add("");
            combomutoku.Items.Add("産前産後育休");
            combomutoku.Items.Add("業務上負傷(通勤除く)");
            combomutoku.Items.Add("感染症予防");
            combomutoku.Items.Add("台風等災害時");
            combomutoku.Items.Add("社外教育等");
            combomutoku.Items.Add("公民権行使");
            combomutoku.Items.Add("生理休暇");
            combomutoku.Items.Add("介護休業");
            //combomutoku.Items.Add("学校休業");
            //combomutoku.Items.Add("会社都合(コロナ影響)");

            RestReason.Items.Add("");
            RestReason.Items.Add("シフトで1日8時間超勤務");
            RestReason.Items.Add("土日祝日年始が休日");
            RestReason.Items.Add("その他 ※個別メモに理由を入力ください");

            ooireason.Items.Add("");
            ooireason.Items.Add("シフトで1日8時間未満勤務がある現場");
            ooireason.Items.Add("次月振休取得予定");
            ooireason.Items.Add("その他 ※個別メモに理由を入力ください");

            kyuugyouriyuu.Items.Add("");
            kyuugyouriyuu.Items.Add("退職に伴う有給消化");
            kyuugyouriyuu.Items.Add("メンタルヘルス(心の健康状態)不調");
            kyuugyouriyuu.Items.Add("体の健康状態不調");
            kyuugyouriyuu.Items.Add("その他");


            //データグリッドビューの背景色変更
            this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.White;
            this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;

            this.DGView.RowsDefaultCellStyle.BackColor = Color.White;
            this.DGView.AlternatingRowsDefaultCellStyle.BackColor =Color.Beige;

            //ソート不可対応
            foreach (DataGridViewColumn c in dataGridView1.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;


            //(0, 0)を現在のセルにする
            dataGridView1.CurrentCell = dataGridView1[0, 0];
            //完了ボタン、Excel出力ボタン等の処理
            SetCheck();

            //所定日数等の取得
            //this.syoteinissuu.Text = sd.SyoteiDay.ToString() + "日"; //所定労働日数
            //this.kyuuzitsu.Text = sd.KyuuzitsuDay.ToString() + "日"; //休日数
            //this.syoteiroudouzikan.Text = sd.RoudouH.ToString() + "時間"; //所定労働時間
            //this.houteiday.Text = sd.HoukyuuDay.ToString() + "日"; //法定休日数
            //this.monthtotal.Text = sd.MtotalDay.ToString() + "日";

            //this.kizyun5.Text = sd.Kizyun5.ToString();
            //this.kizyun4.Text = sd.Kizyun4.ToString();
            //this.kizyun3.Text = sd.Kizyun3.ToString();
            //this.kizyun2.Text = sd.Kizyun2.ToString();
            //this.kizyun1.Text = sd.Kizyun1.ToString();

            Com.InHistory("33_勤怠入力", "", "");
        }

        /// <summary>
        /// データ取得メソッド
        /// </summary>
        /// <returns></returns>
        private void GetData()
        {
            //全対象者データ取得
            //[dbo].[k勤怠社員情報取得]
            dt = co.GetKintaiKihon(1, "");

            //担当別データ
            //[dbo].[担当別一覧表示]
            //DataTable list = co.GetKintaiKihon(9, "");

            string sql = "";
            sql = "select 担当区分 as 区分,担当管理 as 担当,SUM(登録フラグ) as 提出数, null as エラー, null as 警告,COUNT(社員番号) - SUM(登録フラグ) as 未提出 from dbo.全給与対象者情報 ";

            if (Program.loginname == "緑間　拓也" || Program.loginname == "親泊　美和子")
            {
                sql += "where 担当管理 = '緑間　拓也' or (担当区分 = '15_久米島' and 担当事務 = '03_施設') or (担当区分 = '14_宮古島' and 担当事務 = '03_施設')";
                //sql += "where 担当事務 like '%03_施設%' or 担当事務 like '%03_警備%' ";
            }
            //警備
            else if (Program.loginname == "吉里　勇二郎" || Program.loginname == "島　辰之")
            {
                sql += "where 担当管理 = '吉里　勇二郎' or (担当区分 = '15_久米島' and 担当事務 = '03_警備') or (担当区分 = '14_宮古島' and 担当事務 = '03_警備') ";
            }
            else if (Program.loginbusyo == "01_現業")
            {
                sql += "where 担当区分 like '%" + Program.loginbusyo + "%' or (担当区分 = '14_宮古島' and 担当事務 = '01_現業') or (担当区分 = '15_久米島' and 担当事務 = '01_現業')";
            }
            else if (Program.loginname == "宮城　一禎")
            {
                sql += "where 担当区分 like '%" + Program.loginbusyo + "%' or 担当区分 = '04_エンジ' ";
            }
            //TOD0 2023/02/09
            else if (Program.loginname == "太田　朋宏")
            {
                sql += "where 担当管理 = '太田　朋宏' ";
            }
            else if (Program.loginname == "中真　心")
            {
                sql += "where 担当区分 = '11_北部' or 担当区分 = '01_現業' or 担当区分 = '02_客室'";
            }

            //北部PPP/PFI対応
            else if (Program.loginbusyo == "11_北部")
            {
                sql += "where 担当区分 like '%" + Program.loginbusyo + "%' and 担当事務 <> '05_PPP/PFI'";
            }
            else if (Program.loginbusyo == "05_PPP/PFI")
            {
                sql += "where 担当区分 like '%" + Program.loginbusyo + "%' or 担当事務 = '05_PPP/PFI'";
            }


            else
            {
                sql += "where 担当区分 like '%" + Program.loginbusyo + "%' ";
            }

            sql += "group by 担当区分, 担当管理 order by 担当区分, 担当管理";

            DataTable list = Com.GetDB(sql);

            string filtStr = "";
            DataRow[] drYae;
            int errorCt = 0; //エラー件数
            int emergCt = 0; //警告件数
            string[] st;

            //エラー数と警告数を取得しリストに表示
            foreach (DataRow dr in list.Rows)
            {
                //担当管理≒名前
                //担当区分≒区分≒部門
                filtStr = "担当管理 = '" + dr["担当"].ToString() + "' and 担当区分 = '" + dr["区分"].ToString() + "' and 登録フラグ = '1'";
                //filtStr = "担当管理 = '" + dr["担当"].ToString() + "' and 登録フラグ = '1'";


                drYae = dt.Select(filtStr, "");

                errorCt = 0;
                emergCt = 0;

                foreach (DataRow row in drYae)
                {
                    st = co.ErrorCheck(row, "");

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

            foreach (DataRow d in list.Rows)
            {
                if (d[3].ToString() == "0" && d[5].ToString() == "0")
                {
                    KintaiInfo("入力", d[1].ToString());
                }
            }

            //勤怠進捗状況表示
            GetKintaiInfo();

        }

        /// <summary>
        /// データグリッドに一覧表示
        /// </summary>
        private void DisplayGridView()
        {
            //リスト表示の選択セルを指定
            dataGridView1.CurrentCell = dataGridView1[dgvListCol, dgvListRow];

            //対象者表示
            GetListDetail(dgvListCol, dgvListRow, Convert.ToString(dataGridView1.Rows[dgvListRow].Cells[1].Value), Convert.ToString(dataGridView1.Rows[dgvListRow].Cells[0].Value));            
            
            //対象者表示の選択セルを指定
            if (DGView.RowCount == 0)
            {
                AllClear();
            }
            else if (DGView.RowCount == dgvRow)
            {
                DGView.CurrentCell = DGView[dgvCol, dgvRow - 1];

                DataGridViewRow dgr = DGView.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                DataDisp(drv[0].ToString());
            }
            else
            {
                DGView.CurrentCell = DGView[dgvCol, dgvRow];

                DataGridViewRow dgr = DGView.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                DataDisp(drv[0].ToString());
            }
        }

        private void AllClear()
        {
            //TODO　再表示の時に利用していない。。
            number.Text = "";
            entyou.Text = "";
            houteikyuuH.Text = "";
            syoteikyuuH.Text = "";
            zangyouH.Text = "";
            chouzanH.Text = "";
            shinyaH.Text = "";
            //chikoku.Text = "";
            chikokuH.Text = "";
            syotei.Text = "";
            houteikyuu.Text = "";
            syoteikyuu.Text = "";
            yuukyuu.Text = "";
            tokkyuu.Text = "";
            mutoku.Text = "";

            //boushi.Text = "";
            //kyuugyou.Text = "";
            //sonota.Text = "";


            hurikyuu.Text = "";
            koukyuu.Text = "";
            //choukyuu.Text = "";
            todokede.Text = "";
            mutodoke.Text = "";
            kaisuu1.Text = "";
            kaisuu2.Text = "";

            name.Text = "";
            nyuusya.Text = "";
            taisyoku.Text = "";
            kubun.Text = "";

            syokusyu.Text = "";
            genba.Text = "";

            ErrorMsg.Text = "";
            WarningMsg.Text = "";

            zanDays.Text = "";

            kaisuu1.Enabled = true;
            kaisuu2.Enabled = true;
            //choukyuu.Enabled = true;

            //mutoku.Enabled = true;
            //mutoku.Enabled = false;
            //boushi.Enabled = true;
            //kyuugyou.Enabled = true;
            //sonota.Enabled = true;
            
            entyou.Enabled = true;
            chouzanH.Enabled = true;

            //追加分
            houteikyuuH.Enabled = true;
            houteikyuu.Enabled = true;
            syoteikyuuH.Enabled = true;
            syoteikyuu.Enabled = true;
            zangyouH.Enabled = true;
            shinyaH.Enabled = true;

            //再追加
            //chikoku.Enabled = true;
            chikokuH.Enabled = true;
            kaisuu1.Enabled = true;
            kaisuu2.Enabled = true;

            //再々追加
            tokkyuu.Enabled = true;
            hurikyuu.Enabled = true;

            //特休理由選択
            //combotoku.Enabled = true;

            yuukyuu.Enabled = true;


            //先月データのクリア
            label50.Text = "";
            label18.Text = "";
            label19.Text = "";
            label41.Text = "";
            //label42.Text = "";
            label43.Text = "";
            //label44.Text = "";
            label45.Text = "";
            label46.Text = "";
            label47.Text = "";
            label48.Text = "";
            label49.Text = "";
            label51.Text = "";
            label52.Text = "";
            label53.Text = "";
            label54.Text = "";
            //label55.Text = "";
            label56.Text = "";
            label57.Text = "";
            label58.Text = "";
            label59.Text = "";

            //hiddenデータのクリア
            
            //overflg = "";
            yakusyoku = "";
            kubuncode = "";
            hidden = "";
            gengoukei = "";
            hituyouDay = "";

            yuuZan.Text = "";
            zikyuu.Text = "";
            nikkyuu.Text = "";
            yakin.Text = "";
            syukutyoku.Text = "";
            koyou.Text = "";
            syakai.Text = "";
            kinmuh.Text = "";

            zanDays.Text = "";
            zitsuroudouH.Text = "";

            roudouD.Text = "";

            jyuumin.Text = "";

            yuukyuuzyoukyou.Text = "";
            //yuukyuuhissuzan.Text = "";

            kaisu1reason.Text = "";
            kaisu2reason.Text = "";

            toketsu.Text = "";
        }

        /// <summary>
        /// グリッドビュー選択時イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DGView_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = DGView.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            DataDisp(drv[0].ToString());
        }

        /// <summary>
        /// 半券画面へデータ表示
        /// </summary>
        /// <param name="str"></param>
        private void DataDisp(string str)
        {
            //AllClear();

            //先月末退職者(当月勤怠無者)の登録可能対応
            button1.Enabled = true;

            //社員番号を表示
            this.number.Text = str;

            //対象者のみに絞込
            DataRow[] targetDr = dt.Select("社員番号 = " + str, "");

            //
            yuuZan.Text = "";

            yuukyuuzyoukyou.Text = "";
            //yuukyuuhissuzan.Text = "";

            //20160105　前月有給が残る
            this.label49.Text = "";

            //理由選択クリア
            combotoku.SelectedItem = "";
            combomutoku.SelectedItem = "";
            RestReason.SelectedItem = "";
            ooireason.SelectedItem = "";
            kyuugyouriyuu.SelectedItem = "";

            //DataRow[] nenkyuudr = nenkyuudt.Select("社員番号 = " + str, "");

            //前月の勤怠データを取得
            DataTable dtzen = GetKakoData();
            
            //TODO 要確認　
            //テーブルデータ無はスルーする
            if (dtzen.Rows.Count > 0)
            {
                foreach (DataRow row in dtzen.Rows)
                {
                    switch (row[1].ToString())
                    {
                        case "延長時間": this.label50.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "法休時間": this.label18.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "所休時間": this.label19.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "残業時間": this.label41.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        //case "60超残Ｈ": this.label42.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "深夜時間": this.label43.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        //case "遅刻回数": this.label44.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "遅刻時間": this.label45.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "所定": this.label46.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "法休": this.label47.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "所休": this.label48.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "有給": this.label49.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "特休": this.label51.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "生休": this.label52.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "振休": this.label53.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "公休": this.label54.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        //TODO 2019/02月のためだけに残します。以降はラベルも不要です。
                        case "調休": this.label75.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "届欠": this.label56.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "無届": this.label57.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "回数１": this.label58.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        case "回数２": this.label59.Text = Convert.ToDecimal(row["勤怠データ値"]).ToString("#,##0.##;-#,##0.##;#"); break;
                        default: break;
                    }
                }
            }

            //TODO 2019/02月のためだけに残します。以降はラベルも不要です。
            if (this.label54.Text != "" && this.label75.Text != "")
            this.label54.Text = (Convert.ToDecimal(this.label54.Text) + Convert.ToDecimal(this.label75.Text)).ToString();

            //対象者の年休データを取得
            DataRow[] nenkyuudr = nenkyuudt.Select("社員番号 = " + str, "");
            foreach (DataRow r in nenkyuudr)
            {
                yuuZan.Text = r[1].ToString().Replace(".000", "").Replace(".500", ".5");
                if (r[1].ToString() != "") yuuZan.Text += " 日";

                zikyuu.Text = r[2].ToString().Replace(".000", "");
                if (r[2].ToString() != "") zikyuu.Text += " 円";

                nikkyuu.Text = r[3].ToString().Replace(".000", "");
                if (r[3].ToString() != "") nikkyuu.Text += " 円";

                yakin.Text = r[4].ToString().Replace(".000", "");
                if (r[4].ToString() != "") yakin.Text += " 円";


                //TODO 12月勤怠と1月勤怠だけ1000円はいる！
                if (td.StartYMD.ToString("MM") == "12" || td.StartYMD.ToString("MM") == "01")
                {
                    syukutyoku.Text = "1000 円";
                    kaisu2reason.Text = "正月手当";
                }
                //    syukutyoku.Text = r[5].ToString().Replace(".000", "");
                //if (r[5].ToString() != "") syukutyoku.Text += " 円";

                kinmuh.Text = r[6].ToString().Replace(".000", "");
                if (r[6].ToString() != "") kinmuh.Text += " 時間";

                if (r[7].ToString() == "9")
                {
                    koyou.Text = "";
                }
                else
                {
                    koyou.Text = "〇";
                }

                if (r[8].ToString() != "")
                {
                    syakai.Text = "〇";
                }
                else if (r[6].ToString().Replace(".000", "") == "8")
                {
                    syakai.Text = "-";
                }
                else
                {
                    syakai.Text = "";
                }

                roudouD.Text = co.GetKyuukaKubun[r[9].ToString()];

                //10 状況
                //11 必須残
                yuukyuuzyoukyou.Text = r[10].ToString();
                //yuukyuuhissuzan.Text = r[11].ToString();

                kaisu1reason.Text = r[12].ToString();

                
                //kaisu2reason.Text = r[13].ToString();
            }

            foreach (DataRow row in targetDr)
            {
                //入社日チェック必要?
                if (row[30].Equals(DBNull.Value)) return;

                string defH = "0.0";
                string defD = "0";

                entyou.Text = row[1].Equals(DBNull.Value) ? defH : row[1].ToString();   //延長h
                houteikyuuH.Text = row[2].Equals(DBNull.Value) ? defH : row[2].ToString(); //法定休h
                syoteikyuuH.Text = row[3].Equals(DBNull.Value) ? defH : row[3].ToString(); //所定休h              
                zangyouH.Text = row[4].Equals(DBNull.Value) ? defH : row[4].ToString();  //総残業h
                chouzanH.Text = row[5].Equals(DBNull.Value) ? defH : row[5].ToString();  //内60超残h
                shinyaH.Text = row[6].Equals(DBNull.Value) ? defH : row[6].ToString();   //深夜勤h
                //chikoku.Text = row[7].Equals(DBNull.Value) ? defD : row[7].ToString();   //遅刻回数
                chikokuH.Text = row[8].Equals(DBNull.Value) ? defH : row[8].ToString();  //遅刻時間
                syotei.Text = row[9].Equals(DBNull.Value) ? defD : row[9].ToString();    //所定
                houteikyuu.Text = row[10].Equals(DBNull.Value) ? defD : row[10].ToString(); //法定休
                syoteikyuu.Text = row[11].Equals(DBNull.Value) ? defD : row[11].ToString(); //所定休
                yuukyuu.Text = row[12].Equals(DBNull.Value) ? defD : row[12].ToString(); //有給
                tokkyuu.Text = row[13].Equals(DBNull.Value) ? defD : row[13].ToString(); //特休
                mutoku.Text = row[14].Equals(DBNull.Value) ? defD : row[14].ToString(); //無特

                //boushi.Text = row["防止日数"].Equals(DBNull.Value) ? defD : row["防止日数"].ToString(); 
                //kyuugyou.Text = row["休業日数"].Equals(DBNull.Value) ? defD : row["休業日数"].ToString();
                //sonota.Text = row["その他日数"].Equals(DBNull.Value) ? defD : row["その他日数"].ToString();

                hurikyuu.Text = row[15].Equals(DBNull.Value) ? defD : row[15].ToString(); //振休
                koukyuu.Text = row[16].Equals(DBNull.Value) ? defD : row[16].ToString(); //公休
                //choukyuu.Text = row[17].Equals(DBNull.Value) ? defD : row[17].ToString(); //調休
                todokede.Text = row[17].Equals(DBNull.Value) ? defD : row[17].ToString(); //届出
                mutodoke.Text = row[18].Equals(DBNull.Value) ? defD : row[18].ToString(); //無届
                kaisuu1.Text = row[19].Equals(DBNull.Value) ? defD : row[19].ToString(); //回数1
                kaisuu2.Text = row[20].Equals(DBNull.Value) ? defD : row[20].ToString(); //回数2

                name.Text = row[21].ToString(); //名前
                nyuusya.Text = row[30].ToString(); //入社日

                kyuuzitsukubun.Text = row[58].ToString(); //休日区分


                //TODO 処理がかぶっている!!
                syoteisuu = Convert.ToDecimal(row[59].ToString()); //所定
                kyuuzisuu = Convert.ToDecimal(row[60].ToString()); //休日
                houteisuu = Convert.ToDecimal(row[61].ToString()); //法定休日
                roudzikan = Convert.ToDecimal(row[62].ToString()); //労働時間


                string kariflg = "";

                ////退職マスタ対応　退職日がはいってなければ仮退職日を表示
                if (row[31].Equals(DBNull.Value))
                {
                    taisyoku.Text = row["退職日"].ToString();
                    kariflg = "kari";
                }
                else
                {
                    taisyoku.Text = row[31].ToString();
                }

                //taisyoku.Text = row[32].Equals(DBNull.Value) ? row["退職日"].ToString() : row[32].ToString(); //退職日

                kubun.Text = co.GetKubunName[row[29].ToString()]; //給与区分
                kubuncode = row[29].ToString();
                hidden = row[36].ToString();

                combotoku.SelectedItem = row["特休理由"].Equals(DBNull.Value) ? defD : row["特休理由"].ToString(); //特休理由
                combomutoku.SelectedItem = row["無特理由"].Equals(DBNull.Value) ? defD : row["無特理由"].ToString(); //無特理由
                RestReason.SelectedItem = row["休日超過理由"].Equals(DBNull.Value) ? defD : row["休日超過理由"].ToString(); //無特理由
                ooireason.SelectedItem = row["出勤超過理由"].Equals(DBNull.Value) ? defD : row["出勤超過理由"].ToString(); //無特理由
                kyuugyouriyuu.SelectedItem = row["備考"].Equals(DBNull.Value) ? defD : row["備考"].ToString(); //休業理由
                jyuumin.Text = row["住民税"].ToString().Replace(".000", "");

                toketsu.Text = row[63].Equals(DBNull.Value) ? defD : row[63].ToString(); //途欠

                
                textBox1.Text = row["コメント"].ToString();//コメント

                //tiku.Text = row[27].ToString(); //地区
                syokusyu.Text = row[27].ToString(); //職種
                genba.Text = row[28].ToString(); //現場
                               
                //エラーチェック
                string[] st = co.ErrorCheck(row, "");

                ErrorMsg.Text = st[0]; //エラーメッセージ
                hituyouDay = st[1];　 //必要日数
                gengoukei = st[2];　 //合計日数
                //label6.Text = st[3];　 //今月入社・退社表示
                WarningMsg.Text = st[4];

                //当月入社/退社
                //1:入社　2:退社 3:両方
                if (st[3] == "1")
                {
                    nyuusya.ForeColor = Color.Red;
                    taisyoku.ForeColor = Color.Black;
                }
                else if (st[3] == "2")
                {
                    nyuusya.ForeColor = Color.Black;
                    taisyoku.ForeColor = Color.Red;
                }
                else if (st[3] == "3")
                {
                    nyuusya.ForeColor = Color.Red;
                    taisyoku.ForeColor = Color.Red;
                }
                else
                {
                    nyuusya.ForeColor = Color.Black;
                    taisyoku.ForeColor = Color.Black;
                }

                //退職マスタ対応
                if (kariflg == "kari") taisyoku.ForeColor = Color.RoyalBlue;

                //入力不可対応

                //一旦クリア
                kaisuu1.Enabled = true;
                kaisuu2.Enabled = true;
                //choukyuu.Enabled = true;
                //mutoku.Enabled = true;
                //mutoku.Enabled = false;
                //boushi.Enabled = true;
                //kyuugyou.Enabled = true;
                //sonota.Enabled = true;

                entyou.Enabled = true;
                chouzanH.Enabled = true;

                //追加分
                houteikyuuH.Enabled = true;
                houteikyuu.Enabled = true;
                syoteikyuuH.Enabled = true;
                syoteikyuu.Enabled = true;
                zangyouH.Enabled = true;
                shinyaH.Enabled = true;

                //再追加
                //chikoku.Enabled = true;
                chikokuH.Enabled = true;

                //再々追加
                tokkyuu.Enabled = true;
                hurikyuu.Enabled = true;

                yuukyuu.Enabled = true;

                //特休理由
                combotoku.Enabled = true;
                //休日超過理由
                //RestReason.Enabled = false;
                //ooireason.Enabled = false;

                //役職
                yakusyoku = row["役職CD"].ToString();

                //課長職以上
                //課長職:130 係長:135
                if (Convert.ToInt16(yakusyoku) <= 135)
                {
                    //追加分
                    houteikyuuH.Enabled = false;
                    houteikyuu.Enabled = false;
                    syoteikyuuH.Enabled = false;
                    syoteikyuu.Enabled = false;
                    zangyouH.Enabled = false;

                    //深夜手当は原則入力可能
                    //shinyaH.Enabled = false;

                    //再追加
                    //chikoku.Enabled = false;
                    chikokuH.Enabled = false;
                }

                //回数1単価の未登録は入力不可にする
                if (Convert.ToInt16(row[32]) == 0) kaisuu1.Enabled = false;

                //12月勤怠と1月勤怠だけ入力可能
                if (td.StartYMD.ToString("MM") == "12" || td.StartYMD.ToString("MM") == "01")
                {
                    if (row[29].ToString() == "F1")
                    {
                        kaisuu2.Enabled = false;
                    }
                }
                else
                {
                    kaisuu2.Enabled = false;
                }

                //201412 ロワジール対応
                if (genba.Text == "ロワジールホテル那覇" & syokusyu.Text.Substring(0,2) == "客室")
                {
                    kaisuu1.Enabled = false;
                }

                //有給残数無の場合は入力不可にする
                if (yuuZan.Text == "")
                {
                    yuukyuu.Enabled = false;
                }

                //休業理由
                if (syotei.Text == "0.0" || syotei.Text == "0")
                {
                    kyuugyouriyuu.Enabled = true;
                }
                else
                {
                    kyuugyouriyuu.Enabled = false;
                }

                //パートとアルバイトは特休を入力不可にする
                if (row[29].ToString() == "E1" | row[29].ToString() == "F1")
                {
                    //TODO 特休20200401
                    //TODO 特休20251008 入力不可に戻し
                    tokkyuu.Enabled = false;
                    hurikyuu.Enabled = false;

                    //TODO 特休20200401
                    //理由選択も不可にする
                    combotoku.Enabled = false;


                    RestReason.Enabled = false;
                    ooireason.Enabled = false;
                }
                else
                {
                    //日給・月給者
                    //選択済みなら表示
                    if (RestReason.SelectedItem.ToString() != "")
                    {
                        RestReason.Enabled = true;
                    }
                    else
                    {
                        if (roudouD.Text == "３日")
                        {
                            //if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > sd.KyuuzitsuDay + 9)
                            if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > kyuuzisuu + 9)
                            {
                                RestReason.Enabled = true;
                            }
                            else
                            {
                                RestReason.Enabled = false;
                            }
                        }
                        else if (roudouD.Text == "４日")
                        {
                            if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > kyuuzisuu + 4)
                            {
                                RestReason.Enabled = true;
                            }
                            else
                            {
                                RestReason.Enabled = false;
                            }
                        }
                        else
                        {
                            if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > kyuuzisuu)
                            {
                                RestReason.Enabled = true;
                            }
                            else
                            {
                                RestReason.Enabled = false;
                            }
                        }
                    }

                    if (ooireason.SelectedItem.ToString() != "")
                    {
                        ooireason.Enabled = true;
                    }
                    else
                    {
                        if (roudouD.Text == "３日")
                        {
                            if (Convert.ToDecimal(syotei.Text) > syoteisuu + 9)
                            {
                                ooireason.Enabled = true;
                            }
                            else
                            {
                                ooireason.Enabled = false;
                            }
                        }
                        else
                        { 
                            if (Convert.ToDecimal(syotei.Text) > syoteisuu)
                            {
                                ooireason.Enabled = true;
                            }
                            else
                            {
                                ooireason.Enabled = false;
                            }
                        }
                    }
                }

                

                //日給者・月給者は延長を入力不可にする
                if (row[29].ToString() != "E1" & row[29].ToString() != "F1") entyou.Enabled = false;

                //60超は入力不要なので入力不可にする
                chouzanH.Enabled = false;

                //残日数
                if (Convert.ToDecimal(gengoukei) > Convert.ToDecimal(hituyouDay))
                {
                    zanDays.Text = (Convert.ToDecimal(gengoukei) - Convert.ToDecimal(hituyouDay)).ToString() + "日多い　" + "[" + gengoukei + "/" + hituyouDay + "]";
                }
                else if (Convert.ToDecimal(gengoukei) < Convert.ToDecimal(hituyouDay))
                {
                    zanDays.Text = "残り" + (Convert.ToDecimal(hituyouDay) - Convert.ToDecimal(gengoukei)).ToString() + "日　" + "[" + gengoukei + "/" + hituyouDay + "]";
                }
                else
                {
                    zanDays.Text = "[" + gengoukei + "/" + hituyouDay + "]";
                }

                if (kyuuzitsukubun.Text == "20")
                {
                    this.syoteinissuu.Text = sd_dns.SyoteiDay.ToString() + "日"; //所定労働日数
                    this.kyuuzitsu.Text = sd_dns.KyuuzitsuDay.ToString() + "日"; //休日数
                    this.syoteiroudouh.Text = sd_dns.RoudouH.ToString() + "時間"; //所定労働時間
                    this.houteiday.Text = sd_dns.HoukyuuDay.ToString() + "日"; //法定休日数
                    this.monthtotal.Text = sd_dns.MtotalDay.ToString() + "日";

                    this.kizyun5.Text = sd_dns.Kizyun5.ToString();
                    this.kizyun4.Text = sd_dns.Kizyun4.ToString();
                    this.kizyun3.Text = sd_dns.Kizyun3.ToString();
                    this.kizyun2.Text = sd_dns.Kizyun2.ToString();
                    this.kizyun1.Text = sd_dns.Kizyun1.ToString();
                }
                else
                {
                    this.syoteinissuu.Text = sd.SyoteiDay.ToString() + "日"; //所定労働日数
                    this.kyuuzitsu.Text = sd.KyuuzitsuDay.ToString() + "日"; //休日数
                    this.syoteiroudouh.Text = sd.RoudouH.ToString() + "時間"; //所定労働時間
                    this.houteiday.Text = sd.HoukyuuDay.ToString() + "日"; //法定休日数
                    this.monthtotal.Text = sd.MtotalDay.ToString() + "日";

                    this.kizyun5.Text = sd.Kizyun5.ToString();
                    this.kizyun4.Text = sd.Kizyun4.ToString();
                    this.kizyun3.Text = sd.Kizyun3.ToString();
                    this.kizyun2.Text = sd.Kizyun2.ToString();
                    this.kizyun1.Text = sd.Kizyun1.ToString();
                }




                //先月末退職者(当月勤怠無者)の登録可能対応
                if (hituyouDay == "0")
                {
                    button1.Enabled = false;
                }



                moneycolor();
            }
        }

        /// <summary>
        /// 登録・更新ボタンイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //半券情報無場合で登録ボタンクリック時の対応
            if (number.Text == "" || number.Text == "-") return;

            //登録後のフォーカスに利用
            dgvCol = DGView.CurrentCell.ColumnIndex;
            dgvRow = DGView.CurrentCell.RowIndex;

            //入力データのチェック
            ErrorCheck("Up");


             ///エラー有無確認
            if (ErrorMsg.Text != "")
            {
                    MessageBox.Show(ErrorMsg.Text, "登録できません");
            }
            else
            {
                if (WarningMsg.Text != "")
                {
                    DialogResult result = MessageBox.Show(WarningMsg.Text,
                        "警告",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.No) return;
                }

                //insert or update
                if (hidden == "0")
                {
                    InsertCSVData();
                    StatusLabel.Text = name.Text + "さんを登録しました。";
                }
                else
                {
                    UpdateKintai();
                    StatusLabel.Text = name.Text + "さんを登録更新しました。";
                }

                GetData();
                DisplayGridView();
            }

            //フォーカス
            //if (entyou.Enabled)
            //{
            //    entyou.Focus();
            //}
            //else if (houteikyuuH.Enabled)
            //{
            //    houteikyuuH.Focus();
            //}
            //else
            //{
            //    syotei.Focus();
            //}



            syotei.Focus();
            //TODO!
        }

        /// <summary>
        /// データ登録処理
        /// </summary>
        private void InsertCSVData()
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
                    Cmd.CommandText = "[dbo].[Insertkintai]";

                    Cmd.Parameters.Add(new SqlParameter("対象年月", SqlDbType.Date));
                    Cmd.Parameters["対象年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Decimal));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("延長h", SqlDbType.Decimal));
                    Cmd.Parameters["延長h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("法休h", SqlDbType.Decimal));
                    Cmd.Parameters["法休h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所休h", SqlDbType.Decimal));
                    Cmd.Parameters["所休h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("総残h", SqlDbType.Decimal));
                    Cmd.Parameters["総残h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("六十超h", SqlDbType.Decimal));
                    Cmd.Parameters["六十超h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("深夜h", SqlDbType.Decimal));
                    Cmd.Parameters["深夜h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("遅刻回", SqlDbType.Decimal));
                    Cmd.Parameters["遅刻回"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("遅刻h", SqlDbType.Decimal));
                    Cmd.Parameters["遅刻h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所定", SqlDbType.Decimal));
                    Cmd.Parameters["所定"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("法休", SqlDbType.Decimal));
                    Cmd.Parameters["法休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所休", SqlDbType.Decimal));
                    Cmd.Parameters["所休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("有給", SqlDbType.Decimal));
                    Cmd.Parameters["有給"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特休", SqlDbType.Decimal));
                    Cmd.Parameters["特休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無特", SqlDbType.Decimal));
                    Cmd.Parameters["無特"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("振休", SqlDbType.Decimal));
                    Cmd.Parameters["振休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("公休", SqlDbType.Decimal));
                    Cmd.Parameters["公休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("届欠", SqlDbType.Decimal));
                    Cmd.Parameters["届欠"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無届", SqlDbType.Decimal));
                    Cmd.Parameters["無届"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("回数1", SqlDbType.Decimal));
                    Cmd.Parameters["回数1"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("回数2", SqlDbType.Decimal));
                    Cmd.Parameters["回数2"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特休理由", SqlDbType.VarChar));
                    Cmd.Parameters["特休理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無特理由", SqlDbType.VarChar));
                    Cmd.Parameters["無特理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("休日超過理由", SqlDbType.VarChar));
                    Cmd.Parameters["休日超過理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("出勤超過理由", SqlDbType.VarChar));
                    Cmd.Parameters["出勤超過理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("コメント", SqlDbType.VarChar));
                    Cmd.Parameters["コメント"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("最終更新日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["最終更新日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("防止日数", SqlDbType.Decimal));
                    Cmd.Parameters["防止日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("休業日数", SqlDbType.Decimal));
                    Cmd.Parameters["休業日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("その他日数", SqlDbType.Decimal));
                    Cmd.Parameters["その他日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                    Cmd.Parameters["備考"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("途欠", SqlDbType.Decimal));
                    Cmd.Parameters["途欠"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD;
                    Cmd.Parameters["社員番号"].Value = Convert.ToDecimal(number.Text);
                    Cmd.Parameters["延長h"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["法休h"].Value = Convert.ToDecimal(houteikyuuH.Text);
                    Cmd.Parameters["所休h"].Value = Convert.ToDecimal(syoteikyuuH.Text);
                    Cmd.Parameters["総残h"].Value = Convert.ToDecimal(zangyouH.Text);
                    Cmd.Parameters["六十超h"].Value = Convert.ToDecimal(chouzanH.Text);
                    Cmd.Parameters["深夜h"].Value = Convert.ToDecimal(shinyaH.Text);
                    Cmd.Parameters["遅刻回"].Value = 0; //Convert.ToDecimal(chikoku.Text);
                    Cmd.Parameters["遅刻h"].Value = Convert.ToDecimal(chikokuH.Text);
                    Cmd.Parameters["所定"].Value = Convert.ToDecimal(syotei.Text);
                    Cmd.Parameters["法休"].Value = Convert.ToDecimal(houteikyuu.Text);
                    Cmd.Parameters["所休"].Value = Convert.ToDecimal(syoteikyuu.Text);
                    Cmd.Parameters["有給"].Value = Convert.ToDecimal(yuukyuu.Text);
                    Cmd.Parameters["特休"].Value = Convert.ToDecimal(tokkyuu.Text);
                    Cmd.Parameters["無特"].Value = Convert.ToDecimal(mutoku.Text);
                    Cmd.Parameters["振休"].Value = Convert.ToDecimal(hurikyuu.Text);
                    Cmd.Parameters["公休"].Value = Convert.ToDecimal(koukyuu.Text);
                    Cmd.Parameters["届欠"].Value = Convert.ToDecimal(todokede.Text);
                    Cmd.Parameters["無届"].Value = Convert.ToDecimal(mutodoke.Text);
                    Cmd.Parameters["回数1"].Value = Convert.ToDecimal(kaisuu1.Text);
                    Cmd.Parameters["回数2"].Value = Convert.ToDecimal(kaisuu2.Text);
                    Cmd.Parameters["特休理由"].Value = combotoku.SelectedItem.ToString();
                    Cmd.Parameters["無特理由"].Value = combomutoku.SelectedItem.ToString();
                    Cmd.Parameters["休日超過理由"].Value = RestReason.SelectedItem.ToString();
                    Cmd.Parameters["出勤超過理由"].Value = ooireason.SelectedItem.ToString();
                    Cmd.Parameters["コメント"].Value = textBox1.Text;
                    Cmd.Parameters["最終更新日時"].Value = DateTime.Now;
                    Cmd.Parameters["防止日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["休業日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["その他日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["備考"].Value = kyuugyouriyuu.SelectedItem.ToString();
                    Cmd.Parameters["途欠"].Value = Convert.ToDecimal(toketsu.Text);
                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        /// <summary>
        /// 半券表示分のエラーチェック
        /// </summary>
        private void ErrorCheck(string flg)
        {
            DataRow[] targetDr = dt.Select("社員番号 = " + number.Text, "");
            foreach (DataRow row in targetDr)
            {
                row[1] = entyou.Text;
                row[2] = houteikyuuH.Text;
                row[3] = syoteikyuuH.Text;
                row[4] = zangyouH.Text;
                row[5] = chouzanH.Text;
                row[6] = shinyaH.Text;
                row[7] = "0"; // chikoku.Text;
                row[8] = chikokuH.Text;
                row[9] = syotei.Text;
                row[10] = houteikyuu.Text;
                row[11] = syoteikyuu.Text;
                row[12] = yuukyuu.Text;
                row[13] = tokkyuu.Text;
                row[14] = mutoku.Text;
                row[15] = hurikyuu.Text;
                row[16] = koukyuu.Text;
                row[17] = todokede.Text;
                row[18] = mutodoke.Text;
                row[19] = kaisuu1.Text;
                row[20] = kaisuu2.Text;
                row["特休理由"] = combotoku.SelectedItem.ToString();
                row["無特理由"] = combomutoku.SelectedItem.ToString();
                row["休日超過理由"] = RestReason.SelectedItem.ToString();
                row["出勤超過理由"] = ooireason.SelectedItem.ToString();

                row["防止日数"] = 0;
                row["休業日数"] = 0;
                row["その他日数"] = 0;
                row["備考"] = kyuugyouriyuu.SelectedItem.ToString();
                //row["住民税"] = jyuumin.Text;

                string[] st = co.ErrorCheck(row, flg);
                ErrorMsg.Text = st[0]; //エラーメッセージ
                hituyouDay = st[1];　 //必要日数
                gengoukei = st[2];　 //合計日数
                WarningMsg.Text = st[4];  //警告
            }
        }

        /// <summary>
        /// データ更新
        /// </summary>
        private void UpdateKintai()
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
                    Cmd.CommandText = "[dbo].[Updatekintai]";

                    Cmd.Parameters.Add(new SqlParameter("対象年月", SqlDbType.Date));
                    Cmd.Parameters["対象年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Decimal));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("延長h", SqlDbType.Decimal));
                    Cmd.Parameters["延長h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("法休h", SqlDbType.Decimal));
                    Cmd.Parameters["法休h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所休h", SqlDbType.Decimal));
                    Cmd.Parameters["所休h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("総残h", SqlDbType.Decimal));
                    Cmd.Parameters["総残h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("六十超h", SqlDbType.Decimal));
                    Cmd.Parameters["六十超h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("深夜h", SqlDbType.Decimal));
                    Cmd.Parameters["深夜h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("遅刻回", SqlDbType.Decimal));
                    Cmd.Parameters["遅刻回"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("遅刻h", SqlDbType.Decimal));
                    Cmd.Parameters["遅刻h"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所定", SqlDbType.Decimal));
                    Cmd.Parameters["所定"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("法休", SqlDbType.Decimal));
                    Cmd.Parameters["法休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("所休", SqlDbType.Decimal));
                    Cmd.Parameters["所休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("有給", SqlDbType.Decimal));
                    Cmd.Parameters["有給"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特休", SqlDbType.Decimal));
                    Cmd.Parameters["特休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無特", SqlDbType.Decimal));
                    Cmd.Parameters["無特"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("振休", SqlDbType.Decimal));
                    Cmd.Parameters["振休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("公休", SqlDbType.Decimal));
                    Cmd.Parameters["公休"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("届欠", SqlDbType.Decimal));
                    Cmd.Parameters["届欠"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無届", SqlDbType.Decimal));
                    Cmd.Parameters["無届"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("回数1", SqlDbType.Decimal));
                    Cmd.Parameters["回数1"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("回数2", SqlDbType.Decimal));
                    Cmd.Parameters["回数2"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("特休理由", SqlDbType.VarChar));
                    Cmd.Parameters["特休理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("無特理由", SqlDbType.VarChar));
                    Cmd.Parameters["無特理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("休日超過理由", SqlDbType.VarChar));
                    Cmd.Parameters["休日超過理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("出勤超過理由", SqlDbType.VarChar));
                    Cmd.Parameters["出勤超過理由"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("コメント", SqlDbType.VarChar));
                    Cmd.Parameters["コメント"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("最終更新日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["最終更新日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("防止日数", SqlDbType.Decimal));
                    Cmd.Parameters["防止日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("休業日数", SqlDbType.Decimal));
                    Cmd.Parameters["休業日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("その他日数", SqlDbType.Decimal));
                    Cmd.Parameters["その他日数"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                    Cmd.Parameters["備考"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("途欠", SqlDbType.Decimal));
                    Cmd.Parameters["途欠"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD;
                    Cmd.Parameters["社員番号"].Value = Convert.ToDecimal(number.Text);
                    Cmd.Parameters["延長h"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["法休h"].Value = Convert.ToDecimal(houteikyuuH.Text);
                    Cmd.Parameters["所休h"].Value = Convert.ToDecimal(syoteikyuuH.Text);
                    Cmd.Parameters["総残h"].Value = Convert.ToDecimal(zangyouH.Text);
                    Cmd.Parameters["六十超h"].Value = Convert.ToDecimal(chouzanH.Text);
                    Cmd.Parameters["深夜h"].Value = Convert.ToDecimal(shinyaH.Text);
                    Cmd.Parameters["遅刻回"].Value = 0; // Convert.ToDecimal(chikoku.Text);
                    Cmd.Parameters["遅刻h"].Value = Convert.ToDecimal(chikokuH.Text);
                    Cmd.Parameters["所定"].Value = Convert.ToDecimal(syotei.Text);
                    Cmd.Parameters["法休"].Value = Convert.ToDecimal(houteikyuu.Text);
                    Cmd.Parameters["所休"].Value = Convert.ToDecimal(syoteikyuu.Text);
                    Cmd.Parameters["有給"].Value = Convert.ToDecimal(yuukyuu.Text);
                    Cmd.Parameters["特休"].Value = Convert.ToDecimal(tokkyuu.Text);
                    Cmd.Parameters["無特"].Value = Convert.ToDecimal(mutoku.Text);
                    Cmd.Parameters["振休"].Value = Convert.ToDecimal(hurikyuu.Text);
                    Cmd.Parameters["公休"].Value = Convert.ToDecimal(koukyuu.Text);
                    Cmd.Parameters["届欠"].Value = Convert.ToDecimal(todokede.Text);
                    Cmd.Parameters["無届"].Value = Convert.ToDecimal(mutodoke.Text);
                    Cmd.Parameters["回数1"].Value = Convert.ToDecimal(kaisuu1.Text);
                    Cmd.Parameters["回数2"].Value = Convert.ToDecimal(kaisuu2.Text);
                    Cmd.Parameters["特休理由"].Value = combotoku.SelectedItem.ToString();
                    Cmd.Parameters["無特理由"].Value = combomutoku.SelectedItem.ToString();
                    Cmd.Parameters["休日超過理由"].Value = RestReason.SelectedItem.ToString();
                    Cmd.Parameters["出勤超過理由"].Value = ooireason.SelectedItem.ToString();
                    Cmd.Parameters["コメント"].Value = textBox1.Text;
                    Cmd.Parameters["最終更新日時"].Value = DateTime.Now;

                    Cmd.Parameters["防止日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["休業日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["その他日数"].Value = Convert.ToDecimal(0);
                    Cmd.Parameters["備考"].Value = kyuugyouriyuu.SelectedItem.ToString();
                    Cmd.Parameters["途欠"].Value = Convert.ToDecimal(toketsu.Text);
                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }

        }

        /// <summary>
        /// 小数点以下の有無確認
        /// </summary>
        /// <param name="dValue"></param>
        /// <returns></returns>
        public static bool IsDecimal(string dValue)
        {
            if (Convert.ToDecimal(dValue) - Math.Floor(Convert.ToDecimal(dValue)) != 0)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 合計日数の更新
        /// </summary>
        /// <param name="e"></param>
        private void CheckTotalDays(CancelEventArgs e)
        {
            if (!e.Cancel)
            {
                //合計日数
                gengoukei = (Convert.ToDecimal(syotei.Text) + Convert.ToDecimal(houteikyuu.Text) + Convert.ToDecimal(syoteikyuu.Text)
                                    + Convert.ToDecimal(yuukyuu.Text) + Convert.ToDecimal(tokkyuu.Text) + Convert.ToDecimal(mutoku.Text) //+ Convert.ToDecimal(choukyuu.Text)
                                    + Convert.ToDecimal(hurikyuu.Text) + Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(todokede.Text) + Convert.ToDecimal(mutodoke.Text)).ToString();

                if (Convert.ToDecimal(gengoukei) > Convert.ToDecimal(hituyouDay))
                {
                    zanDays.Text = (Convert.ToDecimal(gengoukei) - Convert.ToDecimal(hituyouDay)).ToString() + "日多い　" + "[" + gengoukei + "/" + hituyouDay + "]";
                }
                else if (Convert.ToDecimal(gengoukei) < Convert.ToDecimal(hituyouDay))
                {
                    zanDays.Text = "残り" + (Convert.ToDecimal(hituyouDay) - Convert.ToDecimal(gengoukei)).ToString() + "日　" + "[" + gengoukei + "/" + hituyouDay + "]";
                }
                else
                {
                    zanDays.Text = "[" + gengoukei + "/" + hituyouDay + "]";
                }
            }
        }

        /// <summary>
        /// バリデータチェック
        /// </summary>
        /// <param name="cont"></param>
        /// <param name="e"></param>
        /// <param name="hOrD"></param>
        private void validCommon(Control cont, CancelEventArgs e, string hOrD)
        {
            if (number.Text == "" || number.Text == "-") return;

            //初回エラークリア
            this.errorProvider1.SetError(cont, "");
            ErrorMsg.Text = "";
            WarningMsg.Text = "";

            //全角は半角に
            String result = Strings.StrConv(cont.Text, VbStrConv.Narrow).Trim();

            //空白チェック
            if (result == "")
            {
                this.errorProvider1.SetError(cont, "入力してください");
                e.Cancel = true;

                ErrorMsg.Text += "入力してください" + nl;
                return;
            }

            //数値以外除去 
            Regex regex = new Regex(@"^[-.0-9]+$");
            if (!regex.IsMatch(result))
            {
                ErrorMsg.Text += "数字以外は入力しないでください" + nl;
                this.errorProvider1.SetError(cont, "数字以外は入力しないでください");
                e.Cancel = true;
            }

            //全角入力対応
            cont.Text = result;

            if (e.Cancel == true) return;

            //小数点エラー
            if (hOrD == "H")
            {
                if ((Convert.ToDecimal(result) * 10 % 5) > 0)
                {
                    ErrorMsg.Text += "少数点エラー" + nl;
                    this.errorProvider1.SetError(cont, "少数点エラー");
                    e.Cancel = true; 
                }

            }
            else if (hOrD == "D")
            {
                if (IsDecimal(result))
                {
                    ErrorMsg.Text += "少数点エラー" + nl;
                    this.errorProvider1.SetError(cont, "少数点エラー");
                    e.Cancel = true; return;
                }
            }
            else
            {
            }

            //個別エラー対応
            if (cont.Name == "syotei")
            {
                if (kubun.Text != "臨時社員" && kubun.Text != "パート")
                {
                    if (roudouD.Text == "３日")
                    {
                            //TODO とりあえず無
                    }
                    else
                    {
                        //土日祝祭日の日数で変化
                        if (Convert.ToDecimal(result) > syoteisuu)
                        {
                            ErrorMsg.Text += "所定出勤オーバーです" + Environment.NewLine;
                            ooireason.Enabled = true;
                        }
                    }
                }
            }
            else if (cont.Name == "houteikyuu")
            {
                if (Convert.ToInt32(result) > 0 & Convert.ToDecimal(houteikyuuH.Text) == 0)
                {
                    ErrorMsg.Text += "法定休の時間未入力" + nl;
                }
            }
            else if (cont.Name == "syoteikyuu")
            {
                if (Convert.ToInt32(result) > 10 & Convert.ToDecimal(syotei.Text) == 0)
                {
                    ErrorMsg.Text += "所定休出オーバーです" + nl;
                }

                if (Convert.ToInt32(result) > 0 & Convert.ToDecimal(syoteikyuuH.Text) == 0)
                {
                    ErrorMsg.Text += "所定休出の時間未入力" + nl;
                }
            }
            else if (cont.Name == "hurikyuu")
            {
                if (Convert.ToDecimal(result) > 10)
                {
                    ErrorMsg.Text += "振休オーバーです" + nl;
                }
            }
            else if (cont.Name == "koukyuu")
            {
                if (kubun.Text != "臨時社員" && kubun.Text != "パート") 
                {
                    //TODO とりあえず対象は正社員のみ
                    if (roudouD.Text == "３日")
                    {
                        //if (Convert.ToDecimal(result) > Convert.ToDecimal(sd.KyuuzitsuDay + 9))
                        if (Convert.ToDecimal(result) > kyuuzisuu + 9)
                        {
                            ErrorMsg.Text += "休日オーバーです" + Environment.NewLine;
                            RestReason.Enabled = true;
                        }
                    }
                    else if (roudouD.Text == "４日")
                    {
                        //if (Convert.ToDecimal(result) > Convert.ToDecimal(sd.KyuuzitsuDay + 4))
                        if (Convert.ToDecimal(result) > kyuuzisuu + 4)
                        {
                            ErrorMsg.Text += "休日オーバーです" + Environment.NewLine;
                            RestReason.Enabled = true;
                        }
                    }
                    else
                    {
                        //if (Convert.ToDecimal(result) > Convert.ToDecimal(sd.KyuuzitsuDay))
                        if (Convert.ToDecimal(result) > kyuuzisuu)
                        {
                            ErrorMsg.Text += "休日オーバーです" + Environment.NewLine;
                            RestReason.Enabled = true;
                        }
                    }

                }

            }
            else if (cont.Name == "kaisuu1")
            {
                if (Convert.ToInt32(result) > Convert.ToInt32(hituyouDay))
                {
                    ErrorMsg.Text += "回数1オーバーです" + nl;
                }
            }
            else if (cont.Name == "kaisuu2")
            {
                //TODO 先月回数と今月回数を
                if (td.StartYMD.ToString("MM") == "12")
                {
                    if (Convert.ToInt32(result) > 2)
                    {
                        ErrorMsg.Text += "12月分の正月手当上限は「２」までです。" + nl;
                    }
                }
                else if (td.StartYMD.ToString("MM") == "01")
                {
                    int i = 0;
                    if (label59.Text == "")
                    {
                        i = 0;
                    }
                    else
                    {
                        i = Convert.ToInt32(label59.Text);
                    }

                    if (Convert.ToInt32(result) + i > 3)
                    {
                        ErrorMsg.Text += "正月手当上限は、1月分と12月分合せて「３」までです。" + nl;
                        RestReason.Enabled = true;
                    }
                }


                ////TODO202012 
                //if (Convert.ToInt32(result) > 2)
                //{
                //    ErrorMsg.Text += "回数2オーバーです" + nl;
                //}

                //TODO202012 コメントアウト
                //if (Convert.ToInt32(result) > Convert.ToInt32(hituyouDay))
                //{
                //    ErrorMsg.Text += "回数2オーバーです" + nl;
                //}
            }

            //個別エラー(合計日数)対応
            if (cont.Name == "syotei" | cont.Name == "houteikyuu" | cont.Name == "syoteikyuu" | cont.Name == "yuukyuu" |
                cont.Name == "tokkyuu" | cont.Name == "hurikyuu" | cont.Name == "koukyuu" | cont.Name == "choukyuu" |
                cont.Name == "todokede" | cont.Name == "mutodoke" | cont.Name == "mutoku" |
                cont.Name == "boushi" | cont.Name == "kyuugyou" //| cont.Name == "sonota"
                )
            {
                CheckTotalDays(e);
            }

            //実労働時間対応
            if (cont.Name == "syotei" | cont.Name == "zangyouH" | cont.Name == "syotei" | cont.Name == "entyou" |
                cont.Name == "houteikyuuH" | cont.Name == "syoteikyuuH" | cont.Name == "chikokuH")
            {
                //実労働時間の表示
                decimal dec = Convert.ToDecimal(entyou.Text) + Convert.ToDecimal(zangyouH.Text) + Convert.ToDecimal(syotei.Text) * Convert.ToDecimal(kinmuh.Text.Replace(" 時間", "")) + Convert.ToDecimal(houteikyuuH.Text) + Convert.ToDecimal(syoteikyuuH.Text) - Convert.ToDecimal(chikokuH.Text);
                zitsuroudouH.Text = Convert.ToString(dec);
                zitsuroudouH.Text += "時間";

                moneycolor();
            }

            decimal syoteirou = Convert.ToDecimal(entyou.Text) + Convert.ToDecimal(syotei.Text) * Convert.ToDecimal(kinmuh.Text.Replace(" 時間", "")) - Convert.ToDecimal(chikokuH.Text);

            //固定値→データベース参照へ変更 201803
            //if (syoteirou > sd.RoudouH)
            if (syoteirou > roudzikan)
            {
                WarningMsg.Text += "【警告】" + "所定労働オーバー" + Convert.ToString(syoteirou) + "労基法違反" + nl;
            }

            //係長以上は休出/残業/遅刻は警告を表示
            //if (Convert.ToInt16(yakusyoku.Text) <= 135)

            if (Convert.ToInt16(yakusyoku) <= 135)
            {
                if (cont.Name == "houteikyuuH" | cont.Name == "houteikyuu" | cont.Name == "syoteikyuuH" |
                     cont.Name == "syoteikyuu" | cont.Name == "zangyouH" | cont.Name == "shinyaH")
                {
                    if (Convert.ToDecimal(cont.Text) > 0)
                    {
                        Dictionary<string, string> dic = new Dictionary<string, string>()
                        {
                            {"entyou", "延長"},
                            {"houteikyuuH", "法定休H"},
                            {"houteikyuu", "法定休出"},
                            {"syoteikyuuH", "所定休H"},
                            {"syoteikyuu", "所定休出"},
                            {"zangyouH", "残業"}, 
                            {"shinyaH", "深夜"}
                        };

                        WarningMsg.Text += "【警告】" + "係長職以上に" + dic[cont.Name] + " OK?" + nl;
                    }
                }
            }


            //その他(残業60超対応)
            if ((cont.Name == "zangyouH" || cont.Name == "syoteikyuuH") & Convert.ToDecimal(syoteikyuuH.Text) + Convert.ToDecimal(zangyouH.Text) > 60)
            {
                chouzanH.Text = (Convert.ToDecimal(syoteikyuuH.Text) + Convert.ToDecimal(zangyouH.Text) - 60).ToString();
            }
            else if ((cont.Name == "zangyouH" || cont.Name == "syoteikyuuH") & Convert.ToDecimal(syoteikyuuH.Text) + Convert.ToDecimal(zangyouH.Text) <= 60)
            {
                chouzanH.Text = "0.0";
            }

            //個別警告対応
            if (cont.Name == "entyou" | cont.Name == "houteikyuuH" | cont.Name == "syoteikyuuH" |
                cont.Name == "zangyouH" | cont.Name == "shinyaH")
            {
                if (Convert.ToDecimal(result) >= 100)
                {
                    Dictionary<string, string> dic = new Dictionary<string, string>()
                    {
                        {"entyou", "延長"},
                        {"houteikyuuH", "法定休出"},
                        {"syoteikyuuH", "所定休出"},
                        {"zangyouH", "残業"}, 
                        {"shinyaH", "深夜"}
                    };

                    WarningMsg.Text += "【警告】" + dic[cont.Name] + " 100時間以上OK?" + nl;
                }
            }

            //ロワジール対応
            if (genba.Text == "ロワジールホテル那覇" & syokusyu.Text.Substring(0, 2) == "客室" & (cont.Name == "syotei" || cont.Name == "houteikyuu" || cont.Name == "syoteikyuu") & yakin.Text != "")
            {
                kaisuu1.Text = (Convert.ToInt16(syotei.Text) + Convert.ToInt16(houteikyuu.Text) + Convert.ToInt16(syoteikyuu.Text)).ToString();
            }

            //特休/無特　使用理由選択
            if (cont.Name == "tokkyuu")
            {
                if (Convert.ToDecimal(result) > 0 && combotoku.SelectedItem.ToString() == "")
                {
                    ErrorMsg.Text += "特休理由を選択してください" + nl;
                }

                if (Convert.ToDecimal(result) == 0 && combotoku.SelectedItem.ToString() != "")
                {
                    ErrorMsg.Text += "特休理由を選択してるけど入力無" + nl;
                }
            }
            
            if (cont.Name == "mutoku")
            //else if (cont.Name == "sonota")
            {
                if (Convert.ToDecimal(result) > 0 && combomutoku.SelectedItem.ToString() == "")
                {
                    ErrorMsg.Text += "無休理由を選択してください" + nl;
                }

                if (Convert.ToDecimal(result) == 0 && combomutoku.SelectedItem.ToString() != "")
                {
                    ErrorMsg.Text += "無休理由を選択してるけど入力無" + nl;
                }
            }

            if (kubun.Text != "臨時社員" && kubun.Text != "パート")
            {
                if (roudouD.Text == "３日")
                {
                    if (Convert.ToDecimal(syotei.Text) > sd.SyoteiDay + 9)
                    {
                        ooireason.Enabled = true;
                    }

                    if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > sd.KyuuzitsuDay + 9)
                    {
                        RestReason.Enabled = true;
                    }
                }
                if (roudouD.Text == "４日")
                {
                    if (Convert.ToDecimal(syotei.Text) > sd.SyoteiDay + 4)
                    {
                        ooireason.Enabled = true;
                    }

                    if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > sd.KyuuzitsuDay + 4)
                    {
                        RestReason.Enabled = true;
                    }
                }
                else
                {
                    if (Convert.ToDecimal(syotei.Text) > sd.SyoteiDay)
                    {
                        ooireason.Enabled = true;
                    }

                    if (Convert.ToDecimal(koukyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) > sd.KyuuzitsuDay)
                    {
                        RestReason.Enabled = true;
                    }
                }

            }


            //月給者のみ 月途中入社退社欠勤処理
            if (kubun.Text == "月給者")
            {
                
                if (Convert.ToDateTime(nyuusya.Text) > td.StartYMD)
                {
                    //途中入社
                    toketsu.Text = (syoteisuu - (Convert.ToDecimal(syotei.Text) + Convert.ToDecimal(yuukyuu.Text) + Convert.ToDecimal(tokkyuu.Text) + Convert.ToDecimal(mutoku.Text) + Convert.ToDecimal(hurikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(todokede.Text) + Convert.ToDecimal(mutodoke.Text))).ToString();
                }

                if (taisyoku.Text != "")
                {
                    if (Convert.ToDateTime(taisyoku.Text) < td.EndYMD)
                    {
                        //途中退社
                        toketsu.Text = (syoteisuu - (Convert.ToDecimal(syotei.Text) + Convert.ToDecimal(yuukyuu.Text) + Convert.ToDecimal(tokkyuu.Text) + Convert.ToDecimal(mutoku.Text) + Convert.ToDecimal(hurikyuu.Text) + Convert.ToDecimal(houteikyuu.Text) + Convert.ToDecimal(syoteikyuu.Text) + Convert.ToDecimal(todokede.Text) + Convert.ToDecimal(mutodoke.Text))).ToString();
                    }
                }
            }


            return;
        }

        #region バリデータ
        private void entyou_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void houteikyuuH_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void syoteikyuuH_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void zangyouH_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void chouzanH_Validating(object sender, CancelEventArgs e)
        {
        }

        private void shinyaH_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void chikoku_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void chikokuH_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void syotei_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void houteikyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void syoteikyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void yuukyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void tokkyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void mutoku_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void hurikyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void koukyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void choukyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void todokede_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void mutodoke_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "H");
        }

        private void kaisuu1_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void kaisuu2_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }
        #endregion

        private void DGView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            entyou.Select();
        }

        private void DGView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                entyou.Select();
        }

        private void SetCheck()
        {
            string tname = Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value);
            int errct = Convert.ToInt16(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value);
            int nonct = Convert.ToInt16(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value);


            if (errct + nonct > 0)
            {
                if (tabControl1.SelectedIndex == 1 || tabControl1.SelectedIndex == 2)
                {
                    dataGridView3.DataSource = "";
                    MessageBox.Show(tname + "様担当分" + nl + "エラー数と未提出数がゼロにならないと選択できません。");
                }

                //ボタン処理　Excel出力とチェック完了ボタンを非表示
                button3.Visible = false;
                button4.Visible = false;
                label65.Text = "";
            }
            else
            {
                if (tabControl1.SelectedIndex == 1)
                {
                    //チェックリスト
                    GetCheckData();
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    //修正リスト
                    GetReData();
                }

                //ボタン処理　Excel出力とチェック完了ボタンを表示
                button3.Visible = true;
                button4.Visible = true;
                label65.Text = tname + "様担当分";
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //column
            //0:担当部署
            //1:担当者
            //2:対象者数　削除
            //3:当月入社数 削除
            //4:当月退社数 削除
            //5:未発令数　削除
            //6:提出者数
            //7:エラー数
            //8:警告数
            //9:未提出数

            //row
            //0:技術
            //1:客室
            //2:現業
            //3:事務

            //完了ボタン、Excel出力ボタン等の処理
            SetCheck();

            int result;

            //数値以外はスルー
            if (int.TryParse(dataGridView1.CurrentCell.Value.ToString(), out result))
            {
                //ゼロはスルー
                if (result > 0)
                {
                    dgvListCol = dataGridView1.CurrentCell.ColumnIndex;
                    dgvListRow = dataGridView1.CurrentCell.RowIndex;

                    //詳細情報取得
                    GetListDetail(dgvListCol, dgvListRow, Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value), Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value));

                    //選択行の地区カラムを取得
                    string tiku = dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString();

                    //タイトル表示
                    //counter.Text = dataGridView1.CurrentCell.OwningColumn.DataPropertyName + "　【" + tiku + "】";
                }
            }

            //ソート不可対応
            foreach (DataGridViewColumn c in DGView.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;
        }

        private void GetListDetail(int columnct, int rowct, string flg, string kubun)
        {
            //string filtStr = "担当管理 = '" + flg + "'";

            string filtStr = "";
            if (Program.loginname == "緑間　拓也" || Program.loginname == "親泊　美和子")
            {
                filtStr = "(担当管理 = '緑間　拓也' and 担当区分 = '" + kubun + "') or (担当区分 = '" + kubun + "' and 担当事務 = '03_施設')";
            }
            else if (Program.loginname == "吉里　勇二郎" || Program.loginname == "島　辰之")
            {
                filtStr = "(担当管理 = '吉里　勇二郎' and 担当区分 = '" + kubun + "') or (担当区分 = '" + kubun + "' and 担当事務 = '03_警備')";
            }
            else if (Program.loginbusyo == "01_現業" && kubun != "01_現業")
            {
                filtStr = "担当区分 = '" + kubun + "' and 担当事務 = '01_現業'";
            }
            else if (Program.loginname == "太田　朋宏")
            {
                //filtStr = "担当管理 = '" + flg + "' and 担当区分 = '" + kubun + "'";
                filtStr = "担当区分 = '" + kubun + "' and 担当事務 = '09_管理事務'";
            }
            //else if (Program.loginbusyo == "14_宮古島")
            //{
            //    filtStr = "担当管理 = '" + flg + "' and 担当区分 = '" + kubun + "'";
            //}
            else
            {
                filtStr = "担当管理 = '" + flg + "'";
            }

            //string colname = "";
            switch (columnct)
            {
                //case 2: //対象者数
                //    filtStr += "and 在籍区分 <> '0'";
                //    colname = "対象数";
                //    break;
                case 2: //対象者数
                    filtStr += "and 登録フラグ = '1'";
                    //colname = "提出数";
                    break;
                case 3: //エラー数
                    filtStr += "and 登録フラグ = '1'";
                    //colname = "エラー";
                    break;
                case 4: //警告数
                    filtStr += "and 登録フラグ = '1'";
                    //colname = "警告";
                    break;
                case 5: //未提出数
                    filtStr += "and 登録フラグ = '0' and 在籍区分 <> '0'";
                    //colname = "未提出";
                    break;
            }

            DataRow[] drYae;
            if (Program.loginbusyo == "12_八重山" || Program.loginbusyo == "11_北部" || Program.loginbusyo == "14_宮古島" || Program.loginbusyo == "15_久米島")
            //if (flg == "石川　尚吾" || flg == "篠原　崇" || flg == "大城　森一" || flg == "中真　心" || flg == "宮城　一禎" || flg == "川満　勇人")
            {
                drYae = dt.Select(filtStr, "現場CD,組織CD");
            }
            else if (flg == "喜屋武　大祐" || flg == "高江洲　華子" || flg == "仲里　かおり" || flg == "太田　朋宏" || flg == "新垣　聖悟")
            {
                drYae = dt.Select(filtStr, "組織CD,現場CD");
            }
            else if (flg == "")
            {
                drYae = dt.Select(filtStr, "組織CD,現場CD");
            }
            else
            {
                drYae = dt.Select(filtStr, "現場CD");
            }

            DataTable Disp = new DataTable();
            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("漢字氏名", typeof(string));
            Disp.Columns.Add("カナ名", typeof(string));
            Disp.Columns.Add("組織名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));
            Disp.Columns.Add("状況", typeof(string));

            if (columnct == 3)
            {
                foreach (DataRow row in drYae)
                {
                    string[] st = co.ErrorCheck(row, flg);
                    if (st[0].Length > 0)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号"] = row["社員番号"];
                        nr["漢字氏名"] = row["漢字氏名"];
                        nr["カナ名"] = row["カナ氏名"];
                        nr["組織名"] = row["組織名"];
                        nr["現場名"] = row["現場名"];
                        nr["状況"] = st[0].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }
            else if (columnct == 4)
            {
                foreach (DataRow row in drYae)
                {
                    string[] st = co.ErrorCheck(row, flg);
                    if (st[4].Length > 0)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["社員番号"] = row["社員番号"];
                        nr["漢字氏名"] = row["漢字氏名"];
                        nr["カナ名"] = row["カナ氏名"];
                        nr["組織名"] = row["組織名"];
                        nr["現場名"] = row["現場名"];
                        nr["状況"] = st[4].ToString();
                        Disp.Rows.Add(nr);
                    }
                }
            }
            else if (columnct == 5)
            {
                foreach (DataRow row in drYae)
                {
                    DataRow nr = Disp.NewRow();
                    nr["社員番号"] = row["社員番号"];
                    nr["漢字氏名"] = row["漢字氏名"];
                    nr["カナ名"] = row["カナ氏名"];
                    nr["組織名"] = row["組織名"];
                    nr["現場名"] = row["現場名"];
                    nr["状況"] = row["登録フラグ"].ToString() == "1" ? "登録済" : "未";
                    Disp.Rows.Add(nr);
                }

            }
            else
            {
                foreach (DataRow row in drYae)
                {
                    string[] st = co.ErrorCheck(row, flg);
                    //if (st[4].Length > 0)
                    //{
                    DataRow nr = Disp.NewRow();
                    nr["社員番号"] = row["社員番号"];
                    nr["漢字氏名"] = row["漢字氏名"];
                    nr["カナ名"] = row["カナ氏名"];
                    nr["組織名"] = row["組織名"];
                    nr["現場名"] = row["現場名"];
                    nr["状況"] = st[4].ToString();
                    Disp.Rows.Add(nr);
                    //}
                }
            }

            DGView.DataSource = Disp;
        }

        private void DGView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //TODO:右クリック
            if (e.Button == MouseButtons.Right)
            {
               // MessageBox.Show(e.ColumnIndex.ToString() + e.RowIndex.ToString());

                // ヘッダ以外
                if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
                {

                //    // 右クリックされたセル
                //    DataGridViewCell cell = dgv[e.ColumnIndex, e.RowIndex];
                //    // セルの選択状態を反転
                //    cell.Selected = !cell.Selected;
                }
            }
        }

        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                GetSearchData();
                Com.InHistory("33_勤怠入力_RE", SearchBox.Text, "");
            }
        }

        private void GetSearchData()
        {
            string res = SearchBox.Text.Trim();
            string[] ar = res.Replace("　", " ").Split(' ');

            string result = "";
            int count = 0;

            result = " where 担当管理 like '%" + "%'";
            

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    //if (count == 0)
                    //    result = " where キー like '%" + s + "%'";
                    //else
                        result += " and キー like '%" + s + "%'";

                    count++;
                }
            }

            string strCon = Common.constr;
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter Adapter;
            DataTable dt = new DataTable();

            using (Cn = new SqlConnection(strCon))
            {
                Cmd = Cn.CreateCommand();
                Cmd.CommandText = "select 社員番号, 漢字氏名, カナ氏名, 組織名, 現場名, 担当管理 from dbo.search" + result + " order by 地区名, 組織名, 現場名, 社員番号";
                Adapter = new SqlDataAdapter(Cmd);
                Adapter.Fill(dt);
            }

            DGView.DataSource = dt;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (number.Text == "" || number.Text == "-") return;

            DialogResult result = MessageBox.Show("クリアしてよろしいですか？",
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
                    Cmd.CommandText = "delete from dbo.勤怠データ WHERE 社員番号 = '" + number.Text + "'";
                    using (dr = Cmd.ExecuteReader())
                    {
                        //TODO
                    }
                }
            }

            GetData();
            DisplayGridView();
        }


#region テキストボックスフォーカス取得時の全選択対応
        private bool syotei_flg;
        private bool entyou_flg;
        private bool houteikyuuH_flg;
        private bool syoteikyuuH_flg;
        private bool zangyouH_flg;
        private bool chouzanH_flg;
        private bool shinyaH_flg;
        private bool chikoku_flg;
        private bool chikokuH_flg;
        private bool houteikyuu_flg;
        private bool syoteikyuu_flg;
        private bool yuukyuu_flg;
        private bool tokkyuu_flg;
        private bool mutoku_flg;
        private bool hurikyuu_flg;
        private bool koukyuu_flg;
        private bool choukyuu_flg;
        private bool todokede_flg;
        private bool mutodoke_flg;
        private bool kaisuu1_flg;
        private bool kaisuu2_flg;

        private bool boushi_flg;
        private bool kyuugyou_flg;
        private bool sonota_flg;


        private void syotei_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            syotei_flg = true;
        }

        private void syotei_MouseDown(object sender, MouseEventArgs e)
        {
            if (syotei_flg)
            {
                selectAlltbox((TextBox)sender);
                syotei_flg = false;
            }
        }

        private void selectAlltbox(TextBox cont)
        {
            cont.SelectAll();
        }



        private void entyou_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            entyou_flg = true;
        }

        private void houteikyuuH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            houteikyuuH_flg = true;
        }

        private void syoteikyuuH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            syoteikyuuH_flg = true;
        }

        private void zangyouH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            zangyouH_flg = true;
        }

        private void chouzanH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            chouzanH_flg = true;
        }

        private void shinyaH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            shinyaH_flg = true;
        }

        private void chikoku_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            chikoku_flg = true;
        }

        private void chikokuH_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            chikokuH_flg = true;
        }

        private void houteikyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            houteikyuu_flg = true;
        }

        private void syoteikyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            syoteikyuu_flg = true;
        }

        private void yuukyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            yuukyuu_flg = true;
        }

        private void tokkyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            tokkyuu_flg = true;
        }

        private void mutoku_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            mutoku_flg = true;
        }

        private void hurikyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            hurikyuu_flg = true;
        }

        private void koukyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            koukyuu_flg = true;
        }

        private void choukyuu_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            choukyuu_flg = true;
        }

        private void todokede_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            todokede_flg = true;
        }

        private void mutodoke_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            mutodoke_flg = true;
        }

        private void kaisuu1_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            kaisuu1_flg = true;
        }

        private void kaisuu2_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            kaisuu2_flg = true;
        }



        private void entyou_MouseDown(object sender, MouseEventArgs e)
        {
            if (entyou_flg)
            {
                selectAlltbox((TextBox)sender);
                entyou_flg = false;
            }
        }

        private void houteikyuuH_MouseDown(object sender, MouseEventArgs e)
        {
            if (houteikyuuH_flg)
            {
                selectAlltbox((TextBox)sender);
                houteikyuuH_flg = false;
            }
        }

        private void syoteikyuuH_MouseDown(object sender, MouseEventArgs e)
        {
            if (syoteikyuuH_flg)
            {
                selectAlltbox((TextBox)sender);
                syoteikyuuH_flg = false;
            }
        }

        private void zangyouH_MouseDown(object sender, MouseEventArgs e)
        {
            if (zangyouH_flg)
            {
                selectAlltbox((TextBox)sender);
                zangyouH_flg = false;
            }
        }

        private void chouzanH_MouseDown(object sender, MouseEventArgs e)
        {
            if (chouzanH_flg)
            {
                selectAlltbox((TextBox)sender);
                chouzanH_flg = false;
            }
        }

        private void shinyaH_MouseDown(object sender, MouseEventArgs e)
        {
            if (shinyaH_flg)
            {
                selectAlltbox((TextBox)sender);
                shinyaH_flg = false;
            }
        }

        private void chikoku_MouseDown(object sender, MouseEventArgs e)
        {
            if (chikoku_flg)
            {
                selectAlltbox((TextBox)sender);
                chikoku_flg = false;
            }
        }

        private void chikokuH_MouseDown(object sender, MouseEventArgs e)
        {
            if (chikokuH_flg)
            {
                selectAlltbox((TextBox)sender);
                chikokuH_flg = false;
            }
        }

        private void houteikyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (houteikyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                houteikyuu_flg = false;
            }
        }

        private void syoteikyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (syoteikyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                syoteikyuu_flg = false;
            }
        }

        private void yuukyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (yuukyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                yuukyuu_flg = false;
            }
        }

        private void tokkyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (tokkyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                tokkyuu_flg = false;
            }
        }

        private void mutoku_MouseDown(object sender, MouseEventArgs e)
        {
            if (mutoku_flg)
            {
                selectAlltbox((TextBox)sender);
                mutoku_flg = false;
            }
        }

        private void hurikyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (hurikyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                hurikyuu_flg = false;
            }
        }

        private void koukyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (koukyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                koukyuu_flg = false;
            }
        }

        private void choukyuu_MouseDown(object sender, MouseEventArgs e)
        {
            if (choukyuu_flg)
            {
                selectAlltbox((TextBox)sender);
                choukyuu_flg = false;
            }
        }

        private void todokede_MouseDown(object sender, MouseEventArgs e)
        {
            if (todokede_flg)
            {
                selectAlltbox((TextBox)sender);
                todokede_flg = false;
            }
        }

        private void mutodoke_MouseDown(object sender, MouseEventArgs e)
        {
            if (mutodoke_flg)
            {
                selectAlltbox((TextBox)sender);
                mutodoke_flg = false;
            }
        }

        private void kaisuu1_MouseDown(object sender, MouseEventArgs e)
        {
            if (kaisuu1_flg)
            {
                selectAlltbox((TextBox)sender);
                kaisuu1_flg = false;
            }
        }

        private void kaisuu2_MouseDown(object sender, MouseEventArgs e)
        {
            if (kaisuu2_flg)
            {
                selectAlltbox((TextBox)sender);
                kaisuu2_flg = false;
            }
        }
#endregion

        private void button3_Click(object sender, EventArgs e)
        {
            if (number.Text == "" || number.Text == "-") return;

            string[] sArray = new string[] { this.name.Text, this.number.Text, this.syokusyu.Text, this.genba.Text, this.kubun.Text, this.nyuusya.Text, this.taisyoku.Text, this.kubuncode }; 

            ExData ex = new ExData(sArray);
            //ex.ShowDialog();
            ex.Show();
        }

        private DataTable GetKakoData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "過去勤怠";

                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                        Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("処理年", SqlDbType.VarChar));
                        Cmd.Parameters["処理年"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("処理月", SqlDbType.VarChar));
                        Cmd.Parameters["処理月"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("基準日", SqlDbType.VarChar));
                        Cmd.Parameters["基準日"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("給与支給区分", SqlDbType.VarChar));
                        Cmd.Parameters["給与支給区分"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("勤怠対象開始年月日", SqlDbType.VarChar));
                        Cmd.Parameters["勤怠対象開始年月日"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["社員番号"].Value = number.Text;
                        Cmd.Parameters["給与支給区分"].Value = kubuncode;
                        Cmd.Parameters["処理年"].Value = td.StartYMD.Year.ToString();
                        Cmd.Parameters["処理月"].Value = String.Format("{0:00}", td.StartYMD.Month);
                        Cmd.Parameters["基準日"].Value = td.StartYMD.ToString("yyyy/MM/dd");
                        Cmd.Parameters["勤怠対象開始年月日"].Value = td.StartYMD.ToString("yyyy/MM/dd");

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);

                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
    }

        private void kDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Right)
            {
                    //this.SelectNextControl(this.ActiveControl, !e.Shift, true, true, true);
                    this.SelectNextControl(this.ActiveControl, true, true, true, true);
            }

            if (e.KeyCode == Keys.Left)
            {
                if (!e.Control)
                {
                    this.SelectNextControl(this.ActiveControl, false, true, true, true);
                    //e.Handled = true;
                }
            }
        }

        private void entyou_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void houteikyuuH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void syoteikyuuH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void zangyouH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void chouzanH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void shinyaH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void chikoku_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void chikokuH_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void syotei_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void houteikyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void syoteikyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void yuukyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void tokkyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void mutoku_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void hurikyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void koukyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void choukyuu_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void todokede_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void mutodoke_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void kaisuu1_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void kaisuu2_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void radioButton1_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void radioButton2_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void button1_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            moneycolor();
        }

        private void moneycolor()
        {
            //実労働時間の表示
            decimal dec = Convert.ToDecimal(entyou.Text) + Convert.ToDecimal(zangyouH.Text) + Convert.ToDecimal(syotei.Text) * Convert.ToDecimal(kinmuh.Text.Replace(" 時間", "")) + Convert.ToDecimal(houteikyuuH.Text) + Convert.ToDecimal(syoteikyuuH.Text) - Convert.ToDecimal(chikokuH.Text);
            zitsuroudouH.Text = Convert.ToString(dec);
            zitsuroudouH.Text += "時間";
        }

        private void Nonbeep(KeyPressEventArgs e)
        {
            //イベントを処理済にする
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
            }
        }

        private void entyou_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void houteikyuuH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void syoteikyuuH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void zangyouH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void chouzanH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void shinyaH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void syotei_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void houteikyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void syoteikyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void yuukyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void tokkyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void mutoku_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void hurikyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void koukyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void choukyuu_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void todokede_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void mutodoke_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void kaisuu1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void kaisuu2_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void chikoku_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void chikokuH_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // 現在時を取得
            DateTime datetime_now = DateTime.Now;
            DateTime datetime_set = new DateTime(2013, 11, 06, 17, 30, 0); //年, 月, 日, 時間, 分, 秒
            int time = 0;

            //残り時間を秒に変換
            time = (((datetime_set.Day - datetime_now.Day) * 3600 * 24) + ((datetime_set.Hour - datetime_now.Hour) * 3600) + ((datetime_set.Minute - datetime_now.Minute) * 60) + (datetime_set.Second - datetime_now.Second));

            //残り時間を表示
            //label63.Text = "提出期限まで 残り " + time / 3600 + "時間 " + (time % 3600) / 60 + "分 " + (time % 3600) % 60 + "秒";

            if (datetime_now.ToLongTimeString() == datetime_set.ToLongTimeString())
            {
                //label63.Text = "提出期限は過ぎました。";
                //タイマー停止
                timer1.Stop();
            }
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            InsertSK();
        }

        private void InsertSK()
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
                    Cmd.CommandText = "[dbo].[InsertSK]";

                    Cmd.Parameters.Add(new SqlParameter("対象年月", SqlDbType.Date));
                    Cmd.Parameters["対象年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Decimal));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名", SqlDbType.Decimal));
                    Cmd.Parameters["氏名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("地区", SqlDbType.Decimal));
                    Cmd.Parameters["地区"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("地区コード", SqlDbType.Decimal));
                    Cmd.Parameters["地区コード"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("組織名", SqlDbType.Decimal));
                    Cmd.Parameters["組織名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("組織コード", SqlDbType.Decimal));
                    Cmd.Parameters["組織コード"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("現場名", SqlDbType.Decimal));
                    Cmd.Parameters["現場名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("現場コード", SqlDbType.Decimal));
                    Cmd.Parameters["現場コード"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("日勤時間", SqlDbType.Decimal));
                    Cmd.Parameters["日勤時間"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("残業時間", SqlDbType.Decimal));
                    Cmd.Parameters["残業時間"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("深夜時間", SqlDbType.Decimal));
                    Cmd.Parameters["深夜時間"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                    Cmd.Parameters["コメント"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("最終更新日時", SqlDbType.SmallDateTime));
                    Cmd.Parameters["最終更新日時"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象年月"].Value = td.StartYMD;
                    Cmd.Parameters["社員番号"].Value = Convert.ToDecimal(number.Text);
                    Cmd.Parameters["氏名"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["地区"].Value = Convert.ToDecimal(houteikyuuH.Text);
                    Cmd.Parameters["地区コード"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["組織名"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["組織コード"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["現場名"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["現場コード"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["日勤時間"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["残業時間"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["深夜時間"].Value = Convert.ToDecimal(entyou.Text);
                    Cmd.Parameters["備考"].Value = textBox1.Text;
                    Cmd.Parameters["最終更新日時"].Value = DateTime.Now;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void KintaiInfo(string str, string name)
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

                    if (str == "入力")
                    { 
                        Cmd.CommandText = "dbo.021_勤怠入力完了更新";
                    }
                    else
                    {
                        Cmd.CommandText = "dbo.022_勤怠チェック完了更新";
                    }

                    Cmd.Parameters.Add(new SqlParameter("処理年月", SqlDbType.VarChar));
                    Cmd.Parameters["処理年月"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("担当名", SqlDbType.VarChar));
                    Cmd.Parameters["担当名"].Direction = ParameterDirection.Input;

                    if (str == "入力")
                    {
                        Cmd.Parameters.Add(new SqlParameter("入力完了日時", SqlDbType.SmallDateTime));
                        Cmd.Parameters["入力完了日時"].Direction = ParameterDirection.Input;
                    }
                    else
                    {
                        Cmd.Parameters.Add(new SqlParameter("チェック完了日時", SqlDbType.SmallDateTime));
                        Cmd.Parameters["チェック完了日時"].Direction = ParameterDirection.Input;
                    }

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["処理年月"].Value = td.StartYMD.AddMonths(1).ToString("yyyyMM");
                    Cmd.Parameters["担当名"].Value = name;

                    if (str == "入力")
                    {
                        Cmd.Parameters["入力完了日時"].Value = DateTime.Now;
                    }
                    else
                    {
                        Cmd.Parameters["チェック完了日時"].Value = DateTime.Now;
                    }

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void GetKintaiInfo()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            kintaiinfodt.Clear();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select 担当名, 入力完了日時 as 入力完了, チェック完了日時 as チェック完了 from dbo.勤怠状況管理 order by チェック完了日時 desc, 入力完了日時 desc";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(kintaiinfodt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView2.DataSource = kintaiinfodt;

            //TIPS データグリッドビューの日付カラムのフォーマットを指定する
            dataGridView2.Columns["入力完了"].DefaultCellStyle.Format = "MM/dd HH:mm";
            dataGridView2.Columns["チェック完了"].DefaultCellStyle.Format = "MM/dd HH:mm";

            //担当分が終わってたら、ボタンを変更する
            foreach (DataRow dr in kintaiinfodt.Rows)
            {
                if (dr["担当名"].ToString() == Program.loginname)
                {
                    if (dr["チェック完了"].ToString() != "")
                    {
                        button3.Text = "チェック完了済";
                        button3.Enabled = false;
                        button3.BackColor = Color.Gray;

                        //TODO
                        
                    }
                }
            }
        }

        private void button3_Click_3(object sender, EventArgs e)
        {
            string tname = Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value);
            KintaiInfo("チェック", tname);
            GetKintaiInfo();
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            string tname = Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value);
            int errct = Convert.ToInt16(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value);
            int nonct = Convert.ToInt16(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value);


            //チェックリスト、修正タブを選択できないようにする
            if (e.TabPageIndex == 1 || e.TabPageIndex == 2)
            {
                if (errct + nonct > 0)
                {
                    MessageBox.Show(tname + "様担当分" + nl + "エラー数と未提出数がゼロにならないと選択できません。");
                    e.Cancel = true;
                }
                else
                {
                    if (e.TabPageIndex == 1)
                    {
                        //チェックリスト
                        GetCheckData();
                    }
                    else
                    {
                        //修正リスト
                        GetReData();
                    }
                }
            }
        }

        private void GetReData()
        {
            dataGridView4.DataSource = "";

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable redt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[勤怠修正リスト]";

                        Cmd.Parameters.Add(new SqlParameter("ym", SqlDbType.VarChar));
                        Cmd.Parameters["ym"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("name", SqlDbType.VarChar));
                        Cmd.Parameters["name"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["name"].Value = Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value);

                        Cmd.Parameters["ym"].Value = td.StartYMD.AddMonths(1).ToString("yyyyMM"); ;

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(redt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView4.DataSource = redt;
        }

        private void GetCheckData()
        {
            dataGridView3.DataSource = "";
            ckdt.Clear();

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        string name = Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value);

                        Cmd.CommandType = CommandType.StoredProcedure;

                        //並び順特化
                        if(name == "棚原様")
                        { 
                            Cmd.CommandText = "[dbo].[勤怠チェック_棚原様専用]";
                        }
                         else if (name == "石川　尚吾" || name == "太田　朋宏" || name == "喜屋武　大祐" || name == "篠原　崇" || name == "大城　森一" || name == "中真　心")
                        {
                            Cmd.CommandText = "[dbo].[勤怠チェック_石川様喜屋武専用]";
                        }
                        else
                        {
                            Cmd.CommandText = "[dbo].[勤怠チェック]";
                        }

                        Cmd.Parameters.Add(new SqlParameter("name", SqlDbType.VarChar));
                        Cmd.Parameters["name"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["name"].Value = name;

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(ckdt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            dataGridView3.DataSource = ckdt;

            if (ckdt.Rows.Count == 0) return;

            //dataGridView1.Rows[12].DefaultCellStyle.Format = "0\'%\'";
            //dataGridView1.Rows[13].DefaultCellStyle.Format = "0\'%\'";
            //dataGridView1.Rows[14].DefaultCellStyle.Format = "0\'%\'";


            for (int i = 0; i < ckdt.Columns.Count; i++)
            {
                //項目名以外は右寄せ表示
                if (i < 3)
                {
                    dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else if (i < 11)
                {
                    dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[i].Width = 45;
                }
                else
                {
                    dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[i].Width = 35;
                }

                //ヘッダーの中央表示
                dataGridView3.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            decimal val = 0;
            //セルの行を確認
            if (e.Value != null && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val == 0)
                {
                    //e.CellStyle.ForeColor = Color.Gray;
                    e.Value = null;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetCheckData();

            if (ckdt.Rows.Count == 0)
            {
                MessageBox.Show("おかしいっす。");
                return;
            }

            //ボタン無効化・カーソル変更
            button4.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string fileName = "";
            fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\checklist.xlsx";

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

            //TODO 表示名
            m_MySheet.Cells[1, 1] = "【" + Convert.ToString(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value) + "様担当分】" + "勤怠チェックリスト";
            m_MySheet.Cells[1, 21] = DateTime.Today.ToString("作成日: yyyy/M/d");

            int rows = ckdt.Rows.Count;
            int cols = ckdt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (j < 3 || Convert.ToDecimal(ckdt.Rows[i][j]) != 0)
                    {
                        m_MySheet.Cells[i + 4, j + 1] = ckdt.Rows[i][j];
                    }
                }
            }

            m_MySheet.PageSetup.PrintArea = @"$A$1:$U$" + (rows+3).ToString();

            //int printwide = 50;

            //for (int i = 0; i < 2000; i++)
            //{
            //    if (rows <= i)
            //    {
            //        m_MySheet.PageSetup.PrintArea = @"$A$1:$X$" + (i).ToString();
            //        break;
            //    }
            //}

            string localPass = @"C:\ODIS\CHECKLIST\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_") + "勤怠チェック";
            //string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

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
            button4.Enabled = true;
        }

        private void txtcsv_DragDrop(object sender, DragEventArgs e)
        {
            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            TextBox TargetTextBox = sender as TextBox;

            if (TargetTextBox == null)
            {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定
        }

        private void txtcsv_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグ中のファイルやディレクトリの取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop);

            //ファイルがドラッグされている場合、
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 配列分ループ
                foreach (string sTemp in sFileName)
                {
                    // ファイルパスかチェック
                    if (File.Exists(sTemp) == false)
                    {
                        // ファイルパス以外なので何もしない
                        return;
                    }
                    else
                    {
                        break;
                    }
                }

                // カーソルを[+]へ変更する
                // ここでEffectを変更しないと、以降のイベント（Drop）は発生しない
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            ////TextFieldParser parser = new TextFieldParser(@"C:\temp\勤怠テンプレート.csv", Encoding.GetEncoding("Shift_JIS"));
            //TextFieldParser parser = new TextFieldParser(@txtcsv.Text, Encoding.GetEncoding("Shift_JIS"));
            //parser.TextFieldType = FieldType.Delimited;
            //parser.SetDelimiters(","); // 区切り文字はコンマ

            //// データをすべてクリア
            //dgvcsv.Rows.Clear();

            //DataTable dt = new DataTable();

            //bool flg = true;

            //while (!parser.EndOfData)
            //{
            //    string[] row = parser.ReadFields(); // 1行読み込み

            //    //列追加
            //    if (flg)
            //    {
            //        dgvcsv.ColumnCount = row.Length;
            //        for (int i = 0; i < row.Length; i++)
            //        {
            //            dgvcsv.Columns[i].HeaderText = row[i];
            //            dt.Columns[i].ColumnName = row[i];
            //        }
            //        flg = false;
            //    }
            //    else
            //    {
            //        // 読み込んだデータ(1行をDataGridViewに表示する)
            //        dgvcsv.Rows.Add(row);
            //        dt.Rows.Add(row);

            //        MessageBox.Show(row[1]);
            //    }
            //}

            ////メール送信
            ////SendMail("Insertエラー", batchStdOut.ToString());

            //using (var bulkCopy = new SqlBulkCopy(ODIS.Com.SQLConstr))
            //{
            //    bulkCopy.DestinationTableName = "勤怠データ_test"; //dt.TableName; // テーブル名をSqlBulkCopyに教える
            //    bulkCopy.WriteToServer(dt); // bulkCopy実行
            //}


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
                message.From = new MailAddress("webmaster@oki-daiken.co.jp", "勤怠管理CSVエラー");
                message.To.Add(new MailAddress("kyan@oki-daiken.co.jp"));

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

        private void boushi_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            boushi_flg = true;
        }

        private void kyuugyou_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            kyuugyou_flg = true;
        }

        private void sonota_Enter(object sender, EventArgs e)
        {
            selectAlltbox((TextBox)sender);
            sonota_flg = true;
        }

        private void boushi_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void kyuugyou_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void sonota_KeyDown(object sender, KeyEventArgs e)
        {
            kDown(e);
        }

        private void boushi_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void kyuugyou_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void sonota_KeyPress(object sender, KeyPressEventArgs e)
        {
            Nonbeep(e);
        }

        private void boushi_MouseDown(object sender, MouseEventArgs e)
        {
            if (boushi_flg)
            {
                selectAlltbox((TextBox)sender);
                boushi_flg = false;
            }
        }

        private void kyuugyou_MouseDown(object sender, MouseEventArgs e)
        {
            if (kyuugyou_flg)
            {
                selectAlltbox((TextBox)sender);
                kyuugyou_flg = false;
            }
        }

        private void sonota_MouseDown(object sender, MouseEventArgs e)
        {
            if (sonota_flg)
            {
                selectAlltbox((TextBox)sender);
                sonota_flg = false;
            }
        }

        private void boushi_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void kyuugyou_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void sonota_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e, "D");
        }

        private void boushi_TextChanged(object sender, EventArgs e)
        {
            //SumMutoku();
        }

        private void kyuugyou_TextChanged(object sender, EventArgs e)
        {
            //SumMutoku();
        }

        private void sonota_TextChanged(object sender, EventArgs e)
        {
            //SumMutoku();
        }

        //private void SumMutoku()
        //{
        //    if (boushi.Text == "" || kyuugyou.Text == "" || sonota.Text == "") return;
        //    mutoku.Text = (Convert.ToDecimal(boushi.Text) + Convert.ToDecimal(kyuugyou.Text) + Convert.ToDecimal(sonota.Text)).ToString();
        //}

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void Client_Load(object sender, EventArgs e)
        {

        }

        private void combomutoku_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
