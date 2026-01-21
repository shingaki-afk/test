using C1.C1Excel;
using Microsoft.VisualBasic;
using ODIS.ODIS;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS
{
    public partial class Main : Form
    {
        /// <summary>
        /// 共通クラスのインスタンス
        /// </summary>
        private Common co = new Common();

        private DataTable dt = new DataTable();
        private DataTable dt2 = new DataTable();
        private DataTable aipodt = new DataTable();
        private string nl = Environment.NewLine;
        private DateTime date = new DateTime();
        private bool flg = new bool();

        public Main()
        {
            InitializeComponent();

            //ガベージ
            System.GC.Collect();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //フォントサイズ変更
            dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);

            comboBox1.Items.Add("【情報】操作履歴");
            //comboBox1.Items.Add("【情報】更新履歴");
            comboBox1.Items.Add("【情報】従業員誕生日　直近一週間");
            comboBox1.Items.Add("【情報】免許(資格)・登録手当一覧");
            comboBox1.Items.Add("【情報】免許手当例外一覧");
            comboBox1.Items.Add("【情報】有休残日数一覧");
            comboBox1.Items.Add("【情報】回数1項目一覧");
            comboBox1.Items.Add("【情報】正社員雇用変更一覧(2015/01/01～)");
            comboBox1.Items.Add("【情報】特別手当一覧");
            //comboBox1.Items.Add("【情報】健康診断受診状況一覧");
            //comboBox1.Items.Add("【情報】ストレスチェック8月以降変動一覧");

            comboBox1.Items.Add("【情報】月給者で退職記念品対象者");
            comboBox1.Items.Add("【情報】WEB明細対象者");
            //comboBox1.Items.Add("【情報】固定控除_入社異動_変更理由取得");
            comboBox1.Items.Add("【情報】緊急連絡先入力状況一覧");


            comboBox1.Items.Add("【警告】事業登録 有効期限間近順");
            comboBox1.Items.Add("【警告】従業員定年日　間近順");
            comboBox1.Items.Add("【警告】単年契約終了日　間近順");
            comboBox1.Items.Add("【警告】資格有効期限　間近順");
            comboBox1.Items.Add("【警告】試用期間中従業員");
            comboBox1.Items.Add("【警告】口座名義チェック結果");
            comboBox1.Items.Add("【警告】在留カード有効期限　間近順");

            comboBox1.Items.Add("【エラー】通勤管理エラー一覧");
            comboBox1.Items.Add("【エラー】通勤単価エラー一覧");
            comboBox1.Items.Add("【エラー】資格有効期限切れor未入力");
            comboBox1.Items.Add("【エラー】免許・登録・扶養・通信・離島手当不一致");
            comboBox1.Items.Add("【エラー】未発令一覧");
            comboBox1.Items.Add("【エラー】人事情報エラーチェック結果");
            comboBox1.Items.Add("【エラー】人事情報エラーチェック結果_単年契約社員");
            comboBox1.Items.Add("【エラー】課税区分該当だけど障害情報無");
            comboBox1.Items.Add("【エラー】最賃割れ一覧");

            if (Program.loginname == "喜屋武　大祐" || Program.loginname == "石川　尚吾" || Program.loginname == "高江洲　華子" || Program.loginname == "長山　友恵" || Program.loginname == "太田　朋宏" || Program.loginname == "新垣　聖悟" || Program.loginname == "仲里　かおり")
            {
                comboBox1.Items.Add("【情報】キントーン従業員同期履歴");
                //comboBox1.Items.Add("【情報】キントーン請求書同期履歴");

                //comboBox1.Items.Add("【情報】固定控除_入社_理由取得");
                //comboBox1.Items.Add("【情報】固定控除_異動_理由取得");

                comboBox1.Items.Add("【情報】固定給一覧");
                comboBox1.Items.Add("【情報】退職金合計");
                comboBox1.Items.Add("【情報】次月以降固定給変更一覧");
                comboBox1.Items.Add("【情報】昇格処理状況チェック");
                //comboBox1.Items.Add("固定給平均_組織現場雇用役職別");
                //comboBox1.Items.Add("現場計数にふくまれないひとたち");

                comboBox1.Items.Add("【エラー】端末個人情報不一致");
                comboBox1.Items.Add("【情報と鰓】端末保管中と退職済");
                comboBox1.Items.Add("【鰓】担当テーブル名称不一致");
                comboBox1.Items.Add("【情報】異動削除時の残データ有無");
                comboBox1.Items.Add("【情報】八重山緊急連絡先一覧");
                comboBox1.Items.Add("【警告】実績にあって、予算にないやつ");
            }

            //有給処理最新処理月取得
            DataTable kizyun = new DataTable();
            kizyun = Com.GetDB("select max(付与開始日) as 最新付与開始日 from QUATRO.dbo.NKTTKYUKAF where 会社コード = 'E0'");
            date = Convert.ToDateTime(kizyun.Rows[0]["最新付与開始日"].ToString());

            //年休一覧用コンボボックス
            comboBox3.Items.Add("");
            DataTable tantoubusyo = new DataTable();
            tantoubusyo = Com.GetDB("select distinct 担当区分 from dbo.担当テーブル ");
            foreach (DataRow row in tantoubusyo.Rows)
            {
                comboBox3.Items.Add(row["担当区分"].ToString());
            }

            flg = true;
            comboBox3.SelectedIndex = 0;

            treeView1.Nodes.Add("0_会計関連");
            treeView1.Nodes[0].ImageIndex = 0;
            treeView1.Nodes[0].Nodes.Add("01_会計検索(23/04～)");
            treeView1.Nodes[0].Nodes.Add("01_会計検索(13/04～23/03)");
            treeView1.Nodes[0].Nodes.Add("02_会計検索(05/04～13/03)");
            treeView1.Nodes[0].Nodes.Add("03_科目別損益");
            treeView1.Nodes[0].Nodes.Add("03_科目別損益(～23 / 03)"); 

            treeView1.Nodes[0].Nodes.Add("08_会計チェック");
            treeView1.Nodes[0].Nodes.Add("09_会計連携");

            treeView1.Nodes.Add("1_売上関連");
            treeView1.Nodes[1].ImageIndex = 1;
            treeView1.Nodes[1].Nodes.Add("11_現売上検索");
            treeView1.Nodes[1].Nodes[0].ToolTipText = "2013年1月～";
            treeView1.Nodes[1].Nodes.Add("12_旧売上検索");
            treeView1.Nodes[1].Nodes[1].ToolTipText = "2005年8月～2012年12月";
            treeView1.Nodes[1].Nodes.Add("13_契約一覧");
            treeView1.Nodes[1].Nodes.Add("14_業務管理台帳");
            treeView1.Nodes[1].Nodes.Add("15_計数管理台帳");
            //treeView1.Nodes[1].Nodes.Add("19_請求書データ転送");

            treeView1.Nodes.Add("2_人事関連");
            treeView1.Nodes[2].Nodes.Add("21_従業員検索");

            treeView1.Nodes[2].Nodes.Add("22_資格検索");
            treeView1.Nodes[2].Nodes.Add("23_研修登録・更新");
            treeView1.Nodes[2].Nodes.Add("24_資格有効期限管理");
            treeView1.Nodes[2].Nodes.Add("25_適正人員入力");
            treeView1.Nodes[2].Nodes.Add("26_最賃");
            treeView1.Nodes[2].Nodes.Add("27_技能実習生");

            treeView1.Nodes.Add("3_給与関連");
            treeView1.Nodes[3].ImageIndex = 2;
            treeView1.Nodes[3].Nodes.Add("30_入社入力");
            treeView1.Nodes[3].Nodes.Add("31_異動入力");
            treeView1.Nodes[3].Nodes.Add("32_退職入力");
            treeView1.Nodes[3].Nodes.Add("33_勤怠入力");
            treeView1.Nodes[3].Nodes.Add("34_出向応援入力");
            //treeView1.Nodes[3].Nodes.Add("35_固定控除入力");
            //treeView1.Nodes[3].Nodes.Add("36_保険他入力");
            //treeView1.Nodes[3].Nodes.Add("37_臨時手当入力");
            //treeView1.Nodes[3].Nodes.Add("38_変動控除入力");
            treeView1.Nodes[3].Nodes.Add("39_給与明細表示");
            treeView1.Nodes[3].Nodes[4].ToolTipText = "2012年～現在までの給与明細を閲覧可能です。";
            treeView1.Nodes[3].Nodes.Add("40_チェックリスト");
            treeView1.Nodes[3].Nodes.Add("41_勤怠検索");
            treeView1.Nodes[3].Nodes.Add("42_有給年間取得状況一覧");
            treeView1.Nodes[3].Nodes.Add("43_給与新基準_日給額");
            treeView1.Nodes[3].Nodes.Add("44_給与計算画面");
            treeView1.Nodes[3].Nodes.Add("45_給与計算後資料");
            treeView1.Nodes[3].Nodes.Add("46_駐車場控除入力");
            treeView1.Nodes[3].Nodes.Add("47_退職金");

            if (Program.loginname == "喜屋武　大祐")
            {
                treeView1.Nodes[3].Nodes.Add("48_月支給額平均");
                treeView1.Nodes[3].Nodes.Add("49_年度別総支給額");
            }

            treeView1.Nodes.Add("5_収支予算関連");
            treeView1.Nodes[4].ImageIndex = 3;
            treeView1.Nodes[4].Nodes.Add("51_管理計数");
            treeView1.Nodes[4].Nodes.Add("52_予算更新");
            treeView1.Nodes[4].Nodes.Add("531_予算集計");
            treeView1.Nodes[4].Nodes.Add("532_集計差額");
            treeView1.Nodes[4].Nodes.Add("541_集計職種別");
            treeView1.Nodes[4].Nodes.Add("542_集計職種別_差額");

            //treeView1.Nodes[4].Nodes.Add("55_実績×予算/前年");
            treeView1.Nodes[4].Nodes.Add("55_月別比");
            //treeView1.Nodes[4].Nodes.Add("XX_予算縦横変換中。。");
            treeView1.Nodes[4].Nodes.Add("57_売上グラフ"); 
            treeView1.Nodes[4].Nodes.Add("58_人工単価");
            

            treeView1.Nodes.Add("6_庶務関連");
            treeView1.Nodes[5].Nodes.Add("61_過去稟議(～2020年8月末)");
            treeView1.Nodes[5].Nodes.Add("63_事業登録・ビル管登録一覧");
            treeView1.Nodes[5].Nodes[1].Nodes.Add("631_事業登録・ビル管登録一覧");
            treeView1.Nodes[5].Nodes[1].Nodes.Add("632_ビル管登録現場_更新");
            treeView1.Nodes[5].Nodes[1].Nodes.Add("633_ビル管登録現場_一覧");
            treeView1.Nodes[5].Nodes.Add("64_通勤管理");
            treeView1.Nodes[5].Nodes.Add("66_ハードソフト情報管理");
            treeView1.Nodes[5].Nodes.Add("67_障害者雇用状況");

            treeView1.Nodes[5].Nodes.Add("テスト中");

            treeView1.Nodes.Add("8_リンク");
            treeView1.Nodes[6].Nodes.Add("83_キントーンログイン画面");
            treeView1.Nodes[6].Nodes.Add("83_GSuite GMail");
            treeView1.Nodes[6].Nodes.Add("84_GSuite カレンダー");
            //treeView1.Nodes[6].Nodes.Add("85_rakumo");

            if (Program.loginname == "喜屋武　大祐" || Program.loginname == "太田　朋宏" || Program.loginname == "新垣　聖悟" || Program.loginname == "管理者" || Program.loginname == "RPA用AC")
            {
                treeView1.Nodes.Add("9_システム管理");
                treeView1.Nodes[7].ImageIndex = 4;
                treeView1.Nodes[7].Nodes.Add("91_単年契約情報管理");
                treeView1.Nodes[7].Nodes.Add("92_端末管理テーブル設定");
                treeView1.Nodes[7].Nodes.Add("93_区分管理_資格登録情報");
                //treeView1.Nodes[7].Nodes.Add("94_ファイルサーバーログ");
                treeView1.Nodes[7].Nodes.Add("95_担当テーブル設定");
                treeView1.Nodes[7].Nodes.Add("96_ODISアカウント設定");
                //treeView1.Nodes[7].Nodes.Add("97_rakumo 管理者画面");

                treeView1.Nodes[7].Nodes.Add("手当控除チェック");
                treeView1.Nodes[7].Nodes.Add("振替伝票");
                treeView1.Nodes[7].Nodes.Add("理由管理");
                treeView1.Nodes[7].Nodes.Add("理由管理_人事情報エラーチェック");
                treeView1.Nodes[7].Nodes.Add("口座名義相違認識");
                treeView1.Nodes[7].Nodes.Add("回数管理");
                treeView1.Nodes[7].Nodes.Add("キントーン");
                treeView1.Nodes[7].Nodes.Add("Web明細処理");
                //treeView1.Nodes[7].Nodes.Add("55_定員設定");
                treeView1.Nodes[7].Nodes.Add("SmartHR連携");
                treeView1.Nodes[7].Nodes.Add("Excel計数AI変換テスト");
                treeView1.Nodes[7].Nodes.Add("SmartHRデータ移行"); 

            }
            //全てのノードを展開
            treeView1.ExpandAll();

            //自分自身のバージョン情報を取得する
            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                System.Reflection.Assembly.GetExecutingAssembly().Location);

            GetCombo();

            //SetAppVer();

            foreach (DataRow drw in Program.acdt.Rows)
            {
                comboBox2.Items.Add(drw["名前"].ToString());
            }

            comboBox2.SelectedItem = Program.loginname;

            //TODO 切替後がどうなるか。。
            if (Program.loginname == "喜屋武　大祐" || Program.loginname == "石川　尚吾" || Program.loginname == "高江洲　華子" || Program.loginname == "太田　朋宏" || Program.loginname == "仲里　かおり" || Program.loginname == "新垣　聖悟")
            { 
                comboBox2.Enabled = true;
            }
        }

        private void GetCombo()
        {
            // Random クラスの新しいインスタンスを生成する
            Random cRandom = new System.Random();

            //コンボボックス設定
            comboBox1.SelectedIndex = cRandom.Next(comboBox1.Items.Count - 6);

        }


        private void SetAppVer()
        {
            //初期値
            string ver = "ver:ClickOnceバージョン";
            string update_date = "最終更新日:ClickOnce最終更新日";
            //現在のアプリケーションが ClickOnce アプリケーションかチェック
            //デバッグ時に呼び出すと例外になっちゃうので・・・。
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
                //**バージョン取得
                ver = "ver:" + ad.CurrentVersion.ToString();
                //**最終更新日取得
                update_date = "最終更新日:" + ad.TimeOfLastUpdateCheck.ToLongDateString().ToString();
            }

            //結果の表示
            //label2.Text = ver;
            //label3.Text = update_date;
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //子ノードがあれば対象外
            if (treeView1.SelectedNode.GetNodeCount(true) > 0) return;

            Cursor.Current = Cursors.WaitCursor;

            switch (treeView1.SelectedNode.Text)
            {
                case "01_会計検索(23/04～)": Kaikei_PCA nk_PCA = new Kaikei_PCA(); nk_PCA.Show(); return;
                case "01_会計検索(13/04～23/03)": Kaikei nk = new Kaikei(); nk.Show(); return;
                case "02_会計検索(05/04～13/03)": KakoKaikei kk = new KakoKaikei(); kk.Show(); return;
                case "03_科目別損益": Kamoku kamo = new Kamoku(); kamo.Show(); return;
                case "03_科目別損益(～23 / 03)": Kamoku_Ex kamo_PCA = new Kamoku_Ex(); kamo_PCA.Show(); return;
                case "08_会計チェック": KaikeiCheck kaic = new KaikeiCheck(); kaic.Show(); return;
                case "09_会計連携": KaikeiRenkei kaikeir = new KaikeiRenkei(); kaikeir.Show(); return;

                case "11_現売上検索": Uriage nu = new Uriage(); nu.Show(); return;
                case "12_旧売上検索": kakouriage ku = new kakouriage(); ku.Show(); return;
                case "13_契約一覧": KeiyakuList keil = new KeiyakuList(); keil.Show(); return;
                case "14_業務管理台帳": GyoumuKanri gyoukan = new GyoumuKanri(); gyoukan.Show(); return;
                case "15_計数管理台帳": KeisuuKanri KeisuuK = new KeisuuKanri(); KeisuuK.Show(); return;
                case "17_売上契約": UriageKeiyaku urikei = new UriageKeiyaku(); urikei.Show(); return;

                case "21_従業員検索": Emp emp = new Emp(); emp.Show(); return;
                case "22_資格検索": shikaku sk = new shikaku(); sk.Show(); return;
                case "23_研修登録・更新": Kensyuu ks = new Kensyuu(); ks.Show(); return;
                case "24_資格有効期限管理": ShikakuKigen shiki = new ShikakuKigen(); shiki.Show(); return;
                case "25_適正人員入力": Tekisei ts = new Tekisei(); ts.Show(); return;
                case "26_最賃": Saichin scin = new Saichin(); scin.Show(); return;
                case "27_技能実習生": GinouZ gz = new GinouZ(); gz.Show(); return;
                //case "研修検索": Semi semi = new Semi(); semi.Show(); return;

                case "30_入社入力": Nyuusya ny = new Nyuusya(); ny.Show(); return;
                case "31_異動入力": Genpyou gp = new Genpyou(); gp.Show(); return;
                case "32_退職入力": Taisya ta = new Taisya(); ta.Show(); return;
                case "33_勤怠入力": Client cl = new Client(Program.loginID); cl.Show(); return;
                case "34_出向応援入力": Furikae Furikae = new Furikae(); Furikae.Show(); return;
                //case "35_固定控除入力": KoteiKoujo Rteate = new KoteiKoujo(); Rteate.Show(); return;
                //case "36_保険他入力": Koteietc koteietc = new Koteietc(); koteietc.Show(); return;
                //case "37_臨時手当入力": RinjiTeate rint = new RinjiTeate(); rint.Show(); return;
                case "38_変動控除入力": HendouKoujo henk = new HendouKoujo(); henk.Show(); return;
                case "39_給与明細表示": KMeisai km = new KMeisai(); km.Show(); return;
                case "40_チェックリスト": CheckList CheckList = new CheckList(); CheckList.Show(); return;
                case "41_勤怠検索": KintaiCK kck = new KintaiCK(); kck.Show(); return;
                case "42_有給年間取得状況一覧": YuuKyuu yuku = new YuuKyuu(); yuku.Show(); return;
                case "43_給与新基準_日給額": Nikkyu nikkyu = new Nikkyu(); nikkyu.Show(); return;
                case "44_給与計算画面": ZeeMEC zmec = new ZeeMEC(); zmec.Show(); return;
                case "45_給与計算後資料": AfterCalc_PCA acpca = new AfterCalc_PCA(); acpca.Show(); return;
                case "46_駐車場控除入力": ZigyouUp zig = new ZigyouUp(); zig.Show(); return;
                case "47_退職金": TaisyokuK taik = new TaisyokuK(); taik.Show(); return;
                case "48_月支給額平均": Kyuuyo ky = new Kyuuyo(); ky.Show(); return;
                case "49_年度別総支給額": Soushikyuu sou = new Soushikyuu(); sou.Show(); return;

                case "51_管理計数": KanriKeisuu gkei54 = new KanriKeisuu(); gkei54.Show(); return;
                case "52_予算更新": YosanUp yu_Next = new YosanUp(); yu_Next.Show(); return;
                //case "53_予算集計_54期": YosanTotal54 yt_Next = new YosanTotal54(); yt_Next.Show(); return;

                //case "51_管理計数_53期": KanriKeisuu53 gkei = new KanriKeisuu53(); gkei.Show(); return;
                //case "52_予算更新_53期": YosanUp53 yu = new YosanUp53(); yu.Show(); return;
                case "531_予算集計": YosanTotal yt = new YosanTotal(); yt.Show(); return;
                case "532_集計差額": YosanTotal_Zenkihi ytz = new YosanTotal_Zenkihi(); ytz.Show(); return;
                case "541_集計職種別": YosanTotalSyokusyu yts = new YosanTotalSyokusyu(); yts.Show(); return;
                case "542_集計職種別_差額": YosanTotalSyokusyu_Zenkihi ytsz = new YosanTotalSyokusyu_Zenkihi(); ytsz.Show(); return;

                case "55_実績×予算/前年": ZissekiHi zh = new ZissekiHi(); zh.Show(); return;

                case "55_月別比": YosanTsukibetsu yozenhi = new YosanTsukibetsu(); yozenhi.Show(); return;
                case "XX_予算縦横変換中。。": KamokuYosan kmkuk = new KamokuYosan(); kmkuk.Show(); return;

                case "57_売上グラフ": graph gr = new graph(); gr.Show(); return;
                case "58_人工単価": Ninku ninq = new Ninku(); ninq.Show(); return;
                

                case "61_過去稟議(～2020年8月末)": Ringi rin = new Ringi(); rin.Show(); return;
                case "62_rakumo ワークフロー 運用画面": System.Diagnostics.Process.Start(@"https://a-rakumo.appspot.com/workflow/admin"); return;
                case "63_事業登録一覧":

                    string pass = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\事業登録一覧表.xlsx";

                    DataTable dt = new DataTable();
                    dt = Com.GetDB("select Convert(int,社員番号), 漢字氏名, 地区名, 組織名, 現場名, 給与支給名称, 役職名, 年齢, 在籍年月 from dbo.accessNew where 在籍区分 <> '9'");

                    //手順1：新しいワークブックを作成します。
                    C1XLBook c1XLBook1 = new C1XLBook();

                    c1XLBook1.Load(pass);

                    // 手順2：セルに値を挿入します。
                    XLSheet sheet = c1XLBook1.Sheets["Data"];

                    int rows = dt.Rows.Count;
                    int cols = dt.Columns.Count;

                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            sheet[i, j].Value = dt.Rows[i][j];
                        }
                    }

                    // 手順3：ファイルを保存します。
                    c1XLBook1.Save(pass);

                    System.Diagnostics.Process.Start(pass); return;
                case "631_事業登録・ビル管登録一覧": ZigyouTouroku zt = new ZigyouTouroku(); zt.Show(); return;
                case "632_ビル管登録現場_更新": BillKan bk = new BillKan(); bk.Show(); return;
                case "633_ビル管登録現場_一覧": BillKan_list bk_list = new BillKan_list(); bk_list.Show(); return;
                //case "特定建築物_様式": System.Diagnostics.Process.Start(@"\\192.168.100.11\21_全体共通\10_標準書式\60_法令関連\10_特定建築物\"); return;
                
                case "64_通勤管理": Scramble sc = new Scramble(); sc.Show(); return;
                //case "65_事故・クレーム管理台帳": Report re = new Report(); re.Show(); return;
                case "66_ハードソフト情報管理": HardSoft haso = new HardSoft(); haso.Show(); return;
                case "67_障害者雇用状況": SList sl = new SList(); sl.Show(); return;
                //case "81_標準書式": System.Diagnostics.Process.Start(@"https://oki-daiken.cybozu.com/k/231/"); return;
                //case "82_諸規程集": System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\諸規程集\要項及び諸規程運用状況一覧.xlsx"); return;
                case "83_キントーンログイン画面": System.Diagnostics.Process.Start(@"https://oki-daiken.cybozu.com/"); return;

                case "83_GSuite GMail": System.Diagnostics.Process.Start(@"https://mail.google.com/a/oki-daiken.co.jp"); return;
                case "84_GSuite カレンダー": System.Diagnostics.Process.Start(@"https://calendar.google.com/a/oki-daiken.co.jp"); return;
                //case "85_rakumo": System.Diagnostics.Process.Start(@"https://a-rakumo.appspot.com"); return;
                //case "86_rakumo サポートサイト": System.Diagnostics.Process.Start(@"https://support.rakumo.com/rakumo-support/"); return;

                case "91_単年契約情報管理": Kourei kourei = new Kourei(); kourei.Show(); return;
                case "92_端末管理テーブル設定": TanmatsuK tankan = new TanmatsuK(); tankan.Show(); return;
                case "93_区分管理_資格登録情報": KubunKanri_Shikaku kubun_shikaku = new KubunKanri_Shikaku(); kubun_shikaku.Show(); return;
                //case "94_ファイルサーバーログ": FSLog fsl = new FSLog(); fsl.Show(); return;
                case "95_担当テーブル設定": TantouT tanT = new TantouT(); tanT.Show(); return;
                case "96_ODISアカウント設定": AccountSet acset = new AccountSet(); acset.Show(); return;
                case "97_rakumo 管理者画面": System.Diagnostics.Process.Start(@"https://a-rakumo.appspot.com/admin/"); return;
                //case "目論見更新": Yosan moku = new Yosan(); moku.Show(); return;
                case "手当控除チェック": TeateKoujoCheck krin = new TeateKoujoCheck(); krin.Show(); return;
                case "振替伝票": Denpyou denp = new Denpyou(); denp.Show(); return;
                case "理由管理": RiyuuKanri riyuu = new RiyuuKanri(); riyuu.Show(); return;
                case "理由管理_人事情報エラーチェック": RiyuuKanriZin riyuuz = new RiyuuKanriZin(); riyuuz.Show(); return;
                case "口座名義相違認識": KouzaS kous = new KouzaS(); kous.Show(); return;
                case "回数管理": KaisuuKanri kaikan = new KaisuuKanri(); kaikan.Show(); return;
                case "テスト中": etc etc = new etc(); etc.Show(); return;
                case "キントーン": Kintone kint = new Kintone(); kint.Show(); return;
                case "Web明細処理": WebMeisai webm = new WebMeisai(); webm.Show(); return;
                case "SmartHR連携": SmartHR_Get shr = new SmartHR_Get(); shr.Show(); return;
                case "Excel計数AI変換テスト": Form1 fo = new Form1(); fo.Show(); return;

                case "SmartHRデータ移行": SmartHRData shrd = new SmartHRData(); shrd.Show(); return;
                default: break;
            }

            Cursor.Current = Cursors.Default;
        }






        private void Main_Shown(object sender, EventArgs e)
        {
            //dataGridView1.CurrentCell = null;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void GetComboData()
        {
            //コンボボックス無効化・カーソル変更
            comboBox1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.DataSource = null;
            dt2.Clear();

            //隠す
            tableLayoutPanel1.Visible = false;
            comboBox3.Visible = false;
            textBox1.Visible = false;
            button1.Visible = false;

            //有給使用変数
            string year;
            string month;
            string year_nx;
            string month_nx;
            string taisyoku;
            year = date.Year.ToString();//2019
            month = date.ToString("MM");//01
            year_nx = date.AddMonths(1).Year.ToString();//2019
            month_nx = date.AddMonths(1).ToString("MM");//02
            taisyoku = date.AddMonths(-1).ToString("yyyy/MM/dd");// 2018/12/01

            string month_ex;
            month_ex = date.AddMonths(-1).Month.ToString();

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            switch (comboBox1.SelectedItem.ToString())
            {
                //2019/01/01
                case "【情報】操作履歴":
                    dt = Com.GetDB("select CONVERT(varchar,日時,111) as 日時, convert(char(8),日時,108) as 時刻, 機能, 名前, 検索結果, 検索条件 from dbo.GetHistory");
                    break;
                case "【情報】更新履歴":
                    dt = Com.GetDB("select * from dbo.更新履歴 order by 更新日時 desc");
                    break;
                case "【警告】事業登録 有効期限間近順":
                    dt = Com.GetDB("select * from dbo.z事業登録情報 order by 有効期限");
                    break;
                case "【情報】従業員誕生日　直近一週間":
                    //DateTime.Today.ToString().Substring(5, 5) 
                    if (DateTime.Today.Month.ToString() == "12" && DateTime.Today.AddDays(7).Month.ToString() == "1")
                    {
                        dt = Com.GetDB("select * from [t誕生日取得_年跨ぎ前] order by 誕生日");
                        dt2 = Com.GetDB("select * from [t誕生日取得_年跨ぎ後] order by 誕生日");

                        dt.Merge(dt2);
                    }
                    else
                    {
                        dt = Com.GetDB("select * from t誕生日取得 order by 誕生日");
                    }
                    break;
                case "【警告】従業員定年日　間近順":
                    dt = Com.GetDB("select * from dbo.t定年情報  order by 定年日");
                    break;
                case "【警告】単年契約終了日　間近順":
                    dt = Com.GetDB("select * from t単年契約情報取得 order by 契約終了日");
                    break;
                case "【警告】資格有効期限　間近順":
                    dt = Com.GetDB("select * from s資格有効期限間近順 order by 資格有効期限");
                    break;
                case "【警告】在留カード有効期限　間近順":
                    dt = Com.GetDB("select * from z在留カード有効期限間近順 order by 在留カード有効期限");
                    break;
                case "【警告】試用期間中従業員":
                    dt = Com.GetDB("select * from dbo.s試用期間中一覧 order by 経過日数 desc");
                    break;
                case "【エラー】免許・登録・扶養・通信・離島手当不一致":
                    dt = Com.GetDB("select * from m免許手当不一致 union all select * from t登録手当不一致 union all select * from h扶養手当不一致 union all select * from t通信手当不一致 union all select * from dbo.r離島手当不一致");
                    break;
                case "【情報】免許手当例外一覧":
                    dt = Com.GetDB("select * from dbo.m免許手当例外一覧");
                    break;
                case "【情報】免許(資格)・登録手当一覧":
                    dt = Com.GetDB("select * from dbo.[m免許(資格)・登録手当一覧] order by ID, 連番, 管理コード");
                    break;
                case "【情報】雇用保険と所定労働日数":
                    dt = Com.GetDB("select * from dbo.k雇用保険と所定労働日数");
                    break;
                case "【警告】口座名義チェック結果":
                    dt = Com.GetDB("select * from dbo.k口座名義チェック結果 where 銀行名 = '琉銀'");
                    //dt = Com.GetDB("select * from dbo.k口座名義チェック結果");
                    break;

                case "【情報】通信手当一覧":
                    TargetDays tday = new TargetDays();

                    //string select = "select b.社員番号, b.氏名, b.地区名, b.組織名, b.現場名, a.金額 as 手当額, c.金額 as [ZeeM臨時作業手当設定額" + tday.StartYMD.AddMonths(1).ToString("yy年M月給与") + "], a.備考 as 携帯番号, b.退職年月日 ";
                    //string from = "from dbo.通信手当 a left join dbo.社員基本情報 b on a.社員番号 = b.社員番号 left join QUATRO.dbo.QCTTKINGKH c on a.社員番号 = c.社員番号 and c.処理年 = '" + tday.StartYMD.AddMonths(1).Year.ToString() + "' and c.処理月 = '" + tday.StartYMD.AddMonths(1).ToString("MM") + "' and c.項目ＩＤ = 'A7900' order by 地区CD, 組織CD, 現場CD";

                    string select = "select b.社員番号, b.氏名, b.地区名, b.組織名, b.現場名, タイプ as 手当額, c.金額 as [ZeeM臨時作業手当設定額" + tday.StartYMD.AddMonths(1).ToString("yy年M月給与") + "], a.PC名電番Mail as 携帯番号, a.備考, case when a.組織CD <> b.組織CD or a.現場CD <> b.現場CD then '組織or現場相違' when PC名電番Mail is null or PC名電番Mail = '' then '要電番確認' else '' end as チェック, b.退職年月日 as 退職チェック  ";
                    string from = "from dbo.t端末管理テーブル a left join dbo.社員基本情報 b on a.社員番号 = b.社員番号 left join QUATRO.dbo.QCTTKINGKH c on c.会社コード = 'E0' and a.社員番号 = c.社員番号 and c.処理年 = '" + tday.StartYMD.AddMonths(1).Year.ToString() + "' and c.処理月 = '" + tday.StartYMD.AddMonths(1).ToString("MM") + "' and c.項目ＩＤ = 'A7900' where 使用区分 = '私物' order by 地区CD, b.組織CD, b.現場CD";

                    dt = Com.GetDB(select + from);

                    break;
                //case "【情報】健康診断受診状況一覧":
                //    dt = Com.GetDB("select * from dbo.k健康診断受診状況一覧");
                //    break;
                case "【情報】ストレスチェック8月以降変動一覧":
                    dt = Com.GetDB("select * from dbo.sストレスチェック対象者変動一覧");
                    break;
                    
                case "【エラー】通勤管理エラー一覧":
                    dt = Com.GetDB("select * from dbo.t通勤管理エラー一覧取得 order by 区分");
                    break;
                case "【エラー】通勤単価エラー一覧":
                    dt = Com.GetDB("select * from dbo.t通勤一日単価チェック");
                    break;
                case "【エラー】資格有効期限切れor未入力":
                    dt = Com.GetDB("select * from dbo.[s資格有効期限切れor未入力] ");
                    break;
                case "【エラー】人事情報エラーチェック結果":
                    dt = Com.GetDB("select a.*, b.認識項目 from dbo.z人事情報チェック結果 a left join dbo.z人事情報エラーチェック認識済 b on a.社員番号 = b.社員番号 and a.内容 = b.内容 where 契約社員 is null");
                    break;
                case "【エラー】人事情報エラーチェック結果_単年契約社員":
                    dt = Com.GetDB("select a.*, b.認識項目 from dbo.z人事情報チェック結果 a left join dbo.z人事情報エラーチェック認識済 b on a.社員番号 = b.社員番号 and a.内容 = b.内容 where 契約社員 is not null");
                    break;
                case "更新":
                    //プロステージからデータをとってくる
                    DataTable dttemp = new DataTable();
                    dttemp = Com.GetPosDB("select ym,bumoncode, koujicode, sum(売上) 売上, sum(諸経費) 諸経費 from kpcp01.uriagesyokeihi('202104') group by ym, bumoncode, koujicode");

                    //一旦データ削除
                    DataTable dtr = new DataTable();
                    dtr = Com.GetDB("delete from 売上と経費 where 年月 = '202104'");

                    //ZeeMDBにインサート
                    SqlConnection Cn;
                    SqlCommand Cmd;

                    try
                    {
                        using (Cn = new SqlConnection(ODIS.Com.SQLConstr))
                        {
                            Cn.Open();
                            using (Cmd = Cn.CreateCommand())
                            {
                                using (SqlBulkCopy bulkcopy = new SqlBulkCopy(Cn))
                                {
                                    bulkcopy.BulkCopyTimeout = 660;
                                    bulkcopy.DestinationTableName = "売上と経費";
                                    bulkcopy.WriteToServer(dttemp);
                                    bulkcopy.Close();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("エラー" + ex.ToString());
                        throw;
                    }

                    dt = Com.GetDB("select * from dbo.新計数仮 order by 組織CD, 現場CD");
                    break;

                case "【情報】有休残日数一覧":

                    //コンボボックス対応
                    comboBox3.Visible = true;
                    textBox1.Visible = true;
                    button1.Visible = true;

                    dt = Com.GetDB("select * from dbo.[年休残日数一覧]('" + year + "', '" + month + "', '" + year_nx + "', '" + month_nx + "', '" + taisyoku + "') where 担当区分 like '%" + comboBox3.SelectedItem.ToString() + "%' " + result);
                    dt.Columns["当月/1残数"].ColumnName = month_ex + "/1残数";
                    dt.Columns["当月使用数"].ColumnName = month_ex + "月使用数";
                    dt.Columns["当月末残数"].ColumnName = month_ex + "月末残数";
                    dt.Columns["当月末残内訳_前々回分"].ColumnName = month_ex + "月末残内訳_前々回分";
                    dt.Columns["当月末残内訳_前回分"].ColumnName = month_ex + "月末残内訳_前回分";
                    dt.Columns["当月末年休使用状況_前々回+前回付与数"].ColumnName = month_ex + "月末_前々回+前回付与数";
                    dt.Columns["当月末年休使用状況_前回付与後使用数"].ColumnName = month_ex + "月末_前回付与後使用数";
                    dt.Columns["当月末年休使用状況_消化率"].ColumnName = month_ex + "月末消化率";
                    dt.Columns["当月/31消滅"].ColumnName = month_ex + "/31消滅";
                    dt.Columns["次月/1付与"].ColumnName = date.Month.ToString() + "/1付与";
                    dt.Columns["次月/1残数"].ColumnName = date.Month.ToString() + "/1残数";
                    break;

                case "【エラー】未発令一覧":
                    dt = Com.GetDB("select * from dbo.[未発令一覧] ");
                    break;

                case "【情報】回数1項目一覧":
                    dt = Com.GetDB("select * from dbo.[k回数1項目一覧] order by 項目名");
                    break;

                //case "固定控除一覧表示":
                //    if (Program.loginname == "石川　尚吾" || Program.loginname == "喜屋武　大祐")
                //    {
                //        dt = Com.GetDB("select * from dbo.[固定控除一覧表示] order by 内容, 地区名, 組織名, 現場名");
                //    }
                //    else
                //    {
                //        dt = Com.GetDB("select * from dbo.[固定控除一覧表示] where 内容 not like '2%' order by 内容, 地区名, 組織名, 現場名");
                //    }
                //    break;

                //case "【情報】勤怠入力_入力状況一覧":
                //    //全対象者データ取得
                //    DataTable dtall = co.GetKintaiKihon(1, "");

                //    //担当別データ
                //    DataTable list = co.GetKintaiKihon(9, "");

                //    //エラー数と警告数を取得しリストに表示
                //    foreach (DataRow dr in list.Rows)
                //    {
                //        string filtStr = "担当管理 = '" + dr["担当"].ToString() + "' and 登録フラグ = '1'";
                //        DataRow[] drYae = dtall.Select(filtStr, "");


                //        int errorCt = 0;  //エラー件数
                //        int emergCt = 0; //警告件数

                //        foreach (DataRow row in drYae)
                //        {
                //            string[] st = co.ErrorCheck(row, "");

                //            if (st[0].Length > 0)
                //            {
                //                errorCt++;
                //            }

                //            if (st[4].Length > 0)
                //            {
                //                emergCt++;
                //            }
                //        }

                //        foreach (DataRow d in list.Rows)
                //        {
                //            if (d[1].ToString() == dr["担当"].ToString())
                //            {
                //                d["エラー"] = errorCt.ToString();
                //                d["警告"] = emergCt.ToString();
                //            }
                //        }
                //    }

                //    dt = list;
                //    break;
                //case "【情報】勤怠入力_警告一覧":
                //    //全対象者データ取得
                //    DataTable dtall2 = co.GetKintaiKihon(1, "");

                //    DataRow[] drYae2 = dtall2.Select("登録フラグ = '1'", "現場CD");

                //    DataTable Disp = new DataTable();
                //    Disp.Columns.Add("社員番号", typeof(string));
                //    Disp.Columns.Add("漢字氏名", typeof(string));
                //    Disp.Columns.Add("組織名", typeof(string));
                //    Disp.Columns.Add("現場名", typeof(string));
                //    Disp.Columns.Add("担当管理", typeof(string));
                //    Disp.Columns.Add("状況", typeof(string));
                //    Disp.Columns.Add("休日超過理由", typeof(string));
                //    Disp.Columns.Add("出勤超過理由", typeof(string));
                //    Disp.Columns.Add("メモ", typeof(string));

                //    foreach (DataRow row in drYae2)
                //    {
                //        string[] st = co.ErrorCheck(row, "");
                //        if (st[4].Length > 0 || row["休日超過理由"].ToString().Length > 0 || row["出勤超過理由"].ToString().Length > 0 || row["コメント"].ToString().Length > 0)
                //        {
                //            DataRow nr = Disp.NewRow();
                //            nr["社員番号"] = row["社員番号"];
                //            nr["漢字氏名"] = row["漢字氏名"];
                //            nr["組織名"] = row["組織名"];
                //            nr["現場名"] = row["現場名"];
                //            nr["担当管理"] = row["担当管理"];
                //            nr["状況"] = st[4].ToString();
                //            nr["休日超過理由"] = row["休日超過理由"];
                //            nr["出勤超過理由"] = row["出勤超過理由"];
                //            nr["メモ"] = row["コメント"];
                //            Disp.Rows.Add(nr);
                //        }
                //    }

                //    dt = Disp;
                //    break;
                //case "問合一覧表示":
                //    dt = Com.GetDB("select * from dbo.問合管理");
                //    break;

                case "【情報】特別手当一覧":
                    dt = Com.GetDB("select * from dbo.[t特別と転勤] ");
                    break;
                //case "職種設定情報(月給・日給)":
                //    dt = Com.GetDB("select * from dbo.職種設定表示");
                //    break;
                case "【情報】正社員雇用変更一覧(2015/01/01～)":
                    dt = Com.GetDB("select * from dbo.[雇用体系変更一覧] order by 発令日 desc");
                    break;

                //case "【情報】休業手当":
                //    string sql = "select * from dbo.休業手当表示まとめ ";
                    //dt = Com.GetDB(sql + "order by 組織CD, 現場CD");
                    //break;
                //case "【情報】休業手当_組織現場集計":
                //    dt = Com.GetDB("select * from dbo.休業手当表示_組織別現場別");
                //    break;
                case "【情報】休業申請_雇保加入":
                    dt = Com.GetDB("select * from dbo.休業手当_雇用保険加入 order by 組織CD, 現場CD, カナ氏名");
                    break;
                case "【情報】休業申請_雇保未加入":
                    dt = Com.GetDB("select * from dbo.休業手当_雇用保険未加入 order by 組織CD, 現場CD, カナ氏名");
                    break;
                //case "異動入力で登録された家族情報":
                //    dt = Com.GetDB("select * from QUATRO.dbo.SJMTKAZOKU where 個人識別ＩＤ = '' ");
                //    break;
                case "【情報】固定給一覧":
                    dt = Com.GetDB("select * from dbo.i石川支店長用給与一覧");
                    break;
                case "固定給平均_組織現場雇用役職別":
                    dt = Com.GetDB("select * from dbo.k固定給平均_組織現場雇用役職別");
                    break;
                case "【情報と鰓】端末保管中と退職済":
                    dt = Com.GetDB("select * from dbo.t端末保管中と退職済");
                    break;
                case "【エラー】課税区分該当だけど障害情報無":
                    dt = Com.GetDB("select * from dbo.k課税区分該当だけど障害情報無");
                    break;
                case "【エラー】最賃割れ一覧":
                    dt = Com.GetDB("select * from dbo.s最賃割れ一覧");
                    break;
                case "【情報】次月以降固定給変更一覧":
                    dt = Com.GetDB("select * from dbo.z次月以降固定給変更一覧");
                    break;
                case "【鰓】担当テーブル名称不一致":
                    dt = Com.GetDB("select * from dbo.t担当テーブル名称不一致");
                    break;
                case "現場計数にふくまれないひとたち":
                    dt = Com.GetDB("select* from dbo.社員基本情報 a left join dbo.k固定給一覧 b on a.社員番号 = b.社員番号 where 在籍区分<> '9' and 現場CD like '%9900' and 組織名<> '役員室'");
                    break;
                case "【情報】退職金合計":
                    dt = Com.GetDB("select sum(退職金_自己都合) as 退職金自己都合算出合計, sum(退職金_会社都合) as 退職金会社都合算出合計 from dbo.t退職金リスト");
                    break;
                case "【エラー】端末個人情報不一致":
                    dt = Com.GetDB("select * from dbo.t端末個人情報不一致");
                    break;
                case "【情報】異動削除時の残データ有無":
                    dt = Com.GetDB("select * from dbo.i異動削除時の残データ有無");
                    break;
                case "【情報】八重山緊急連絡先一覧":
                    dt = Com.GetDB("select 漢字氏名, カナ氏名, 組織CD, 組織名, 現場CD,現場名,郵便番号, 住所１,住所２,住所３,住所４,  b.* from dbo.accessNew a left join dbo.k緊急連絡先 b on a.社員番号 = b.社員番号 where 在籍区分 <> '9' and 組織CD like '3%' order by 組織CD, 現場CD");
                    break;
                case "【警告】実績にあって、予算にないやつ":
                    dt = Com.GetDB("select * from dbo.kanrikeisuu a left join dbo.yosankanri b on a.年月 = b.年月 and a.部門CD = b.部門CD and a.現場CD = b.現場CD where b.年月 is null and a.年月 > '202403'");
                    break;

                case "【情報】キントーン従業員同期履歴":
                    dt = Com.GetDB("select top 240 * from dbo.KintoneSync order by 処理日時 desc");
                    break;
                //case "【情報】キントーン請求書同期履歴":
                //    dt = Com.GetDB("select top 240 * from dbo.KintoneInvoiceSync order by 処理日時 desc");
                //    break;
                case "【情報】月給者で退職記念品対象者":
                    dt = Com.GetDB("select * from dbo.g月給者で退職記念品対象者テーブル");
                    break;
                case "【情報】WEB明細対象者":
                    dt = Com.GetDB("select case when a.承諾区分 = '1' then '同意済' else '' end 同意区分, a.社員番号, b.地区名, b.組織名, b.現場名,b.氏名,a.備考 from dbo.web明細対象者 a left join dbo.s社員基本情報 b on a.社員番号 = b.社員番号 order by 地区CD, 組織CD, 現場CD");
                    break;
                case "【情報】緊急連絡先入力状況一覧":
                    dt = Com.GetDB("select * from k緊急連絡先入力状況一覧 order by 担当区分 ");
                    break;
                case "【情報】固定控除_入社_理由取得":
                    dt = Com.GetDB("select * from dbo.k固定控除_入社入力");
                    break;
                case "【情報】固定控除_異動_理由取得":
                    dt = Com.GetDB("select * from dbo.k固定控除_異動入力");
                    break;
                case "【情報】昇格処理状況チェック":
                    dt = Com.GetDB("select * from dbo.s昇格処理状況チェック order by 年月, 組織CD");
                    break;

            }


            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            dataGridView1.DataSource = dt;

            DispChange();

            //全て入力した後に列幅を自動調節する
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            comboBox1.Enabled = true;

            Com.InHistory(comboBox1.SelectedItem.ToString(), "", "");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataRow row in Program.acdt.Rows)
            {
                if (row["名前"].ToString() == comboBox2.SelectedItem.ToString())
                {
                    Program.loginID = row["ID"].ToString();
                    Program.access = row["権限"].ToString();
                    //Program.logintiku = row["地区"].ToString();
                    Program.loginbusyo = row["部署"].ToString();
                    Program.loginname = row["名前"].ToString();
                    Program.dispZinzi = row["人事検索権限"].Equals(DBNull.Value) ? 0 : (int)row["人事検索権限"];
                    //1  非生産部門
                    //2
                    //3　PPP/PFI
                    //4  現業・客室
                    //5  施設・エンジ
                    //6
                    //7  八重山
                    //8  北部支店
                    //9  多面展開
                    //10 宮古島支店
                    //11 久米島支店
                    //99 フル権限(総務関係、役員)
                    Program.yakusyokucd = row["役職CD"].ToString();
                }
            }
        }

        private void DispChange()
        {
            if (dt.Rows.Count == 0) return;

            if (comboBox1.SelectedItem.ToString() == "【エラー】免許・登録・扶養・通信・離島手当不一致")
            {
                dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox1.SelectedItem.ToString() == "【情報】免許手当例外一覧")
            {
                dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
            }
            else if (comboBox1.SelectedItem.ToString() == "【情報】免許(資格)・登録手当一覧")
            {
                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";
            }
            //else if (comboBox1.SelectedItem.ToString() == "【情報】有給管理一覧(前回付与10日以上で年5日未取得の方)")
            //{
            //    dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
            //    dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";

            //    dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //    dataGridView1.Columns["reskey"].Visible = false;
            //}
            else if (comboBox1.SelectedItem.ToString() == "【情報】有休残日数一覧")
            {
                dataGridView1.Columns[12].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[13].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[14].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[15].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[16].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[17].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[18].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[19].DefaultCellStyle.Format = "0\'%\'";
                dataGridView1.Columns[20].DefaultCellStyle.Format = "#,##0.#";
                //dataGridView1.Columns[21].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[22].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[23].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[24].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[25].DefaultCellStyle.Format = "#,##0.#";
                dataGridView1.Columns[26].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView1.Columns[21].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[24].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                if (Program.loginname == "石川　尚吾" || Program.loginname == "喜屋武　大祐")
                {
                }
                else
                {
                    dataGridView1.Columns["時給"].Visible = false;
                    dataGridView1.Columns["勤務時間"].Visible = false;
                    dataGridView1.Columns["残数金額換算"].Visible = false;
                }

                dataGridView1.Columns["reskey"].Visible = false;
            }
            //else if (comboBox1.SelectedItem.ToString() == "固定控除一覧表示")
            //{
            //    dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            //    dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
            //    dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //}
            else if (comboBox1.SelectedItem.ToString() == "【情報】通信手当一覧")
            {
                dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox1.SelectedItem.ToString() == "【情報】休業手当")
            {
                dataGridView1.Columns[08].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[09].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[11].DefaultCellStyle.Format = "#,0";
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
                dataGridView1.Columns[08].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[09].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
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
            }
            //else if (comboBox1.SelectedItem.ToString() == "【情報】健康診断受診状況一覧")
            //{
            //    for (int i = 0; i <= 9; i++)
            //    {
            //        dataGridView1.Columns[i].Width = 130;
            //        dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
            //        dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    }

            //    dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.GreenYellow;
            //    dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.GreenYellow;
            //    dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.GreenYellow;

            //    dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.Aqua;
            //    dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.Aqua;
            //    dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.Aqua;

            //    dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.Cornsilk;
            //    dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.Cornsilk;
            //    dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.Cornsilk;

            //    dataGridView1.Columns[3].DefaultCellStyle.Format = "N1";
            //    dataGridView1.Columns[6].DefaultCellStyle.Format = "N1";
            //    dataGridView1.Columns[9].DefaultCellStyle.Format = "N1";

            //}
            else if (comboBox1.SelectedItem.ToString() == "【情報】回数1項目一覧")
            {
                    dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (comboBox1.SelectedItem.ToString() == "【情報】緊急連絡先入力状況一覧")
            {
                //TODO 
                dataGridView1.Columns[1].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[2].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[3].DefaultCellStyle.Format = "0.0\'%\'";
                
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }




        }

        //private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //{
        //}
        //CellFromatting からCellPaintingへ変更
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (comboBox1.SelectedItem.ToString() == "【警告】単年契約終了日　間近順")
            {
                //セルの列を確認
                decimal val = 0;
                //if (e.Value != null && e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                if (e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val <= 0 || e.Value.ToString() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
                else if (e.ColumnIndex == 0 && e.Value.ToString() == "")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
            else if (comboBox1.SelectedItem.ToString() == "【警告】資格有効期限　間近順")
            {
                //セルの列を確認
                decimal val = 0;
                //if (e.Value != null && e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                if (e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val <= 0 || e.Value.ToString() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    }
                    else if (val <= 60 || e.Value.ToString() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
                else if (e.ColumnIndex == 0 && e.Value.ToString() == "")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
            else if (comboBox1.SelectedItem.ToString() == "【警告】在留カード有効期限　間近順")
            {
                //セルの列を確認
                decimal val = 0;
                //if (e.Value != null && e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                if (e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val <= 0 || e.Value.ToString() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    }
                    else if (val <= 60 || e.Value.ToString() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
                else if (e.ColumnIndex == 0 && e.Value.ToString() == "")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
            else if (comboBox1.SelectedItem.ToString() == "【情報】有休残日数一覧")
            {
                //セルの列を確認
                decimal val = 0;
                if (e.Value != null && e.ColumnIndex == 22 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val >= 0)
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
            }
            else if (comboBox1.SelectedItem.ToString() == "【警告】試用期間中従業員")
            {
                //セルの列を確認
                decimal val = 0;
                if (e.Value != null && e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val >= 180)
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
            }           
            else if (comboBox1.SelectedItem.ToString() == "【警告】事業登録 有効期限間近順")
            {
                //セルの列を確認
                decimal val = 0;
                if (e.Value != null && e.ColumnIndex == 0 && decimal.TryParse(e.Value.ToString(), out val))
                {
                    //セルの値により、背景色を変更する
                    if (val <= 0)
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    }
                }

                //if (e.Value != null && e.ColumnIndex == 4 && decimal.TryParse(e.Value.ToString(), out val))
                if (e.ColumnIndex == 4)
                {
                    //セルの値により、背景色を変更する
                    if (e.Value.ToString() != "" && e.Value.ToString() != "在籍状況")
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
        }



        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //初期処理エラー対応
            if (flg)
            {
                flg = false;
                return;
            }

            GetComboData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                GetComboData();
            }
        }
    }
}
