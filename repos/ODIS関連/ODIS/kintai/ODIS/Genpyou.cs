using C1.C1Excel;
using Microsoft.VisualBasic;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;



namespace ODIS.ODIS
{
    public partial class Genpyou : Form
    {
        private string nl = Environment.NewLine;

        private System.Data.SqlTypes.SqlString SqlStrNull = System.Data.SqlTypes.SqlString.Null;
        private System.Data.SqlTypes.SqlDecimal SqlDecNull = System.Data.SqlTypes.SqlDecimal.Null;


        private DataTable soshikidt = new DataTable();
        private DataTable syain = new DataTable();
        private DataTable exdt = new DataTable();
        private DataTable SelectDisp = new DataTable();

        private bool soshikiflg = false;
        private bool genbaflg = false;

        DataTable Keiken = new DataTable(); //社外経験
        DataTable Gaku = new DataTable();   //学歴
        DataTable Syoku = new DataTable();  //職種
        DataTable Nen = new DataTable();    //年齢

        //選択行
        private int dgvRow = 0;

        //null対応
        private DateTime zerodt = new System.DateTime(2022, 12, 31, 0, 0, 0, 0);

        //最賃額
        private decimal saichin = 0;

        public Genpyou()
        {

            if (Convert.ToInt16(Program.access) == 1)
            {
                MessageBox.Show("入力権限がありません。");
                Com.InHistory("31_異動入力権限無", "", "");
                return;
            }

            InitializeComponent();

            //初期設定
            IniSet();

            //異動データ一覧取得
            GetIdouData();

            Com.InHistory("31_異動入力", "", "");
        }

        public Genpyou(string str)
        {
            if (Convert.ToInt16(Program.access) == 1)
            {
                MessageBox.Show("入力権限がありません。");
                Com.InHistory("31_異動入力権限無", "", "");
                return;
            }

            InitializeComponent();

            //初期設定
            IniSet();

            //異動日設定
            string[] receiveIdouD = SelectIdouDay.ShowMiniForm("");　//Form2を開く
            idoudaynew.Value = Convert.ToDateTime(receiveIdouD[0]);

            //異動データ一覧取得
            GetIdouData();

            //現行データ取得
            GetData(str);

            //資格・家族情報取得
            DataDispShikaku();
            DataDispKazoku();

            Com.InHistory("31_異動入力(従業員検索経由)", "", "");
        }

        //異動データ取得
        private void GetIdouData()
        {
            DataTable dt = new DataTable();
            string sql = "select * from dbo.i異動データ取得 where ";

            if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
            {
                sql = sql + "(担当区分  in ('03_施設', '04_エンジ') or (担当区分 in ('14_宮古島','15_久米島') and 担当事務 in ('03_施設', '04_エンジ')) or 更新者 in ('" + Program.loginname + "'))";
            }
            else if (Program.loginname == "宮城　一禎")
            {
                sql = sql + "(担当区分  in ('04_エンジ') or (担当区分 in ('15_久米島')) or 更新者 in ('" + Program.loginname + "'))";
            }
            else if (Program.loginname == "金城　智之" || Program.loginname == "佐久川　昌佳" || Program.loginname == "太田　朋宏")
            {
                sql = sql + "担当区分  like '%%' ";
            }
            //TODO 2503大濱さん宮古島応援のため
            else if (Program.loginname == "大浜　綾希子")
            {
                sql = sql + "(担当区分  in ('01_現業') or (担当区分 in ('15_久米島') and 担当事務 = '01_現業') or 更新者 in ('" + Program.loginname + "'))";
            }
            else if (Program.loginname == "佐久間　みどり")
            {
                sql = sql + "(担当区分  in ('02_客室','14_宮古島') or 更新者 in ('" + Program.loginname + "'))";
            }
            else
            {
                sql = sql + "(更新者 in ('" + Program.loginname + "') or 担当区分  like '%" + Program.loginbusyo + "%') ";
            }

            //年月絞り込み
            sql = sql + " and 異動年月日 like '" + ymlist.SelectedItem + "%'";
            //sql = sql + " order by 更新日時";
            sql = sql + " order by No";
            dt = Com.GetDB(sql);

            dispdgv.DataSource = dt;
        }

        //新規ボタン
        private void button3_Click(object sender, EventArgs e)
        {
            //異動日設定
            string[] receiveIdouD = SelectIdouDay.ShowMiniForm("");　//Form2を開く
            idoudaynew.Value = Convert.ToDateTime(receiveIdouD[0]);

            //Form2に送るテキスト
            string sendText = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectEmpNyu.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //入力フォームクリア
            DataReset();

            //TODO リセットされたので再度異動日を設定
            idoudaynew.Value = Convert.ToDateTime(receiveIdouD[0]);

            //現行データ取得
            GetData(receiveText[0]);

            //資格・家族情報取得
            DataDispShikaku();
            DataDispKazoku();

            CheckNoDay(receiveText[0], "");
        }

        private void CheckNoDay(string num, string no)
        {

            //同一人物、同一異動日がないかチェック
            string sql = "select * from dbo.i異動データ取得 where 社員番号 = '" + num + "' and 異動年月日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' and No <> '" + no + "' ";
            DataTable dt = Com.GetDB(sql);

            if (dt.Rows.Count == 1)
            {
                string number = dt.Rows[0][0].ToString();
                string date = dt.Rows[0][1].ToString();
                string name = dt.Rows[0][3].ToString();

                MessageBox.Show(name + "様は、異動日：" + date + "で既に登録されています。No." + number + nl + "別異動日での登録内容であれば、異動日の変更をしてください。" + nl + "同一異動日での登録内容であれば、No:" + number + "を選択、更新してください。");
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }
        }

        private void GetHatsurei()
        {
            string kyuu = kyuuyo.SelectedItem?.ToString().Substring(0, 1); //給与支給区分
            string kyuu_m = "";
            if (kyuuyo_m.Text.Length > 0)
            {
                kyuu_m = kyuuyo_m.Text?.ToString().Substring(0, 1); //給与支給区分_前
            }
            else
            {
                kyuu_m = kyuuyo_m.Text?.ToString(); //給与支給区分_前
            }

            //所属と現場
            if (soshikiflg)
            {
                if (genbaflg)
                {
                    hatsurei.SelectedItem = "04　所属異動・現場異動";
                }
                else
                {
                    hatsurei.SelectedItem = "02　所属異動";
                }

            }
            else if (genbaflg)
            {
                hatsurei.SelectedItem = "03　現場異動";
            }
            else
            {
                //TODO [01　給与変更等（人事発令なし）]でもいいかも
                hatsurei.SelectedItem = "01　給与変更等（人事発令なし）";
                //hatsurei.SelectedIndex = -1;
            }

            //役職
            if (yakusyoku.SelectedItem?.ToString() == null || yakusyoku_m.Text.ToString() == "-")
            {

            }
            else if (yakusyoku.SelectedItem?.ToString() == yakusyoku_m.Text.ToString())
            {

            }
            else
            {
                if (Convert.ToInt16(yakusyoku.SelectedItem?.ToString().Substring(0, 4)) == Convert.ToInt16(yakusyoku_m.Text.ToString().Substring(0, 4)))
                {
                    //TODO [01　給与変更等（人事発令なし）]でもいいかも
                    hatsurei.SelectedIndex = -1;
                }
                else if (Convert.ToInt16(yakusyoku.SelectedItem?.ToString().Substring(0, 4)) < Convert.ToInt16(yakusyoku_m.Text.ToString().Substring(0, 4)))
                {
                    hatsurei.SelectedItem = "05　昇進昇給（役職変更）";
                }
                else
                {
                    hatsurei.SelectedItem = "06　降格降給（役職変更）";
                }
            }

            //試用期間で月給者を選択
            //if (shiyou_m.Text == "試用期間" && kyuu == "C" && kyuu_m != "C")
            //{
            //    hatsurei.SelectedItem = "07　本採用（試用→正社員）";
            //    //TODO
            //    return;
            //}

            //試用期間で月給者を選択
            if (shiyou_m.Text == "01　試用期間" && shiyou.Text == "")
            {
                hatsurei.SelectedItem = "07　本採用（試用→正社員）";
                return;
            }


            //雇用体系変更
            if (kyuuyo.Text == kyuuyo_m.Text)
            {

            }
            else
            {
                if (kyuu_m == "E" && (kyuu == "C" || kyuu == "D"))
                {
                    hatsurei.SelectedItem = "10　雇用体系変更（パート→正社員）";
                }
                else if (kyuu_m == "F" && (kyuu == "C" || kyuu == "D"))
                {
                    hatsurei.SelectedItem = "11　雇用体系変更（アルバイト→正社員）";
                }
                else if (kyuu_m == "F" && kyuu == "E")
                {
                    hatsurei.SelectedItem = "12　雇用体系変更（アルバイト→パート）";
                }
                else if (kyuu_m == "E" && kyuu == "F")
                {
                    hatsurei.SelectedItem = "13　雇用体系変更（パート→アルバイト）";
                }
                else if ((kyuu_m == "C" || kyuu_m == "D") && kyuu == "E")
                {
                    hatsurei.SelectedItem = "14　雇用体系変更（正社員→パート）";
                }
                else if ((kyuu_m == "C" || kyuu_m == "D") && kyuu == "E")
                {
                    hatsurei.SelectedItem = "15　雇用体系変更（正社員→アルバイト）";
                }
                else if ((kyuu_m == "D") && kyuu == "C" && shiyou_m.Text != "01　試用期間")
                {
                    hatsurei.SelectedItem = "09　雇用体系変更（日給→月給）";
                }
                else
                {
                    //hatsurei.SelectedItem = "07　本採用（試用→正社員）";
                }
            }

            //定年功労雇用
            if (keiyaku.SelectedItem?.ToString() != keiyaku_m.Text)
            {
                if (keiyaku.SelectedItem?.ToString() == "")
                {
                    hatsurei.SelectedItem = "20　定年功労雇用";
                }
            }

        }

        private void IniSet()
        {
            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dispdgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            shikakudgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            kazokudgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //選択モードを行単位での選択のみにする
            dispdgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            shikakudgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            kazokudgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //ソート不可対応
            foreach (DataGridViewColumn c in dispdgv.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            foreach (DataGridViewColumn c in shikakudgv.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            foreach (DataGridViewColumn c in kazokudgv.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            //地区コード
            tiku.Items.Add("1　本社");
            tiku.Items.Add("2　那覇");
            tiku.Items.Add("3　八重山");
            tiku.Items.Add("4　北部");
            tiku.Items.Add("5　広域");
            tiku.Items.Add("6　宮古島");
            tiku.Items.Add("7　久米島");

            //発令区分
            hatsurei.Items.Add("01　給与変更等（人事発令なし）");
            hatsurei.Items.Add("02　所属異動");
            hatsurei.Items.Add("03　現場異動");
            hatsurei.Items.Add("04　所属異動・現場異動");
            hatsurei.Items.Add("05　昇進昇給（役職変更）");
            hatsurei.Items.Add("06　降格降給（役職変更）");

            //
            hatsurei.Items.Add("07　本採用（試用→正社員）");
            hatsurei.Items.Add("09　雇用体系変更（日給→月給）");
            hatsurei.Items.Add("10　雇用体系変更（パート→正社員）");
            hatsurei.Items.Add("11　雇用体系変更（アルバイト→正社員）");
            hatsurei.Items.Add("12　雇用体系変更（アルバイト→パート）");
            hatsurei.Items.Add("13　雇用体系変更（パート→アルバイト）");
            hatsurei.Items.Add("14　雇用体系変更（正社員→パート）");
            hatsurei.Items.Add("15　雇用体系変更（正社員→アルバイト）");

            hatsurei.Items.Add("20　定年功労雇用");
            hatsurei.Items.Add("21　契約雇用更新");
            hatsurei.Items.Add("22　その他変更");

            //職種
            Syoku = Com.GetDB("select * from dbo.K_職務給_職種");
            comboBoxSyoku.Items.Add("");
            foreach (DataRow row in Syoku.Rows)
            {
                comboBoxSyoku.Items.Add(row["備考"]);
            }

            //社外経験
            Keiken = Com.GetDB("select * from dbo.K_技能給_B社外経験");
            comboBoxKeiken.Items.Add("");
            foreach (DataRow row in Keiken.Rows)
            {
                comboBoxKeiken.Items.Add(row["備考"]);
            }
            comboBoxKeiken.Items.Add("【基準外】");

            //最終学歴
            Gaku = Com.GetDB("select * from dbo.K_技能給_C最終学歴");
            comboBoxGaku.Items.Add("");
            foreach (DataRow row in Gaku.Rows)
            {
                comboBoxGaku.Items.Add(row["備考"]);
            }
            comboBoxGaku.Items.Add("【基準外】");


            //年齢
            Nen = Com.GetDB("select * from dbo.K_技能給_A年齢");


            //家族生年月日
            for (int i = 2; i < 64; i++)
            {
                warekicb.Items.Add("昭和" + i + "年");
            }

            warekicb.Items.Add("平成元年");

            for (int i = 2; i < 31; i++)
            {
                warekicb.Items.Add("平成" + i + "年");
            }

            warekicb.Items.Add("令和元年");

            for (int i = 2; i < 8; i++)
            {
                warekicb.Items.Add("令和" + i + "年");
            }

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable hatsu = new DataTable();
            DataTable yaku = new DataTable();
            DataTable toukyuu = new DataTable();
            DataTable goubou = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        //TODO 今日の日付で一旦設定する　本来はいらない
                        //string d = Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd");
                        string d = DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd");


                        //社員情報取得
                        Cmd.CommandText = "select * from dbo.s社員基本情報_期間指定('" + d + "') a left join dbo.k固定給期間指定('" + d + "') b on a.社員番号 = b.社員番号 left join dbo.z税金情報 c on a.社員番号 = c.社員番号 left join [005_住所] d on a.社員番号 = d.社員番号 left join dbo.t通勤管理テーブル e on a.社員番号 = e.社員番号 and e.管理No = '1' where a.在籍区分 <> '9'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(syain);

                        //役職一覧
                        //Cmd.CommandText = "select 管理コード, 摘要 from QUATRO.dbo.QCMTCODED where 適用終了日='9999/12/31' AND 情報キー='SJMT092' order by ソート順";
                        //da = new SqlDataAdapter(Cmd);
                        //da.Fill(yaku);

                        ////役職一覧設定
                        //foreach (DataRow drw in yaku.Rows)
                        //{
                        //    yakusyoku.Items.Add(drw["管理コード"].ToString() + "　" + drw["摘要"].ToString());
                        //}

                        //等級一覧
                        //Cmd.CommandText = "select 等級コード, 等級名称 from QUATRO.dbo.SJMTTOKYU where 適用終了日 = '9999/12/31' and 雇用体系 = '01'";
                        //da = new SqlDataAdapter(Cmd);
                        //da.Fill(toukyuu);

                        ////等級一覧設定
                        //this.shiyou.Items.Add("");
                        //foreach (DataRow drw in toukyuu.Rows)
                        //{
                        //    this.shiyou.Items.Add(drw["等級コード"].ToString() + "　" + drw["等級名称"].ToString());
                        //}
                        ////無効化
                        //this.shiyou.Enabled = false;


                        //組織一覧
                        Cmd.CommandText = "select distinct a.組織CD, a.組織名 from dbo.担当テーブル a where 定員数 > 0";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(soshikidt);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            //給与支給区分設定
            kyuuyo.Items.Add("A1" + "　" + "役員");
            //kyuuyo.Items.Add("B1" + "　" + "兼務役員");
            kyuuyo.Items.Add("C1" + "　" + "月給者");
            //kyuuyo.Items.Add("C2" + "　" + "功労月給者");
            kyuuyo.Items.Add("D1" + "　" + "日給者");
            //kyuuyo.Items.Add("D2" + "　" + "功労日給者");
            kyuuyo.Items.Add("E1" + "　" + "パート");
            kyuuyo.Items.Add("F1" + "　" + "アルバイト");

            //役職一覧設定
            yakusyoku.Items.Add("0180" + "　" + "係員");
            yakusyoku.Items.Add("0170" + "　" + "サブチーフ");
            yakusyoku.Items.Add("0160" + "　" + "チーフ");
            yakusyoku.Items.Add("0150" + "　" + "副主任");
            yakusyoku.Items.Add("0140" + "　" + "主任");
            yakusyoku.Items.Add("0135" + "　" + "係長");
            yakusyoku.Items.Add("0132" + "　" + "技術係長");
            yakusyoku.Items.Add("0130" + "　" + "課長");
            yakusyoku.Items.Add("0122" + "　" + "技術課長");
            yakusyoku.Items.Add("0120" + "　" + "副部長");
            yakusyoku.Items.Add("0112" + "　" + "技術副部長");
            yakusyoku.Items.Add("0110" + "　" + "部長");
            yakusyoku.Items.Add("0102" + "　" + "技術部長");
            yakusyoku.Items.Add("0050" + "　" + "常務取締役");
            yakusyoku.Items.Add("0020" + "　" + "代表取締役社長");
            yakusyoku.Items.Add("0045" + "　" + "取締役相談役");

            //comboBox1.Items.Add("0");
            //comboBox1.Items.Add("50000");
            //foreach (DataRow drw in yaku.Rows)
            //{
            //    yakusyoku.Items.Add(drw["管理コード"].ToString() + "　" + drw["摘要"].ToString());
            //}

            //契約社員設定
            keiyaku.Items.Add("");
            keiyaku.Items.Add("10" + "　" + "一般契約社員");
            keiyaku.Items.Add("20" + "　" + "単年契約社員");
            keiyaku.Items.Add("30" + "　" + "技能実習生");
            keiyaku.Items.Add("31" + "　" + "特技能実習生");

            //友の会区分
            tomokubun.Items.Add("");
            tomokubun.Items.Add("1　非加入");
            tomokubun.Items.Add("2　アルバイト加入");
            //tomokubun.SelectedIndex = 0;

            //社員区分
            //kyuuzitsukubun.Items.Add("01" + "　" + "月給者");
            //kyuuzitsukubun.Items.Add("02" + "　" + "日給者");
            //kyuuzitsukubun.Items.Add("03" + "　" + "パート");
            //kyuuzitsukubun.Items.Add("04" + "　" + "アルバイト");
            //kyuuzitsukubun.Enabled = false;

            //休日区分
            kyuuzitsukubun.Items.Add("10" + "　" + "年間最低数");
            kyuuzitsukubun.Items.Add("20" + "　" + "土日祝");

            //試用期間
            this.shiyou.Items.Add("");
            this.shiyou.Items.Add("01" + "　" + "試用期間");

            //休暇付与区分
            kyuuka.Items.Add("0" + "　" + "5日以上");
            kyuuka.Items.Add("1" + "　" + "4日");
            kyuuka.Items.Add("2" + "　" + "3日");
            kyuuka.Items.Add("3" + "　" + "2日");
            kyuuka.Items.Add("4" + "　" + "1日");
            kyuuka.Items.Add("9" + "　" + "付与なし");

            //勤務時間
            kinmu.Items.Add("8");
            kinmu.Items.Add("7");
            kinmu.Items.Add("6");
            kinmu.Items.Add("5");
            kinmu.Items.Add("4");
            kinmu.Items.Add("3");
            kinmu.Items.Add("2");
            kinmu.Items.Add("1");

            //税表区分
            zeikubun.Items.Add("1　甲");
            zeikubun.Items.Add("2　乙");
            //zeikubun.Items.Add("3非居住");

            //障害区分
            syougai.Items.Add("");
            syougai.Items.Add("1　普通");
            syougai.Items.Add("2　特別");

            //寡フ区分
            kahu.Items.Add("");
            kahu.Items.Add("1　寡フ");
            kahu.Items.Add("2　ひとり親");

            //勤労　外国人　災害
            kinrou.Items.Add("");
            kinrou.Items.Add("1　○");
            gaikoku.Items.Add("");
            gaikoku.Items.Add("1　○");
            saigai.Items.Add("");
            saigai.Items.Add("1　○");

            //通勤手当区分
            tuukinteatekubun.Items.Add("");
            tuukinteatekubun.Items.Add("1 実費精算");

            //通勤区分
            tuukinkubun.Items.Add("1 車");
            tuukinkubun.Items.Add("2 バイク");
            //tuukinkubun.Items.Add("3 徒歩・自転車");
            tuukinkubun.Items.Add("4 バス・モノレール");
            tuukinkubun.Items.Add("5 送迎(会社)");
            tuukinkubun.Items.Add("6 送迎(知人・親族)");
            tuukinkubun.Items.Add("7 業務車両");
            tuukinkubun.Items.Add("8 徒歩");
            tuukinkubun.Items.Add("9 自転車");
            tuukinkubun.SelectedIndex = 0;
            

            //家族タブ
            //続柄区分
            DataTable zokugaradt = new DataTable();
            zokugaradt = Com.GetDB("select 管理コード + '　' + 摘要 as 続柄 from QUATRO.dbo.QCMTCODED where 情報キー = 'SJMT030' and 適用終了日 = '9999/12/31' and 管理コード <> '0'");
            foreach (DataRow drw in zokugaradt.Rows)
            {
                zokugara.Items.Add(drw["続柄"].ToString());
            }

            //資格
            //DataTable shikakudt = new DataTable();
            //shikakudt = Com.GetDB("select 管理コード +'　' + 摘要 as 資格 from QUATRO.dbo.QCMTCODED c where c.情報キー = 'SJMT095' and c.適用終了日 = '9999/12/31'");

            //foreach (DataRow drw in shikakudt.Rows)
            //{
            //    shikakucombo.Items.Add(drw["資格"].ToString());
            //}


            //同居区分
            doukyokubun.Items.Add("0　該当しない");
            doukyokubun.Items.Add("1　該当する");
            doukyokubun.SelectedIndex = 1;

            //税扶養区分
            huyoukubun.Items.Add("0　該当しない");
            huyoukubun.Items.Add("1　該当する");
            huyoukubun.SelectedIndex = 0;

            //源泉区分
            gensenkubun.Items.Add("0　該当しない");
            gensenkubun.Items.Add("1　該当する");
            gensenkubun.SelectedIndex = 0;

            //健保区分
            kenpokanyuu.Items.Add("0　該当しない");
            kenpokanyuu.Items.Add("1　該当する");
            kenpokanyuu.SelectedIndex = 0;

            //世帯主区分
            setainushi.Items.Add("0　該当しない");
            setainushi.Items.Add("1　該当する");
            setainushi.SelectedIndex = 0;

            //障害者
            syougaikubun.Items.Add("0　該当しない");
            syougaikubun.Items.Add("1　普通");
            syougaikubun.Items.Add("2　特別");
            syougaikubun.SelectedIndex = 0;

            //時給
            //zikyuu.Items.Add("0");
            //for (int i = 820; i <= 1500; i++)
            //{
            //    zikyuu.Items.Add((i).ToString());
            //}

            //日給
            //nikkyuu.Items.Add("0");
            //for (int i = 6560; i <= 15000; i++)
            //{
            //    nikkyuu.Items.Add((i * 1).ToString());
            //}

            //通信手当
            tuushin.Items.Add("0");
            tuushin.Items.Add("2000");

            //車両手当
            syaryou.Items.Add("0");
            syaryou.Items.Add("5000");

            //for (int i = 612; i <= 1500; i++)
            //{
            //    nikkyuu.Items.Add((i * 10).ToString());
            //}

            //TODO 絞り込み設定自動化
            ymlist.Items.Add("2020/04");
            ymlist.Items.Add("2020/05");
            ymlist.Items.Add("2020/06");
            ymlist.Items.Add("2020/07");
            ymlist.Items.Add("2020/08");
            ymlist.Items.Add("2020/09");
            ymlist.Items.Add("2020/10");
            ymlist.Items.Add("2020/11");
            ymlist.Items.Add("2020/12");
            ymlist.Items.Add("2021/01");
            ymlist.Items.Add("2021/02");
            ymlist.Items.Add("2021/03");
            ymlist.Items.Add("2021/04");
            ymlist.Items.Add("2021/05");
            ymlist.Items.Add("2021/06");
            ymlist.Items.Add("2021/07");
            ymlist.Items.Add("2021/08");
            ymlist.Items.Add("2021/09");
            ymlist.Items.Add("2021/10");
            ymlist.Items.Add("2021/11");
            ymlist.Items.Add("2021/12");
            ymlist.Items.Add("2022/01");
            ymlist.Items.Add("2022/02");
            ymlist.Items.Add("2022/03");
            ymlist.Items.Add("2022/04");
            ymlist.Items.Add("2022/05");
            ymlist.Items.Add("2022/06");
            ymlist.Items.Add("2022/07");
            ymlist.Items.Add("2022/08");
            ymlist.Items.Add("2022/09");
            ymlist.Items.Add("2022/10");
            ymlist.Items.Add("2022/11");
            ymlist.Items.Add("2022/12");
            ymlist.Items.Add("2023/01");
            ymlist.Items.Add("2023/02");
            ymlist.Items.Add("2023/03");
            ymlist.Items.Add("2023/04");
            ymlist.Items.Add("2023/05");
            ymlist.Items.Add("2023/06");
            ymlist.Items.Add("2023/07");
            ymlist.Items.Add("2023/08");
            ymlist.Items.Add("2023/09");
            ymlist.Items.Add("2023/10");
            ymlist.Items.Add("2023/11");
            ymlist.Items.Add("2023/12");
            ymlist.Items.Add("2024/01");
            ymlist.Items.Add("2024/02");
            ymlist.Items.Add("2024/03");
            ymlist.Items.Add("2024/04");
            ymlist.Items.Add("2024/05");
            ymlist.Items.Add("2024/06");
            ymlist.Items.Add("2024/07");
            ymlist.Items.Add("2024/08");
            ymlist.Items.Add("2024/09");
            ymlist.Items.Add("2024/10");
            ymlist.Items.Add("2024/11");
            ymlist.Items.Add("2024/12");
            ymlist.Items.Add("2025/01");
            ymlist.Items.Add("2025/02");
            ymlist.Items.Add("2025/03");
            ymlist.Items.Add("2025/04");
            ymlist.Items.Add("2025/05");
            ymlist.Items.Add("2025/06");
            ymlist.Items.Add("2025/07");
            ymlist.Items.Add("2025/08");
            ymlist.Items.Add("2025/09");
            ymlist.Items.Add("2025/10");
            ymlist.Items.Add("2025/11");
            ymlist.Items.Add("2025/12");
            ymlist.Items.Add("2026/01");
            ymlist.Items.Add("2026/02");
            ymlist.Items.Add("2026/03");

            //TODO 今日の日付の5日前の年月がデフォルト表示
            //ymlist.SelectedIndex = ymlist.FindString(Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM"));
            ymlist.SelectedIndex = ymlist.FindString(DateTime.Now.AddDays(-5).ToString("yyyy/MM"));


            //労働条件
            //労働条件の設定
            koyoukubun.Items.Add("");
            koyoukubun.Items.Add("1 期間の定めあり");
            koyoukubun.Items.Add("2 期間の定めなし");

            koushinkubun.Items.Add("");
            koushinkubun.Items.Add("1 自動更新");
            koushinkubun.Items.Add("2 更新する場合があり得る");
            koushinkubun.Items.Add("3 契約の更新はしない");

            //休日勤務
            kyuusyustu.Items.Add("");
            kyuusyustu.Items.Add("1 有");
            kyuusyustu.Items.Add("2 無");

            //時間外労働
            zikangairoudou.Items.Add("");
            zikangairoudou.Items.Add("1 有");
            zikangairoudou.Items.Add("2 無");

            //夜間勤務
            yakankinmu.Items.Add("");
            yakankinmu.Items.Add("1 有");
            yakankinmu.Items.Add("2 無");

            //定年
            teinen.Items.Add("");
            teinen.Items.Add("1 該当無");
            teinen.Items.Add("2 満65才");

            //賞与
            syouyo.Items.Add("");
            syouyo.Items.Add("1 有");
            syouyo.Items.Add("2 無");

            //退職金
            taisyokukin.Items.Add("");
            taisyokukin.Items.Add("1 有");
            taisyokukin.Items.Add("2 無");

            //休日回数 
            kyuujitsukaisuu.Items.Add("");
            kyuujitsukaisuu.Items.Add("1 年間107日　月変形週労40時間  ※2月が29日の暦日の場合は108日");
            kyuujitsukaisuu.Items.Add("2 1ヶ月につき  4日～10日  (週5以上勤務)"); //週5
            kyuujitsukaisuu.Items.Add("3 1ヶ月につき 12日～15日  (週4勤務)"); //週4
            kyuujitsukaisuu.Items.Add("4 1ヶ月につき 16日～20日  (週3勤務)"); //週3
            kyuujitsukaisuu.Items.Add("5 1ヶ月につき 20日～25日  (週2勤務)"); //週2
            kyuujitsukaisuu.Items.Add("6 1ヶ月につき 23日～27日  (週1勤務)"); //週1
        }

        private DataTable genbadt = new DataTable();
        private void GetGenba()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select distinct 現場CD, 現場名 from dbo.担当テーブル where 定員数 > 0 and 組織CD = '" + soshiki.SelectedItem.ToString().Substring(0, 5) + "'";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(genbadt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
        }

        private string DataInsertUpdate()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            string no = "";

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[i異動データ登録更新_new]";

                        Cmd.Parameters.Add(new SqlParameter("No", SqlDbType.Int)); Cmd.Parameters["No"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("異動年月日", SqlDbType.Char)); Cmd.Parameters["異動年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("姓フリ", SqlDbType.VarChar)); Cmd.Parameters["姓フリ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("姓", SqlDbType.VarChar)); Cmd.Parameters["姓"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("現郵便番号", SqlDbType.VarChar)); Cmd.Parameters["現郵便番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現住所", SqlDbType.VarChar)); Cmd.Parameters["現住所"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("郵便番号", SqlDbType.VarChar)); Cmd.Parameters["郵便番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("住所", SqlDbType.VarChar)); Cmd.Parameters["住所"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("住所フラグ", SqlDbType.Char)); Cmd.Parameters["住所フラグ"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("地区", SqlDbType.Char)); Cmd.Parameters["地区"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("所属組織", SqlDbType.Char)); Cmd.Parameters["所属組織"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現場名", SqlDbType.Char)); Cmd.Parameters["現場名"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("役職", SqlDbType.Char)); Cmd.Parameters["役職"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("給与支給区分", SqlDbType.Char)); Cmd.Parameters["給与支給区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("試用期間", SqlDbType.Char)); Cmd.Parameters["試用期間"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("休日区分", SqlDbType.Char)); Cmd.Parameters["休日区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("等級", SqlDbType.Char)); Cmd.Parameters["等級"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("契約社員", SqlDbType.Char)); Cmd.Parameters["契約社員"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("友の会区分", SqlDbType.Char)); Cmd.Parameters["友の会区分"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("税表区分", SqlDbType.VarChar)); Cmd.Parameters["税表区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("本人障害", SqlDbType.VarChar)); Cmd.Parameters["本人障害"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("寡フ", SqlDbType.VarChar)); Cmd.Parameters["寡フ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("勤労学生", SqlDbType.VarChar)); Cmd.Parameters["勤労学生"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("災害", SqlDbType.VarChar)); Cmd.Parameters["災害"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("外国人", SqlDbType.VarChar)); Cmd.Parameters["外国人"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("休暇付与区分", SqlDbType.Char)); Cmd.Parameters["休暇付与区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("基本勤務時間", SqlDbType.Decimal)); Cmd.Parameters["基本勤務時間"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("時給", SqlDbType.Decimal)); Cmd.Parameters["時給"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("日給", SqlDbType.Decimal)); Cmd.Parameters["日給"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("回数1", SqlDbType.Decimal)); Cmd.Parameters["回数1"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("回数2", SqlDbType.Decimal)); Cmd.Parameters["回数2"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("本給", SqlDbType.Decimal)); Cmd.Parameters["本給"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("職務技能給", SqlDbType.Decimal)); Cmd.Parameters["職務技能給"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("調整手当", SqlDbType.Decimal)); Cmd.Parameters["調整手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("特別手当", SqlDbType.Decimal)); Cmd.Parameters["特別手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("役職手当", SqlDbType.Decimal)); Cmd.Parameters["役職手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現場手当", SqlDbType.Decimal)); Cmd.Parameters["現場手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("離島手当", SqlDbType.Decimal)); Cmd.Parameters["離島手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("免許手当", SqlDbType.Decimal)); Cmd.Parameters["免許手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("扶養手当", SqlDbType.Decimal)); Cmd.Parameters["扶養手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("転勤手当", SqlDbType.Decimal)); Cmd.Parameters["転勤手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当非", SqlDbType.Decimal)); Cmd.Parameters["通勤手当非"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当課", SqlDbType.Decimal)); Cmd.Parameters["通勤手当課"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("登録手当", SqlDbType.Decimal)); Cmd.Parameters["登録手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通信手当", SqlDbType.Decimal)); Cmd.Parameters["通信手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("車両手当", SqlDbType.Decimal)); Cmd.Parameters["車両手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("友の会", SqlDbType.Decimal)); Cmd.Parameters["友の会"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("固定1", SqlDbType.Decimal)); Cmd.Parameters["固定1"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("固定2", SqlDbType.Decimal)); Cmd.Parameters["固定2"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("職種", SqlDbType.VarChar)); Cmd.Parameters["職種"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("最終学歴", SqlDbType.VarChar)); Cmd.Parameters["最終学歴"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社外経験", SqlDbType.VarChar)); Cmd.Parameters["社外経験"].Direction = ParameterDirection.Input;

                        //Cmd.Parameters.Add(new SqlParameter("通勤手当区分", SqlDbType.VarChar)); Cmd.Parameters["通勤手当区分"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("通勤手段区分", SqlDbType.VarChar)); Cmd.Parameters["通勤手段区分"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("片道通勤距離", SqlDbType.Decimal)); Cmd.Parameters["片道通勤距離"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("片道料金", SqlDbType.Decimal)); Cmd.Parameters["片道料金"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("発令区分", SqlDbType.VarChar)); Cmd.Parameters["発令区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考1", SqlDbType.VarChar)); Cmd.Parameters["備考1"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考2", SqlDbType.VarChar)); Cmd.Parameters["備考2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考3", SqlDbType.VarChar)); Cmd.Parameters["備考3"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考4", SqlDbType.VarChar)); Cmd.Parameters["備考4"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考5", SqlDbType.VarChar)); Cmd.Parameters["備考5"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("固定控除1理由", SqlDbType.VarChar)); Cmd.Parameters["固定控除1理由"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("固定控除2理由", SqlDbType.VarChar)); Cmd.Parameters["固定控除2理由"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.DateTime)); Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar)); Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("通勤1日単価", SqlDbType.Decimal)); Cmd.Parameters["通勤1日単価"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["No"].Value = lbl_no.Text;
                        Cmd.Parameters["異動年月日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");
                        Cmd.Parameters["社員番号"].Value = this.syainno.Text;

                        //Cmd.Parameters["現郵便番号"].Value = yuubin.Text;
                        //Cmd.Parameters["現住所"].Value = zyuusyo.Text;
                        //Cmd.Parameters["姓フリ"].Value = seihuri.Text;
                        //Cmd.Parameters["姓"].Value = sei.Text;

                        Cmd.Parameters["姓フリ"].Value = seihuri.Text == seihuri_m.Text ? SqlStrNull : seihuri.Text;
                        Cmd.Parameters["姓"].Value = sei.Text == sei_m.Text ? SqlStrNull : sei.Text;

                        Cmd.Parameters["現郵便番号"].Value = yuubin.Text == yuubin_m.Text ? SqlStrNull : yuubin.Text;
                        Cmd.Parameters["現住所"].Value = zyuusyo.Text == zyuusyo_m.Text ? SqlStrNull : zyuusyo.Text;
                        Cmd.Parameters["郵便番号"].Value = yuubin2.Text == yuubin_m2.Text ? SqlStrNull : yuubin2.Text;
                        Cmd.Parameters["住所"].Value = zyuusyo2.Text == zyuusyo_m2.Text ? SqlStrNull : zyuusyo2.Text;
                        //TODO
                        Cmd.Parameters["住所フラグ"].Value = zyuusyocheck.Checked ? 1 : 0;

                        Cmd.Parameters["地区"].Value = tiku.SelectedItem?.ToString() == tiku_m.Text ? SqlStrNull : tiku.SelectedItem?.ToString();
                        Cmd.Parameters["所属組織"].Value = soshiki.SelectedItem?.ToString() == soshiki_m.Text ? SqlStrNull : soshiki.SelectedItem?.ToString();
                        //TODO 同現場で職種変更だけするとうまくいかない
                        //TODO とりあえず戻したからうまくいかない状態になっているかもしれない
                        //TODO 比較対象を組織追加20200819
                        Cmd.Parameters["現場名"].Value = genba.SelectedItem?.ToString() == genba_m.Text && soshiki.SelectedItem?.ToString() == soshiki_m.Text ? SqlStrNull : genba.SelectedItem?.ToString();
                        //Cmd.Parameters["現場名"].Value = genba.SelectedItem?.ToString();

                        Cmd.Parameters["役職"].Value = yakusyoku.SelectedItem?.ToString() == yakusyoku_m.Text ? SqlStrNull : yakusyoku.SelectedItem?.ToString();
                        Cmd.Parameters["給与支給区分"].Value = kyuuyo.SelectedItem?.ToString() == kyuuyo_m.Text ? SqlStrNull : kyuuyo.SelectedItem?.ToString();
                        Cmd.Parameters["等級"].Value = SqlStrNull;
                        //Cmd.Parameters["社員区分"].Value = kyuuzitsukubun.SelectedItem?.ToString() == syain_m.Text ? SqlStrNull : kyuuzitsukubun.SelectedItem?.ToString();
                        Cmd.Parameters["休日区分"].Value = kyuuzitsukubun.SelectedItem?.ToString() == kyuuzitsu_m.Text ? SqlStrNull : kyuuzitsukubun.SelectedItem?.ToString();

                        Cmd.Parameters["試用期間"].Value = shiyou.SelectedItem?.ToString() == shiyou_m.Text ? SqlStrNull : shiyou.SelectedItem?.ToString();
                        Cmd.Parameters["契約社員"].Value = keiyaku.SelectedItem?.ToString() == keiyaku_m.Text ? SqlStrNull : keiyaku.SelectedItem?.ToString();

                        Cmd.Parameters["友の会区分"].Value = tomokubun.SelectedItem?.ToString() == tomokubun_m.Text ? SqlStrNull : tomokubun.SelectedItem?.ToString();

                        Cmd.Parameters["税表区分"].Value = zeikubun.SelectedItem?.ToString() == zeikubun_m.Text ? SqlStrNull : zeikubun.SelectedItem?.ToString();
                        Cmd.Parameters["本人障害"].Value = syougai.SelectedItem?.ToString() == syougai_m.Text ? SqlStrNull : syougai.SelectedItem?.ToString();
                        Cmd.Parameters["寡フ"].Value = kahu.SelectedItem?.ToString() == kahu_m.Text ? SqlStrNull : kahu.SelectedItem?.ToString();
                        Cmd.Parameters["勤労学生"].Value = kinrou.SelectedItem?.ToString() == kinrou_m.Text ? SqlStrNull : kinrou.SelectedItem?.ToString();
                        Cmd.Parameters["災害"].Value = saigai.SelectedItem?.ToString() == saigai_m.Text ? SqlStrNull : saigai.SelectedItem?.ToString();
                        Cmd.Parameters["外国人"].Value = gaikoku.SelectedItem?.ToString() == gaikoku_m.Text ? SqlStrNull : gaikoku.SelectedItem?.ToString();

                        Cmd.Parameters["休暇付与区分"].Value = kyuuka.SelectedItem?.ToString() == kyuuka_m.Text ? SqlStrNull : kyuuka.SelectedItem?.ToString();


                        Cmd.Parameters["基本勤務時間"].Value = kinmu.SelectedItem?.ToString() == kinmu_m.Text ? SqlDecNull : Convert.ToDecimal(kinmu.Text);

                        if (zikyuu.Enabled)
                        {
                            Cmd.Parameters["時給"].Value = zikyuu.Value.ToString() == zikyuu_m.Text ? SqlDecNull : Convert.ToDecimal(zikyuu.Value);
                        }
                        else
                        {
                            Cmd.Parameters["時給"].Value = SqlDecNull;
                        }

                        if (nikkyuu.Enabled)
                        {
                            Cmd.Parameters["日給"].Value = nikkyuu.Value.ToString() == nikkyuu_m.Text ? SqlDecNull : Convert.ToDecimal(nikkyuu.Value);
                        }
                        else
                        {
                            Cmd.Parameters["日給"].Value = SqlDecNull;
                        }

                        Cmd.Parameters["回数1"].Value = kaisuu1.Value.ToString() == kaisuu1_m.Text ? SqlDecNull : Convert.ToDecimal(kaisuu1.Value);
                        //Cmd.Parameters["回数2"].Value = kaisuu2.Value.ToString() == kaisuu2_m.Text ? SqlDecNull : Convert.ToDecimal(kaisuu2.Value);
                        Cmd.Parameters["回数2"].Value = SqlDecNull;

                        Cmd.Parameters["本給"].Value = honkyuu.Value.ToString() == honkyuu_m.Text ? SqlDecNull : Convert.ToDecimal(honkyuu.Value.ToString());
                        Cmd.Parameters["職務技能給"].Value = syokumu.Value.ToString() == syokumu_m.Text ? SqlDecNull : Convert.ToDecimal(syokumu.Value.ToString());
                        //Cmd.Parameters["調整手当"].Value = tyousei.Value.ToString() == tyousei_m.Text ? SqlDecNull : Convert.ToDecimal(tyousei.Value);
                        Cmd.Parameters["調整手当"].Value = SqlDecNull;


                        Cmd.Parameters["特別手当"].Value = tokubetsu.Value.ToString() == tokubetsu_m.Text ? SqlDecNull : Convert.ToDecimal(tokubetsu.Value);
                        Cmd.Parameters["役職手当"].Value = yakuteate.Text.ToString() == yakuteate_m.Text ? SqlDecNull : Convert.ToDecimal(yakuteate.Text);
                        //Cmd.Parameters["現場手当"].Value = genbateate.Text.ToString() == genbateate_m.Text ? SqlDecNull : Convert.ToDecimal(genbateate.Text);

                        //TODO 2020/12/17
                        Cmd.Parameters["現場手当"].Value = SqlDecNull;
                        Cmd.Parameters["離島手当"].Value = ritou.Text.ToString() == ritou_m.Text ? SqlDecNull : Convert.ToDecimal(ritou.Text);

                        Cmd.Parameters["免許手当"].Value = menkyo.Text.ToString() == menkyo_m.Text ? SqlDecNull : Convert.ToDecimal(menkyo.Text);
                        Cmd.Parameters["扶養手当"].Value = huyou.Text.ToString() == huyou_m.Text ? SqlDecNull : Convert.ToDecimal(huyou.Text);
                        Cmd.Parameters["転勤手当"].Value = syukkou.Value.ToString() == syukkou_m.Text ? SqlDecNull : Convert.ToDecimal(syukkou.Value);

                        Cmd.Parameters["通勤手当非"].Value = tuukinhi.Text == tuukinhi_m.Text ? SqlDecNull : Convert.ToDecimal(tuukinhi.Text);
                        Cmd.Parameters["通勤手当課"].Value = tuukinka.Text == tuukinka_m.Text ? SqlDecNull : Convert.ToDecimal(tuukinka.Text);

                        Cmd.Parameters["登録手当"].Value = touroku.Text == touroku_m.Text ? SqlDecNull : Convert.ToDecimal(touroku.Text);
                        Cmd.Parameters["通信手当"].Value = tuushin.Text == tuushin_m.Text ? SqlDecNull : Convert.ToDecimal(tuushin.Text);
                        Cmd.Parameters["車両手当"].Value = syaryou.Text == syaryou_m.Text ? SqlDecNull : Convert.ToDecimal(syaryou.Text);

                        Cmd.Parameters["友の会"].Value = tomonokai.Text == tomonokai_m.Text ? SqlDecNull : Convert.ToDecimal(tomonokai.Text);
                        Cmd.Parameters["固定1"].Value = SqlDecNull;
                        Cmd.Parameters["固定2"].Value = SqlDecNull;

                        Cmd.Parameters["職種"].Value = comboBoxSyoku.SelectedItem?.ToString() == syokusyu_m.Text ? SqlStrNull : comboBoxSyoku.SelectedItem?.ToString();
                        Cmd.Parameters["最終学歴"].Value = comboBoxGaku.SelectedItem?.ToString() == gakureki_m.Text ? SqlStrNull : comboBoxGaku.SelectedItem?.ToString();
                        Cmd.Parameters["社外経験"].Value = comboBoxKeiken.SelectedItem?.ToString() == keiken_m.Text ? SqlStrNull : comboBoxKeiken.SelectedItem?.ToString();

                        //Cmd.Parameters["通勤手当区分"].Value = tuukinteatekubun.SelectedItem?.ToString() == tuukinteatekubun_m.Text ? SqlStrNull : tuukinteatekubun.SelectedItem?.ToString();
                        //Cmd.Parameters["通勤手段区分"].Value = tuukinkubun.SelectedItem?.ToString() == tuukinkubun_m.Text ? SqlStrNull : tuukinkubun.SelectedItem?.ToString();

                        //Cmd.Parameters["片道通勤距離"].Value = Convert.ToDecimal(katakyori.Value);
                        //Cmd.Parameters["片道通勤距離"].Value = katakyori.Value.ToString() == katakyori_m.Text ? SqlDecNull : Convert.ToDecimal(katakyori.Value);
                        //Cmd.Parameters["片道料金"].Value = SqlStrNull;


                        Cmd.Parameters["発令区分"].Value = hatsurei.Text;
                        Cmd.Parameters["備考1"].Value = bikou1.Text;
                        Cmd.Parameters["備考2"].Value = bikou2.Text;
                        Cmd.Parameters["備考3"].Value = bikou3.Text;
                        Cmd.Parameters["備考4"].Value = bikou4.Text;
                        Cmd.Parameters["備考5"].Value = tokureason.Text;
                        Cmd.Parameters["固定控除1理由"].Value = "";
                        Cmd.Parameters["固定控除2理由"].Value = "";
                        Cmd.Parameters["更新日時"].Value = DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒");
                        Cmd.Parameters["更新者"].Value = Program.loginname;
                        //Cmd.Parameters["通勤1日単価"].Value = tuutanka.Text == tuutanka_m.Text ? SqlDecNull : Convert.ToInt64(tuutanka.Text);
                        SqlDataReader dr = Cmd.ExecuteReader();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            return no;
        }


        private string DeleteInfo()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            string no = "";

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[i異動削除履歴登録]";

                        Cmd.Parameters.Add(new SqlParameter("No", SqlDbType.Int)); Cmd.Parameters["No"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("異動年月日", SqlDbType.Char)); Cmd.Parameters["異動年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("名前", SqlDbType.VarChar)); Cmd.Parameters["名前"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar)); Cmd.Parameters["備考"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.DateTime)); Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar)); Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;

                        Cmd.Parameters["No"].Value = lbl_no.Text;
                        Cmd.Parameters["異動年月日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");
                        Cmd.Parameters["社員番号"].Value = this.syainno.Text;
                        Cmd.Parameters["名前"].Value = name.Text;
                        Cmd.Parameters["備考"].Value = bikou1.Text;
                        Cmd.Parameters["更新日時"].Value = DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒");
                        Cmd.Parameters["更新者"].Value = Program.loginname;

                        SqlDataReader dr = Cmd.ExecuteReader();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            return no;
        }

        private void TourokuKoushin()
        {
            //TODO 制限処理
            //TODO 発令整理20200330
            //TOOD 下記関数呼出前も制限処理必要だよ
            if (syainno.Text == "")
            {
                MessageBox.Show("誰も選んでないす");
                return;
            }

            if (idoudaynew.Value.Equals(DBNull.Value))
            {
                MessageBox.Show("異動日は必須です。");
                return;
            }

            if (kyuuyo.Text == "D1　日給者")
            {
                if (comboBoxSyoku.SelectedItem?.ToString() == "" || comboBoxGaku.SelectedItem?.ToString() == "" || comboBoxKeiken.SelectedItem?.ToString() == "")
                { 
                    MessageBox.Show("職種、最終学歴、社外経験を選択してください。");
                    return;
                }
            }

            if (tokubetsu.Value > 0 && tokureason.Text.Length == 0)
            {
                MessageBox.Show("特別手当支給場合は、特別手当理由は必須です。");
                return;
            }

            //if (kotei1.Value > 0 && koteik1reason.Text.Length == 0)
            //{
            //    MessageBox.Show("固定控除1変更場合は、固定控除1変更理由は必須です。");
            //    return;
            //}
            
            //if (kotei2.Value > 0 && koteik2reason.Text.Length == 0)
            //{
            //    MessageBox.Show("固定控除2変更場合は、固定控除2変更理由は必須です。");
            //    return;
            //}


            GetHatsurei();

            if (hatsurei.Text == "")
            {
                hatsurei.SelectedIndex = 0;
            }

            DateTime idoud = Convert.ToDateTime(idoudaynew.Value);

            //異動日が変化しているかチェック
            DataTable dtetc = new DataTable();
            dtetc = Com.GetDB("select 異動年月日 from dbo.i異動データ where No = '" + lbl_no.Text + "'");

            if (dtetc.Rows.Count > 0)
            {
                if (dtetc.Rows[0][0].ToString() != Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd"))
                {
                    //異動家族データがあれば、適用開始日も変更!
                    DataTable dt = new DataTable();
                    string sql = "update QUATRO.dbo.SJMTKAZOKU set 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' where 会社コード = 'E0' and ユーザ任意フィールド１ is not null and 社員番号 = '" + syainno.Text + "' and 適用開始日 = '" + dtetc.Rows[0][0].ToString() + "'";
                    dt = Com.GetDB(sql);

                    //異動家族データがあれば、適用終了日も変更!
                    DataTable dt3 = new DataTable();
                    string sql3 = "update QUATRO.dbo.SJMTKAZOKU set 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "' where 会社コード = 'E0' and ユーザ任意フィールド１ is not null and 社員番号 = '" + syainno.Text + "' and 適用終了日 = '" + Convert.ToDateTime(dtetc.Rows[0][0]).AddDays(-1).ToString("yyyy/MM/dd") + "'";
                    dt3 = Com.GetDB(sql3);


                    //異動資格データがあれば、適用開始日も変更!
                    string sql2 = "update QUATRO.dbo.SJMTSHIKAK set 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' where 会社コード = 'E0' and 個人識別ＩＤ = '' and 社員番号 = '" + syainno.Text + "' and 適用開始日 = (select 異動年月日 from dbo.i異動データ where No = " + lbl_no.Text + ")";
                    dt = Com.GetDB(sql2);

                    //通勤手当テーブルデータがあれば、適用開始日も変更!
                    DataTable dt4 = new DataTable();
                    string sql4 = "update dbo.t通勤手当元データ set 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' where 社員番号 = '" + syainno.Text + "' and 適用開始日 = '" + dtetc.Rows[0][0].ToString() + "'";
                    dt4 = Com.GetDB(sql4);

                    //通勤手当テーブルデータがあれば、適用終了日も変更!
                    DataTable dt5 = new DataTable();
                    string sql5 = "update dbo.t通勤手当元データ set 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "' where 社員番号 = '" + syainno.Text + "' and 適用終了日 = '" + Convert.ToDateTime(dtetc.Rows[0][0]).AddDays(-1).ToString("yyyy/MM/dd") + "'";
                    dt5 = Com.GetDB(sql5);
                }
            }

            DataTable dtmae = new DataTable();
            DataTable dtato = new DataTable();

            dtmae = Com.GetDB("select count(*) from dbo.i異動データ");

            string no = DataInsertUpdate();

            //通勤手当元データ登録更新
            SetTuukinMotoData();

            //労働条件登録更新
            //TODO エラーになるからとりあえずコメントアウト！
            //UpdateRoudou();
            dtato = Com.GetDB("select count(*) from dbo.i異動データ");



            //登録した異動日に合わせる
            ymlist.SelectedIndex = ymlist.FindString(Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM"));


            //一覧表示は最終更新順！必ず最後の行に入らなければならない！
            //dgvRow = dispdgv.Rows.Count - 1;

            //dispdgv.CurrentCell = dispdgv[1, dgvRow];

            //登録後のフォーカスに利用
            //TODO 一覧表示は登録順または追加は必ず最後の行に入らなければならない！
            //if (dispdgv.CurrentCell == null)
            if (dtmae.Rows[0][0].ToString() != dtato.Rows[0][0].ToString())
            {
                //新規追加場合
                dgvRow = dispdgv.Rows.Count;
            }
            else
            {
                dgvRow = dispdgv.CurrentCell.RowIndex;
            }

            //異動一覧データ取得
            GetIdouData();


            //何のためにやっているか不明
            if (no != "")
            {
                dgvRow = dispdgv.Rows.Count - 1;
            }

            dispdgv.CurrentCell = dispdgv[1, dgvRow];

            //入力フォームクリア
            DataReset();

            DataGridViewRow dgr = dispdgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            //異動日が一旦リセットされているので、強制的に登録直前の異動日を格納

            idoudaynew.Value = Convert.ToDateTime(idoud);

            //現行データを取得
            GetData(drv[2].ToString());

            //異動データを取得
            GetDataIdou(drv[0].ToString(), drv[2].ToString());

            if (no == "")
            {
                MessageBox.Show("更新しました。");
            }
            else
            {
                MessageBox.Show("登録しました。 社員番号は" + drv[1].ToString() + "です。");
            }
        }
        //登録
        private void button1_Click(object sender, EventArgs e)
        {
            TourokuKoushin();

        }



        private void GetData(string num)
        {
            //改訂前データ取得
            //DataRow[] dr = syain.Select("社員番号 = '" + num + "'");


            string d = Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd");

            //処理年
            string y = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy");
            //処理月
            string m = Convert.ToDateTime(idoudaynew.Value).ToString("MM");

            DataTable dt = new DataTable();

            string sql = "select * from dbo.i異動前データ取得('" + d + "', '" + num + "', '" + y + "', '" + m + "')";
            dt = Com.GetDB(sql);

            foreach (DataRow drw in dt.Rows)
            {
                //基本情報
                syainno.Text = drw["社員番号"].ToString();
                kana.Text = drw["カナ名"].ToString();
                name.Text = drw["氏名"].ToString();
                seinengappi.Text = drw["生年月日"].ToString();
                nyuusya.Text = drw["入社年月日"].ToString();
                kinzoku.Text = drw["在籍年月"].ToString();
                seibetsu.Text = drw["性別区分"].ToString() == "1" ? "男性" : "女性";
                kozinshikibetsuid.Text = drw["個人識別ID"].ToString();


                shiyou_m.Text = drw["試用"].ToString();

                //区分情報
                tiku_m.Text = drw["地区CD"].ToString() + "　" + drw["地区名"].ToString();
                soshiki_m.Text = drw["組織CD"].ToString() + "　" + drw["組織名"].ToString();
                genba_m.Text = drw["現場CD"].ToString() + "　" + drw["現場名"].ToString();
                kyuuyo_m.Text = drw["給与支給区分"].ToString() + "　" + drw["給与支給区分名"].ToString();

                syokusyu_m.Text = drw["職種"].ToString();
                gakureki_m.Text = drw["最終学歴"].ToString();
                keiken_m.Text = drw["社外経験"].ToString();
                nyuusyaold_m.Text = drw["入社時年齢"].ToString();
                hyoukakyuu_m.Text = drw["評価額"].ToString();

                if (drw["基準外"].ToString() == "")
                {
                    kizyungai_m.Text = "0";
                }
                else
                {
                    kizyungai_m.Text = drw["基準外"].ToString();
                }


                yakusyoku_m.Text = drw["役職CD"].ToString() + "　" + drw["役職名"].ToString();
                
                
                string kk = "";
                if (drw["契約社員"].ToString() == "10")
                {
                    kk = "10　一般契約社員";
                }
                else if (drw["契約社員"].ToString() == "20")
                {
                    kk = "20　単年契約社員";
                }
                else if (drw["契約社員"].ToString() == "30")
                {
                    kk = "30　技・特実習生";
                }
                else
                {

                }
                keiyaku_m.Text = kk;

                string tk = "";
                if (drw["友の会区分"].ToString() == "1")
                {
                    tk = "1　非加入";
                }
                else if (drw["友の会区分"].ToString() == "2")
                {
                    tk = "2　アルバイト加入";
                }
                else
                {

                }
                tomokubun_m.Text = tk;


                //区分固定情報
                kyuuzitsu_m.Text = drw["休日区分"].ToString() + "　" + drw["休日区分名"].ToString();
                //toukyuu_m.Text = (drw["等級"].ToString() + "　" + drw["等級名称"].ToString()).Trim();

                //氏名変更
                sei_m.Text = drw["氏名"].ToString().Split('　')[0];
                seihuri_m.Text = drw["カナ名"].ToString().Split('　')[0];

                //郵便住所
                yuubin_m.Text = drw["現郵便番号"].ToString();
                zyuusyo_m.Text = drw["現住所"].ToString();

                //住民登録先
                yuubin_m2.Text = drw["郵便番号"].ToString();
                zyuusyo_m2.Text = drw["住所"].ToString();

                //時給等
                zikyuu_m.Text = Convert.ToInt32(drw["時給"]).ToString();
                nikkyuu_m.Text = Convert.ToInt32(drw["日給"]).ToString();
                kaisuu1_m.Text = Convert.ToInt32(drw["回数1単価"]).ToString();
                //kaisuu2_m.Text = Convert.ToInt32(drw["回数2単価"]).ToString();
                kyuuka_m.Text = drw["休暇付与区分"].ToString() + "　" + drw["週労働数"].ToString();
                kinmu_m.Text = drw["勤務時間"].ToString().Substring(0, 1);

                tuukinteatekubun_m.Text = drw["通勤手当区分"].ToString();
                tuukinkubun_m.Text = drw["通勤方法"].ToString(); //通勤手段区分
                katakyori_m.Text = drw["片道距離"].ToString();
                tuutanka_m.Text = Convert.ToInt32(drw["通勤1日単価"]).ToString();

                decimal tankad = 0;
                decimal hid = 0;
                decimal kad = 0;

                Com.CalcTuukin(drw["通勤方法"].ToString(), Convert.ToDecimal(drw["片道距離"].ToString()), Convert.ToDecimal(Getrday(kyuuka_m.Text)), ref tankad, ref hid, ref kad);

                //if (Convert.ToInt32(drw["通勤1日単価"]).ToString() != Convert.ToInt32(tankad).ToString())
                //{ 
                //    MessageBox.Show("おかしー。喜屋武へ連絡ねがいます。");
                //}

                if (drw["通勤手当区分"].ToString().Trim() == "")
                {
                    tuukinhi_m.Text = Convert.ToInt64(hid).ToString();//通勤非課税
                    tuukinka_m.Text = Convert.ToInt64(kad).ToString();//通勤課税
                }
                else
                {
                    tuukinhi_m.Text = "0";//通勤非課税
                    tuukinka_m.Text = "0";//通勤課税
                }

                //税表区分
                if (drw["税表区分"].ToString() == "1")
                {
                    zeikubun_m.Text = "1　甲";
                }
                else if (drw["税表区分"].ToString() == "2")
                {
                    zeikubun_m.Text = "2　乙";
                }
                else
                {
                    zeikubun_m.Text = "非居住";
                }

                //本人障害区分
                if (drw["本人特障"].ToString() == "1")
                {
                    syougai_m.Text = "2　特別";
                }
                else if (drw["本人普障"].ToString() == "1")
                {
                    syougai_m.Text = "1　普通";
                }
                else
                {
                    syougai_m.Text = "";
                }

                //寡フ区分
                if (drw["本人ひとり親"].ToString() == "1")
                {
                    kahu_m.Text = "2　ひとり親";
                }
                else if (drw["本人寡フ"].ToString() == "1")
                {
                    kahu_m.Text = "1　寡フ";
                }
                else
                {
                    kahu_m.Text = "";
                }

                kinrou_m.Text = drw["本人勤労"].ToString() == "0" ? "" : "1　○";
                gaikoku_m.Text = drw["本人外国人"].ToString() == "0" ? "" : "1　○";
                saigai_m.Text = drw["本人災害"].ToString() == "0" ? "" : "1　○";

                //日給者・パート・アルバイトの本給に暫定給を表示



                double hon = 0;
                if (zikyuu_m.Text != "0")
                {
                    hon = Math.Round(Convert.ToInt32(zikyuu_m.Text) * Convert.ToInt32(kinmu_m.Text) * Getrday(kyuuka_m.Text));
                    honkyuu_m.Text = hon.ToString();
                    honkyuu_m.ForeColor = Color.Red;
                    //honkyuu.ForeColor = Color.Red;
                }
                else if (nikkyuu_m.Text != "0")
                {
                    hon = Math.Round(Convert.ToInt32(nikkyuu_m.Text) * Getrday(kyuuka_m.Text));
                    honkyuu_m.Text = hon.ToString();
                    honkyuu_m.ForeColor = Color.Red;
                    //honkyuu.ForeColor = Color.Red;
                }
                else
                {
                    hon = Convert.ToInt32(drw["本給"]);
                    honkyuu_m.Text = hon > 0 ? hon.ToString() : "";
                }

                int syo = Convert.ToInt32(drw["職務技能給"]);
                int tyou = Convert.ToInt32(drw["調整手当"]);
                int toku = Convert.ToInt32(drw["特別手当"]);
                int yaku = Convert.ToInt32(drw["役職手当"]);
                //int gen = Convert.ToInt32(drw["現場手当"]);
                int men = Convert.ToInt32(drw["免許手当"]);
                int rito = Convert.ToInt32(drw["離島手当"]);
                int syu = Convert.ToInt32(drw["転勤手当"]);
                int tou = Convert.ToInt32(drw["登録手当"]);
                int tuu = Convert.ToInt32(drw["通信手当"]);
                int sya = Convert.ToInt32(drw["車両手当"]);

                //特別手当の理由
                tokureason.Text = drw["特別手当内容"].ToString();

                //固定控除の理由
                //koteik1reason.Text = drw["固定控除1内容"].ToString();
                //koteik2reason.Text = drw["固定控除2内容"].ToString();

                syokumu_m.Text = syo >= 0 ? syo.ToString() : "";
                //tyousei_m.Text = tyou >= 0 ? tyou.ToString() : "";
                tokubetsu_m.Text = toku >= 0 ? toku.ToString() : "";
                yakuteate_m.Text = yaku >= 0 ? yaku.ToString() : "";
                //genbateate_m.Text = gen >= 0 ? gen.ToString() : "";
                menkyo_m.Text = men >= 0 ? men.ToString() : "";
                ritou_m.Text = rito >= 0 ? rito.ToString() : "";
                syukkou_m.Text = syu > 0 ? syu.ToString() : "0";//転勤手当
                touroku_m.Text = tou > 0 ? tou.ToString() : "0";//登録手当
                tuushin_m.Text = tuu > 0 ? tuu.ToString() : "0";//通信手当
                syaryou_m.Text = sya > 0 ? sya.ToString() : "0";//車両手当

                //現行の基準内賃金の算出
                int ki = Convert.ToInt32(hon + syo + tyou + toku + yaku + men + rito + syu + tou + tuu + sya);
                kizyun_m.Text = ki.ToString();


                int tuhi = Convert.ToInt32(drw["通勤非課税"]);
                int tuka = Convert.ToInt32(drw["通勤課税"]);
                int hu = Convert.ToInt32(drw["扶養手当"]);

                huyou_m.Text = hu >= 0 ? hu.ToString() : "";

                //syokumu_m.Text = syo > 0 ? syo.ToString() : "0"s//職務技能
                //tuukinhi_m.Text = tuhi > 0 ? tuhi.ToString() : "0";//通勤非課税
                //tuukinka_m.Text = tuka > 0 ? tuka.ToString() : "0";//通勤課税

                shikyuu_m.Text = (ki + hu + Convert.ToInt32(tuukinhi_m.Text) + Convert.ToInt32(tuukinka_m.Text)).ToString();



                int tomo = Convert.ToInt32(drw["友の会"]);
                int ko1 = Convert.ToInt32(drw["固定他1"]);
                int ko2 = Convert.ToInt32(drw["固定他2"]);

                //
                tomonokai_m.Text = tomo > 0 ? tomo.ToString() : "0";
                //kotei1_m.Text = ko1 > 0 ? ko1.ToString() : "0";
                //kotei2_m.Text = ko2 > 0 ? ko2.ToString() : "0";

                //kouzyo_m.Text = (tomo + ko1 + ko2).ToString();

                //職務技能給
                Syokumu_Sum_m_Calc();
            }



            #region 変更前データを変更後データへ反映
            tiku.SelectedItem = tiku_m.Text;　　　//地区
            soshiki.SelectedItem = soshiki_m.Text;　 //組織
            genba.SelectedItem = genba_m.Text;     //現場

            keiyaku.SelectedItem = keiyaku_m.Text;  　  //契約社員
            kyuuyo.SelectedItem = kyuuyo_m.Text;    //給与支給区分

            yakusyoku.SelectedItem = yakusyoku_m.Text; //役職

            //契約社員↑に移動する前の位置 20230104


            tomokubun.SelectedItem = tomokubun_m.Text; //友の会区分

            kyuuzitsukubun.SelectedItem = kyuuzitsu_m.Text; 　  //休日区分
            shiyou.SelectedItem = shiyou_m.Text;     //試用区分

            comboBoxSyoku.SelectedItem = syokusyu_m.Text;    //職種
            comboBoxGaku.SelectedItem = gakureki_m.Text;    //最終学歴
            comboBoxKeiken.SelectedItem = keiken_m.Text;    //社外経験
            nyuusyaold.Text = nyuusyaold_m.Text; //入社年齢

            //TODO 年齢給を入れてみる 20200409
            nennreikyuu.Text = nennreikyuu_m.Text;

            kizyungai.Text = kizyungai_m.Text; //基準外

            hyoukakyuu.Text = hyoukakyuu_m.Text; //評価

            zeikubun.Text = zeikubun_m.Text; //税表区分
            syougai.Text = syougai_m.Text; //本人障害
            kahu.Text = kahu_m.Text; //寡フ
            kinrou.Text = kinrou_m.Text; //勤労
            gaikoku.Text = gaikoku_m.Text; //外国人
            saigai.Text = saigai_m.Text; //災害

            sei.Text = sei_m.Text; //苗字
            seihuri.Text = seihuri_m.Text; //苗字カナ

            yuubin.Text = yuubin_m.Text; //郵便番号
            zyuusyo.Text = zyuusyo_m.Text; //住所

            yuubin2.Text = yuubin_m2.Text; //郵便番号
            zyuusyo2.Text = zyuusyo_m2.Text; //住所

            zikyuu.Text = zikyuu_m.Text; //時給
            nikkyuu.Text = nikkyuu_m.Text; //日給
            kaisuu1.Text = kaisuu1_m.Text; //回数1
            //kaisuu2.Text = kaisuu2_m.Text; //回数2
            kyuuka.Text = kyuuka_m.Text; //休暇付与区分
            kinmu.Text = kinmu_m.Text; //勤務時間

            tuukinteatekubun.SelectedItem = tuukinteatekubun_m.Text; //通勤手当区分
            tuukinkubun.SelectedItem = tuukinkubun_m.Text; //通勤手段区分
            
            tuutanka.Text = tuutanka_m.Text; //通勤1日単価

            katakyori.Value = Convert.ToDecimal(katakyori_m.Text); //片道距離



            //TODO コメントアウト 2022/09/27
            //honkyuu.Value = Convert.ToDecimal(honkyuu_m.Text); //本給


            syokumu.Value = Convert.ToDecimal(syokumu_m.Text); //職務技能給
            //tyousei.Value = Convert.ToDecimal(tyousei_m.Text); //調整手当
            tokubetsu.Value = Convert.ToDecimal(tokubetsu_m.Text); //特別手当
            yakuteate.Text = yakuteate_m.Text; //役職手当
            //genbateate.Text = genbateate_m.Text; //現場手当
            ritou.Text = ritou_m.Text;
            syukkou.Value = Convert.ToDecimal(syukkou_m.Text); //転勤手当　処理もれ確認2022/06/13
            //TODO コメントアウト! 2020/12/17
            //menkyo.Text = menkyo_m.Text; //免許手当
            //huyou.Text = huyou_m.Text; //扶養手当

            kizyun.Text = kizyun_m.Text; //基準内賃金

            tomonokai.Text = tomonokai_m.Text;　//友の会
            //kotei1.Text = kotei1_m.Text;　　　　//固定1
            //kotei2.Text = kotei2_m.Text;        //固定2
            #endregion

            Disp();
        }

        //異動データ取得
        private void GetDataIdou(string kanrino, string num)
        {
            //改訂後データ取得
            DataRow[] dr = syain.Select("社員番号 = '" + num + "'");

            DataTable dt = new DataTable();
            ///dt = Com.GetDB("select * from dbo.i異動データ where 異動年月日 = '" + date + "' and 社員番号 = '" + num + "'");
            dt = Com.GetDB("select * from dbo.i異動データ a left join dbo.t通勤手当元データ b on a.社員番号 = b.社員番号 and a.異動年月日 between b.適用開始日 and b.適用終了日 where No = '" + kanrino + "'");

            foreach (DataRow row in dt.Rows)
            {
                //異動先・改訂値への反映
                if (!row["No"].Equals(DBNull.Value)) lbl_no.Text = row["No"].ToString();   //No
                if (!row["異動年月日"].Equals(DBNull.Value)) idoudaynew.Value = Convert.ToDateTime(row["異動年月日"]);   //異動日

                if (!row["地区"].Equals(DBNull.Value)) tiku.SelectedIndex = tiku.FindString(row["地区"].ToString());   //地区
                if (!row["所属組織"].Equals(DBNull.Value)) soshiki.SelectedIndex = soshiki.FindString(row["所属組織"].ToString());  //組織
                if (!row["現場名"].Equals(DBNull.Value)) genba.SelectedIndex = genba.FindString(row["現場名"].ToString());     //現場
                if (!row["契約社員"].Equals(DBNull.Value)) keiyaku.SelectedIndex = keiyaku.FindString(row["契約社員"].ToString());   //契約社員
                if (!row["給与支給区分"].Equals(DBNull.Value)) kyuuyo.SelectedIndex = kyuuyo.FindString(row["給与支給区分"].ToString());    //給与支給区分
                if (!row["役職"].Equals(DBNull.Value)) yakusyoku.SelectedIndex = yakusyoku.FindString(row["役職"].ToString()); //役職

                if (!row["友の会区分"].Equals(DBNull.Value)) tomokubun.SelectedIndex = tomokubun.FindString(row["友の会区分"].ToString());   //契約社員

                if (!row["休日区分"].Equals(DBNull.Value)) kyuuzitsukubun.SelectedIndex = kyuuzitsukubun.FindString(row["休日区分"].ToString());      //社員区分
                if (!row["試用期間"].Equals(DBNull.Value)) shiyou.SelectedIndex = shiyou.FindString(row["試用期間"].ToString());   //試用期間

                if (!row["職種"].Equals(DBNull.Value)) comboBoxSyoku.SelectedIndex = comboBoxSyoku.FindString(row["職種"].ToString());    //職種
                if (!row["最終学歴"].Equals(DBNull.Value)) comboBoxGaku.SelectedIndex = comboBoxGaku.FindString(row["最終学歴"].ToString());    //最終学歴
                if (!row["社外経験"].Equals(DBNull.Value)) comboBoxKeiken.SelectedIndex = comboBoxKeiken.FindString(row["社外経験"].ToString());    //社外経験

                if (!row["時給"].Equals(DBNull.Value)) zikyuu.Value = Convert.ToDecimal(row["時給"].ToString()); //時給
                if (!row["日給"].Equals(DBNull.Value)) nikkyuu.Value = Convert.ToDecimal(row["日給"].ToString()); //日給
                if (!row["回数1"].Equals(DBNull.Value)) kaisuu1.Value = Convert.ToDecimal(row["回数1"].ToString()); //回数1
                //if (!row["回数2"].Equals(DBNull.Value)) kaisuu2.Value = Convert.ToDecimal(row["回数2"].ToString()); //回数2

                if (!row["休暇付与区分"].Equals(DBNull.Value)) kyuuka.SelectedIndex = kyuuka.FindString(row["休暇付与区分"].ToString());  //休暇付与区分
                if (!row["基本勤務時間"].Equals(DBNull.Value)) kinmu.SelectedIndex = kinmu.FindString(row["基本勤務時間"].ToString());  //勤務時間

                if (!row["通勤手当区分"].Equals(DBNull.Value))
                {
                    tuukinteatekubun.SelectedIndex = tuukinteatekubun.FindString(row["通勤手当区分"].ToString());  //通勤手当区分
                }
                else
                {
                    tuukinteatekubun.SelectedIndex = 0;
                }
                if (!row["通勤方法"].Equals(DBNull.Value)) tuukinkubun.SelectedIndex = tuukinkubun.FindString(row["通勤方法"].ToString());  //通勤手段
                if (!row["片道距離"].Equals(DBNull.Value)) katakyori.Value = Convert.ToDecimal(row["片道距離"].ToString()); //通勤距離
                //if (!row["片道料金"].Equals(DBNull.Value)) tuutanka.Text = row["片道料金"].ToString(); //片道料金
                if (!row["通勤1日単価"].Equals(DBNull.Value)) tuutanka.Text = Convert.ToInt32(row["通勤1日単価"]).ToString(); //片道料金

                if (!row["姓"].Equals(DBNull.Value)) sei.Text = row["姓"].ToString(); //姓
                if (!row["姓フリ"].Equals(DBNull.Value)) seihuri.Text = row["姓フリ"].ToString(); //姓フリガナ
                if (!row["現郵便番号"].Equals(DBNull.Value)) yuubin.Text = row["現郵便番号"].ToString(); //郵便番号
                if (!row["現住所"].Equals(DBNull.Value)) zyuusyo.Text = row["現住所"].ToString(); //現住所
                if (!row["郵便番号"].Equals(DBNull.Value)) yuubin2.Text = row["郵便番号"].ToString(); //郵便番号
                if (!row["住所"].Equals(DBNull.Value)) zyuusyo2.Text = row["住所"].ToString(); //現住所

                if (!row["税表区分"].Equals(DBNull.Value)) zeikubun.SelectedIndex = zeikubun.FindString(row["税表区分"].ToString());  //税表区分
                if (!row["本人障害"].Equals(DBNull.Value)) syougai.SelectedIndex = syougai.FindString(row["本人障害"].ToString()); //本人障害
                if (!row["寡フ"].Equals(DBNull.Value)) kahu.SelectedIndex = kahu.FindString(row["寡フ"].ToString());  //寡フ
                if (!row["勤労学生"].Equals(DBNull.Value)) kinrou.SelectedIndex = kinrou.FindString(row["勤労学生"].ToString());  //勤労
                if (!row["外国人"].Equals(DBNull.Value)) gaikoku.SelectedIndex = gaikoku.FindString(row["外国人"].ToString());  //外国人
                if (!row["災害"].Equals(DBNull.Value)) saigai.SelectedIndex = saigai.FindString(row["災害"].ToString());  //災害

                if (!row["本給"].Equals(DBNull.Value)) honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()); //本給
                if (!row["職務技能給"].Equals(DBNull.Value)) syokumu.Value = Convert.ToDecimal(row["職務技能給"].ToString()); //職務技能給
                //if (!row["調整手当"].Equals(DBNull.Value)) tyousei.Value = Convert.ToDecimal(row["調整手当"].ToString()); //調整手当
                if (!row["特別手当"].Equals(DBNull.Value)) tokubetsu.Value = Convert.ToDecimal(row["特別手当"].ToString()); //特別手当
                if (!row["役職手当"].Equals(DBNull.Value)) yakuteate.Text = row["役職手当"].ToString(); //役職手当
                //if (!row["現場手当"].Equals(DBNull.Value)) menkyo.Text = row["現場手当"].ToString(); //現場手当
                if (!row["免許手当"].Equals(DBNull.Value)) menkyoato.Text = row["免許手当"].ToString(); //免許手当
                if (!row["扶養手当"].Equals(DBNull.Value)) huyouato.Text = row["扶養手当"].ToString(); //扶養手当

                if (!row["転勤手当"].Equals(DBNull.Value)) syukkou.Value = Convert.ToDecimal(row["転勤手当"].ToString()); //転勤手当
                if (!row["通勤手当非"].Equals(DBNull.Value)) tuukinhi2.Text = row["通勤手当非"].ToString(); //通勤手当(非)
                if (!row["通勤手当課"].Equals(DBNull.Value)) tuukinka2.Text = row["通勤手当課"].ToString(); //通勤手当(課)
                if (!row["登録手当"].Equals(DBNull.Value)) tourokuato.Text = row["登録手当"].ToString();    //登録手当
                if (!row["通信手当"].Equals(DBNull.Value)) tuushin.Text = row["通信手当"].ToString();    //通信手当
                if (!row["車両手当"].Equals(DBNull.Value)) syaryou.Text = row["車両手当"].ToString();    //通信手当

                if (!row["友の会"].Equals(DBNull.Value)) tomonokai.Text = row["友の会"].ToString(); //友の会
                //if (!row["固定1"].Equals(DBNull.Value)) kotei1.Value = Convert.ToDecimal(row["固定1"].ToString());    //固定1
                //if (!row["固定2"].Equals(DBNull.Value)) kotei2.Value = Convert.ToDecimal(row["固定2"].ToString());    //固定2

                if(!row["発令区分"].Equals(DBNull.Value)) hatsurei.Text = row["発令区分"].ToString();
                if (!row["備考1"].Equals(DBNull.Value)) bikou1.Text = row["備考1"].ToString();
                if (!row["備考2"].Equals(DBNull.Value)) bikou2.Text = row["備考2"].ToString();
                if (!row["備考3"].Equals(DBNull.Value)) bikou3.Text = row["備考3"].ToString();
                if (!row["備考4"].Equals(DBNull.Value)) bikou4.Text = row["備考4"].ToString();
                if (!row["備考5"].Equals(DBNull.Value)) tokureason.Text = row["備考5"].ToString();

                //if (!row["固定控除1理由"].Equals(DBNull.Value)) koteik1reason.Text = row["固定控除1理由"].ToString();
                //if (!row["固定控除2理由"].Equals(DBNull.Value)) koteik2reason.Text = row["固定控除2理由"].ToString();
                
                Disp();
            }

            //年齢給取得
            GetNenreiKyuu();

            //資格・家族情報取得
            DataDispShikaku();
            DataDispKazoku();

            //資格・家族情報一覧
            shikakudgv.CurrentCell = null;
            kazokudgv.CurrentCell = null;
        }


        //通勤手当元データ登録更新
        private void SetTuukinMotoData()
        {

            DataTable checkdt = Com.GetDB("select 通勤手当区分, 通勤方法, 片道距離, 通勤1日単価 from t通勤手当元データ where 社員番号 = '" + syainno.Text + "' and 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

            if (checkdt.Rows.Count > 0)
            {
                //入力済データ値と同じならスルー
                if (tuukinteatekubun.SelectedItem?.ToString() == checkdt.Rows[0][0].ToString())
                {
                    if (katakyori.Value.ToString() == checkdt.Rows[0][2].ToString())
                    {
                        if (tuukinkubun.SelectedItem?.ToString() == checkdt.Rows[0][1].ToString())
                        {
                            if (tuutanka.Text == checkdt.Rows[0][3].ToString())
                            {
                                return;
                            }
                        }
                    }
                }
            }
            else
            {
                //異動前データ値と同じならスルー
                if (tuukinteatekubun.SelectedItem?.ToString() == tuukinteatekubun_m.Text)
                {
                    if (katakyori.Value.ToString() == katakyori_m.Text)
                    {
                        if (tuukinkubun.SelectedItem?.ToString() == tuukinkubun_m.Text)
                        {
                            if (tuutanka.Text == tuutanka_m.Text)
                            {
                                return;
                            }
                        }
                    }
                }
            }

            //t通勤管理テーブルの通勤方法を更新
            DataTable tuukindt = Com.GetDB("update t通勤管理テーブル set 通勤方法 = '" + tuukinkubun.SelectedItem?.ToString() + "' where 社員番号 = '" + syainno.Text + "' and 管理No = '1'");

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                //トランザクション
                //using (SqlTransaction transaction = Cn.BeginTransaction())
                //{ 

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;

                    //①追加　　　　ないのはありえないから処理は、無し
                    //②更新と追加　適用終了日が'9999/12/31'データの適用開始日が異動日より過去
                    //③更新　　　　適用終了日が'9999/12/31'データの適用開始日が異動日である
                    //④例外　　　　適用終了日が'9999/12/31'データの適用開始日が異動日より未来

                    DataTable flgdata = Com.GetDB("select 適用開始日 from t通勤手当元データ where 社員番号 = '" + syainno.Text + "' and 適用終了日 = '9999/12/31'");

                    //適用開始日が異動日より若い
                    if (Convert.ToDateTime(flgdata.Rows[0][0]) < Convert.ToDateTime(idoudaynew.Value))
                    {
                        //②の更新
                        DataTable update = Com.GetDB("update t通勤手当元データ set 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "' where 社員番号 = '" + syainno.Text + "' and 適用終了日 = '9999/12/31'");

                        //②の追加
                        Cmd.CommandText = "[dbo].[t通勤手当元データインサート]";
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("適用開始日", SqlDbType.Char)); Cmd.Parameters["適用開始日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("適用終了日", SqlDbType.Char)); Cmd.Parameters["適用終了日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当区分", SqlDbType.Char)); Cmd.Parameters["通勤手当区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤方法", SqlDbType.Char)); Cmd.Parameters["通勤方法"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("片道距離", SqlDbType.Char)); Cmd.Parameters["片道距離"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤1日単価", SqlDbType.Char)); Cmd.Parameters["通勤1日単価"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["社員番号"].Value = this.syainno.Text;
                        Cmd.Parameters["適用開始日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");
                        Cmd.Parameters["適用終了日"].Value = "9999/12/31";
                        //Cmd.Parameters["通勤手当区分"].Value = tuukinteatekubun.SelectedItem?.ToString();
                        
                        Cmd.Parameters["通勤手当区分"].Value = tuukinteatekubun.SelectedItem?.ToString() == "" ? SqlStrNull : tuukinteatekubun.SelectedItem?.ToString();
                        Cmd.Parameters["通勤方法"].Value = tuukinkubun.SelectedItem?.ToString();
                        Cmd.Parameters["片道距離"].Value = katakyori.Value.ToString();
                        Cmd.Parameters["通勤1日単価"].Value = tuutanka.Text;

                        //using (dr = Cmd.ExecuteReader())
                        //{
                        //    int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        //}
                        using (dr = Cmd.ExecuteReader())
                        {
                        }
                    }
                    else if (Convert.ToDateTime(flgdata.Rows[0][0]) == Convert.ToDateTime(idoudaynew.Value))
                    {
                        //③更新　　適用開始日が異動日である
                        Cmd.CommandText = "[dbo].[t通勤手当元データアップデート]";
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("適用開始日", SqlDbType.Char)); Cmd.Parameters["適用開始日"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("適用終了日", SqlDbType.Char)); Cmd.Parameters["適用終了日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当区分", SqlDbType.Char)); Cmd.Parameters["通勤手当区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤方法", SqlDbType.Char)); Cmd.Parameters["通勤方法"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("片道距離", SqlDbType.Decimal)); Cmd.Parameters["片道距離"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤1日単価", SqlDbType.Decimal)); Cmd.Parameters["通勤1日単価"].Direction = ParameterDirection.Input;
                        
                        Cmd.Parameters["社員番号"].Value = this.syainno.Text;
                        Cmd.Parameters["適用開始日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");
                        //Cmd.Parameters["適用終了日"].Value = this.syainno.Text;
                        Cmd.Parameters["通勤手当区分"].Value = tuukinteatekubun.SelectedItem?.ToString() == "" ? SqlStrNull : tuukinteatekubun.SelectedItem?.ToString();
                        Cmd.Parameters["通勤方法"].Value = tuukinkubun.SelectedItem?.ToString();
                        Cmd.Parameters["片道距離"].Value = katakyori.Value.ToString();
                        Cmd.Parameters["通勤1日単価"].Value = tuutanka.Text;

                        using (dr = Cmd.ExecuteReader())
                        {
                            //TODO
                        }
                    }
                    else
                    {
                        MessageBox.Show("未来日付で更新データがあるみたいですけど。。");
                        return;
                    }



                //}
                }
            }
        }

        //データ出力
        private void button2_Click(object sender, EventArgs e)
        {
            if (name.Text == "-")
            {
                MessageBox.Show("いやいや");
                return;
            }

            if (hatsurei.Text == "")
            {
                MessageBox.Show("一度、登録ボタンをおしてから出力してください");
                return;
            }

            if (katakyori.Value == 0　&& tuukinteatekubun.SelectedItem?.ToString() != "1 実費精算")
            {
                string msg = "通勤距離が入力されていません。" + nl + "　2020年12月より通勤手段や契約区分に限らず、全従業員必須入力となります。" + nl + "　お手数ですがGoogleMapにて距離測定し、ご入力頂くようお願い致します。" + nl + "　※100m以内でしたら0.1を入力してください。" + nl;
                MessageBox.Show(msg);
                return;
            }

            if (katakyori.Value.ToString() == katakyori_m.Text && (genba.SelectedItem.ToString() != genba_m.Text || zyuusyo.Text != zyuusyo_m.Text))
            {
                string msg = "住所または現場が変更されましたが、通勤距離が変更されてません。通勤距離は変わらずでよろしいですか？" + nl;
                       
                DialogResult result = MessageBox.Show(msg,
                                "警告",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Exclamation,
                                MessageBoxDefaultButton.Button2);

                //何が選択されたか調べる
                if (result == DialogResult.No) return;
            }


            //GetHatsurei();

            //if (hatsurei.Text == "")
            //{
            //    hatsurei.SelectedIndex = 0;
            //}

            //TODO　必須項目未入力有無チェック

            //全体登録更新
            DataInsertUpdate();

            //TODO　テーブルより再取得必要か


            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button2.Enabled = false;

            string fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\20_異動登録.xlsx";

            //手順1：新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            c1XLBook1.Load(fileName);



            //=====================================================
            //TODO 背景色を変更する処理
            //XLSheet sheet1 = c1XLBook1.Sheets[1];
            //sheet1[1, 1].Style.BackColor = Color.re;
            //=====================================================

            // 手順2：セルに値を挿入します。
            XLSheet sheet = c1XLBook1.Sheets[4];


            sheet[0, 1].Value = idoudaynew.Text;//異動日
            sheet[1, 1].Value = hatsurei.Text.Split('　')[1];//発令区分 
            sheet[2, 1].Value = syainno.Text;//社員番号

            sheet[3, 1].Value = kana.Text; //フリガナ
            sheet[4, 1].Value = name.Text; //氏名
            sheet[5, 1].Value = seinengappi.Text; //生年月日
            sheet[6, 1].Value = seibetsu.Text; //性別

            sheet[0, 4].Value = comboBoxKeiken.Text; //最終学歴
            sheet[1, 4].Value = comboBoxGaku.Text; //社外経験

            sheet[2, 4].Value = nyuusya.Text; //入社年月日
            sheet[3, 4].Value = kinzoku.Text; //勤続年数
            sheet[4, 4].Value = nyuusyaold.Text; //入社時年齢


            //異動元

            //地区
            sheet[10, 1].Value = tiku_m.Text.Split('　')[0];
            sheet[10, 2].Value = tiku_m.Text.Split('　')[1];

            //組織
            sheet[11, 1].Value = soshiki_m.Text.Split('　')[0];
            sheet[11, 2].Value = soshiki_m.Text.Split('　')[1];

            //現場名
            sheet[12, 1].Value = genba_m.Text.Split('　')[0];
            sheet[12, 2].Value = genba_m.Text.Split('　')[1];

            //給与支給区分
            sheet[13, 2].Value = kyuuyo_m.Text.Split('　')[1];

            //休日区分
            sheet[14, 2].Value = kyuuzitsu_m.Text.Split('　')[1];

            //職種
            sheet[15, 2].Value = syokusyu_m.Text;

            //役職
            sheet[16, 2].Value = yakusyoku_m.Text.Split('　')[1];

            //試用期間
            sheet[17, 2].Value = shiyou_m.Text == "" ? "" : shiyou_m.Text.Split('　')[1];

            //契約社員
            sheet[18, 2].Value = keiyaku_m.Text == "" ? "" : keiyaku_m.Text.Split('　')[1];

            //友の会区分
            sheet[19, 2].Value = tomokubun_m.Text == "" ? "" : tomokubun_m.Text.Split('　')[1];


            //異動先
            //地区
            sheet[10, 3].Value = tiku.Text.Split('　')[0];
            sheet[10, 4].Value = tiku.Text.Split('　')[1];

            //組織
            sheet[11, 3].Value = soshiki.Text.Split('　')[0];
            sheet[11, 4].Value = soshiki.Text.Split('　')[1];

            //現場名
            sheet[12, 3].Value = genba.Text.Split('　')[0];
            sheet[12, 4].Value = genba.Text.Split('　')[1];

            //給与支給区分
            sheet[13, 4].Value = kyuuyo.Text == "" ? "" : kyuuyo.Text.Split('　')[1];

            //休日区分
            sheet[14, 4].Value = kyuuzitsukubun.Text == "" ? "" : kyuuzitsukubun.Text.Split('　')[1];

            //職種
            sheet[15, 4].Value = comboBoxSyoku.Text;

            //役職
            sheet[16, 4].Value = yakusyoku.Text.Split('　')[1];

            //試用期間
            sheet[17, 4].Value = shiyou.Text == "" ? "" : shiyou.Text.Split('　')[1];

            //契約社員
            sheet[18, 4].Value = keiyaku.Text == "" ? "" : keiyaku.Text.Split('　')[1];

            //友の会区分
            sheet[19, 4].Value = tomokubun.Text == "" ? "" : tomokubun.Text.Split('　')[1];


            //25　税表区分
            sheet[25, 1].Value = zeikubun_m.Text == "" ? "" : zeikubun_m.Text.Split('　')[1];
            sheet[25, 2].Value = zeikubun.Text == "" ? "" : zeikubun.Text.Split('　')[1];

            //26　本人障害
            sheet[26, 1].Value = syougai_m.Text == "" ? "" : syougai_m.Text.Split('　')[1];
            sheet[26, 2].Value = syougai.Text == "" ? "" : syougai.Text.Split('　')[1];

            //27　寡フ
            sheet[27, 1].Value = kahu_m.Text == "" ? "" : kahu_m.Text.Split('　')[1];
            sheet[27, 2].Value = kahu.Text == "" ? "" : kahu.Text.Split('　')[1];

            //28　勤労
            sheet[28, 1].Value = kinrou_m.Text == "" ? "" : kinrou_m.Text.Split('　')[1];
            sheet[28, 2].Value = kinrou.Text == "" ? "" : kinrou.Text.Split('　')[1];

            //29　外国人
            sheet[29, 1].Value = gaikoku_m.Text == "" ? "" : gaikoku_m.Text.Split('　')[1];
            sheet[29, 2].Value = gaikoku.Text == "" ? "" : gaikoku.Text.Split('　')[1];

            //30　災害
            sheet[30, 1].Value = saigai_m.Text == "" ? "" : saigai_m.Text.Split('　')[1];
            sheet[30, 2].Value = saigai.Text == "" ? "" : saigai.Text.Split('　')[1];

            //時給
            sheet[36, 1].Value = Convert.ToInt32(zikyuu_m.Text);
            sheet[36, 2].Value = Convert.ToInt32(zikyuu.Value);
            //日給
            sheet[37, 1].Value = nikkyuu_m.Text;
            sheet[37, 2].Value = nikkyuu.Text;
            //回数1
            sheet[38, 1].Value = Convert.ToInt32(kaisuu1_m.Text);
            sheet[38, 2].Value = Convert.ToInt32(kaisuu1.Value);
            //回数2
            //sheet[39, 1].Value = kaisuu2_m.Text;
            //sheet[39, 2].Value = kaisuu2.Text;
            //休暇付与区分
            sheet[40, 1].Value = kyuuka_m.Text == "" ? "" : kyuuka_m.Text.Split('　')[1];
            sheet[40, 2].Value = kyuuka.Text == "" ? "" : kyuuka.Text.Split('　')[1];
            //勤務時間
            sheet[41, 1].Value = kinmu_m.Text;
            sheet[41, 2].Value = kinmu.Text;


            //通勤手当区分
            sheet[44, 1].Value = "";
            sheet[44, 2].Value = "";

            //通勤手段区分
            sheet[45, 1].Value = tuukinkubun_m.Text;
            sheet[45, 2].Value = tuukinkubun.Text;

            //通勤距離
            sheet[46, 1].Value = katakyori_m.Text;
            sheet[46, 2].Value = katakyori.Text;

            //片道料金
            sheet[47, 1].Value = tuutanka_m.Text;
            sheet[47, 2].Value = tuutanka.Text;


            //現郵便番号
            sheet[51, 1].Value = yuubin_m.Text;
            sheet[51, 2].Value = yuubin.Text;

            //現住所
            sheet[52, 1].Value = zyuusyo_m.Text;
            sheet[52, 2].Value = zyuusyo.Text;

            //郵便番号
            sheet[53, 1].Value = yuubin_m2.Text;
            sheet[53, 2].Value = yuubin2.Text;

            //住所
            sheet[54, 1].Value = zyuusyo_m2.Text;
            sheet[54, 2].Value = zyuusyo2.Text;

            //本給
            sheet[57, 1].Value = Convert.ToInt32(honkyuu_m.Text);
            sheet[57, 2].Value = Convert.ToInt32(honkyuu.Value);

            //職務技能給
            sheet[58, 1].Value = Convert.ToInt32(syokumu_m.Text);
            sheet[58, 2].Value = Convert.ToInt32(syokumu.Value);
            //特別手当
            sheet[59, 1].Value = Convert.ToInt32(tokubetsu_m.Text);
            sheet[59, 2].Value = Convert.ToInt32(tokubetsu.Value);
            //役職手当
            sheet[60, 1].Value = Convert.ToInt32(yakuteate_m.Text);
            sheet[60, 2].Value = Convert.ToInt32(yakuteate.Text);
            //免許手当
            sheet[61, 1].Value = Convert.ToInt32(menkyo_m.Text);
            sheet[61, 2].Value = Convert.ToInt32(menkyo.Text);
            //扶養手当
            sheet[62, 1].Value = Convert.ToInt32(huyou_m.Text);
            sheet[62, 2].Value = Convert.ToInt32(huyou.Text);
            //離島手当
            sheet[63, 1].Value =Convert.ToInt32(ritou_m.Text);
            sheet[63, 2].Value = Convert.ToInt32(ritou.Text);


            //転勤手当
            sheet[64, 1].Value = Convert.ToInt32(syukkou_m.Text);
            sheet[64, 2].Value = Convert.ToInt32(syukkou.Value);
            //通勤手当(非)
            sheet[65, 1].Value = Convert.ToInt32(tuukinhi_m.Text);
            sheet[65, 2].Value = Convert.ToInt32(tuukinhi.Text);
            //通勤手当(課)
            sheet[66, 1].Value = Convert.ToInt32(tuukinka_m.Text);
            sheet[66, 2].Value = Convert.ToInt32(tuukinka.Text);
            //登録手当
            sheet[67, 1].Value = Convert.ToInt32(touroku_m.Text);
            sheet[67, 2].Value = Convert.ToInt32(touroku.Text);
            //通信手当
            sheet[68, 1].Value = Convert.ToInt32(tuushin_m.Text);
            sheet[68, 2].Value = Convert.ToInt32(tuushin.Text);
            //車両手当
            sheet[69, 1].Value = Convert.ToInt32(syaryou_m.Text);
            sheet[69, 2].Value = Convert.ToInt32(syaryou.Text);


            int yaku = Convert.ToInt32(yakusyoku.SelectedItem.ToString().Substring(0, 4));
            int yaku_m = Convert.ToInt32(yakusyoku_m.Text.ToString().Substring(0, 4));
            string kubun = kyuuyo.Text == "" ? "" : kyuuyo.Text.Split('　')[1];
            string kubun_m = kyuuyo_m.Text == "" ? "" : kyuuyo_m.Text.Split('　')[1];

            //職種
            //if (comboBoxSyoku.SelectedItem == null || comboBoxSyoku.SelectedItem.ToString() == "" || yaku <= 135)
            //{

            if (yaku_m > 135 && kubun_m == "月給者")
            {
                //職務給
                sheet[70, 1].Value = Convert.ToInt32(syokumukyuu_m.Text);
                //学歴給
                sheet[71, 1].Value = Convert.ToInt32(gakurekikyuu_m.Text);
                //経験給
                sheet[72, 1].Value = Convert.ToInt32(keikenkyuu_m.Text);
                //年齢給
                sheet[73, 1].Value = Convert.ToInt32(nennreikyuu_m.Text);
                //基準外
                sheet[74, 1].Value = Convert.ToInt32(kizyungai_m.Text);
                //評価給
                sheet[75, 1].Value = Convert.ToInt32(hyoukakyuu_m.Text);
            }

            if (yaku > 135 && kubun == "月給者")
            {
                //職務給
                sheet[70, 2].Value = Convert.ToInt32(syokumukyuu.Text);
                //学歴給
                sheet[71, 2].Value = Convert.ToInt32(gakurekikyuu.Text);
                //経験給
                sheet[72, 2].Value = Convert.ToInt32(keikenkyuu.Text);
                //年齢給
                sheet[73, 2].Value = Convert.ToInt32(nennreikyuu.Text);
                //基準外
                sheet[74, 2].Value = Convert.ToInt32(kizyungai.Text);
                //評価給
                sheet[75, 2].Value = Convert.ToInt32(hyoukakyuu.Text);
            }

            ////友の会
            sheet[78, 1].Value = Convert.ToInt32(tomonokai_m.Text);
            sheet[78, 2].Value = Convert.ToInt32(tomonokai.Text);
            ////固定控除1
            sheet[79, 1].Value = 0;
            sheet[79, 2].Value = 0;
            ////固定控除2
            sheet[80, 1].Value = 0;
            sheet[80, 2].Value = 0;

            //その他情報
            sheet[57, 5].Value = DateTime.Now.ToString("yyyy年MM月dd日"); //作成日
            sheet[58, 5].Value = Program.loginname; //作成者
            sheet[59, 5].Value = Convert.ToDateTime(idoudaynew.Value).AddMonths(1).ToString("yyyy年MM月支給 ") + Convert.ToDateTime(idoudaynew.Value).ToString("(yyyy年MM月分)"); //給与計算月度　YYYY年MM月支給 (XXXX分)

            //備考1～4
            sheet[64, 5].Value = bikou1.Text;
            sheet[65, 5].Value = bikou2.Text;
            sheet[66, 5].Value = bikou3.Text;
            sheet[67, 5].Value = bikou4.Text;

            //名前変更
            sheet[84, 1].Value = sei_m.Text == sei.Text ? "" : sei.Text;
            sheet[84, 2].Value = seihuri_m.Text == seihuri.Text ? "" : seihuri.Text;

            //家族情報
            DataTable kdt = new DataTable();
            kdt = Com.GetDB("select * from dbo.k家族情報取得_異動('" + syainno.Text + "', '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' ) order by 家族識別ＩＤ");

            //家族データ
            int r3 = kdt.Rows.Count;
            int c3 = kdt.Columns.Count;

            for (int i = 0; i < r3; i++)
            {
                for (int j = 0; j < c3; j++)
                {
                    sheet[j, i + 7].Value = kdt.Rows[i][j];
                }
            }

            //免許手当
            //資格
            DataTable sdt = new DataTable();
            sdt = Com.GetDB("select * from s資格データ取得 where LEFT(資格コード,1) <> 'T' and 社員番号 = '" + syainno.Text + "' and 適用終了日 > '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

            //資格データ
            int r2 = sdt.Rows.Count;
            int c2 = sdt.Columns.Count;

            for (int i = 0; i < r2; i++)
            {
                for (int j = 0; j < c2; j++)
                {
                    sheet[j, i + 20].Value = sdt.Rows[i][j];
                }
            }

            //登録手当
            DataTable tdt = new DataTable();
            tdt = Com.GetDB("select * from s資格データ取得 where LEFT(資格コード,1) = 'T' and 社員番号 = '" + syainno.Text + "' and 適用終了日 > '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

            //資格データ
            int r4 = tdt.Rows.Count;
            int c4 = tdt.Columns.Count;

            for (int i = 0; i < r4; i++)
            {
                for (int j = 0; j < c4; j++)
                {
                    sheet[j+15, i + 20].Value = tdt.Rows[i][j];
                }
            }


            string localPass = @"C:\ODIS\IDOU\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_") + name.Text.Replace(" ", "").Replace("　", "");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            // 手順3：ファイルを保存します。
            c1XLBook1.Save(exlName + ".xlsx");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;

            //excel出力
            System.Diagnostics.Process.Start(exlName + ".xlsx");
        }


        private void honkyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }


        private void validCommon(Control cont, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError(cont, "");

            //全角は半角に
            String result = Microsoft.VisualBasic.Strings.StrConv(cont.Text, VbStrConv.Narrow).Trim();

            ////空白チェック
            if (result == "") return;

            //数値以外除去 
            Regex regex = new Regex(@"^[0-9]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError(cont, "数字以外は入力しないでください");
                e.Cancel = true;
            }

            //全角入力対応
            cont.Text = result;
        }

        //対象者変更
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            //ダブルクリックでしかだめ！ 現行不要
        }

        //対象者変更時の全コントロールリセット
        private void DataReset()
        {
            //基本情報
            DateTime firstDayOfMonth1 = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            syainno.Text = "";
            //TODO
            idoudaynew.Value = null;
            //idouday.Value = firstDayOfMonth1;
            kana.Text = "";
            name.Text = "";
            seinengappi.Text = "";
            nyuusya.Text = "";
            kinzoku.Text = "";
            seibetsu.Text = "";
            shiyou_m.Text = "";
            nyuusyaold.Text = "";
            lbl_no.Text = "0";

            hatsurei.SelectedIndex = -1;

            //項目情報 変更後
            tiku.SelectedIndex = -1;
            soshiki.SelectedIndex = -1;
            genba.SelectedIndex = -1;
            kyuuyo.SelectedIndex = -1;
            yakusyoku.SelectedIndex = -1;
            keiyaku.SelectedIndex = -1;
            tomokubun.SelectedIndex = -1;

            kyuuzitsukubun.SelectedIndex = -1;
            shiyou.SelectedIndex = -1;

            comboBoxSyoku.SelectedIndex = -1;
            comboBoxGaku.SelectedIndex = -1;
            comboBoxKeiken.SelectedIndex = -1;

            syokumukyuu.Text = "0";
            gakurekikyuu.Text = "0";
            keikenkyuu.Text = "0";
            nennreikyuu.Text = "0";
            kizyungai.Text = "0";
            hyoukakyuu.Text = "0";

            syokumukyuu_m.Text = "0";
            gakurekikyuu_m.Text = "0";
            keikenkyuu_m.Text = "0";
            nennreikyuu_m.Text = "0";
            kizyungai_m.Text = "0";
            hyoukakyuu_m.Text = "0";
            //項目情報 変更前                  
            tiku_m.Text = "";
            soshiki_m.Text = "";
            genba_m.Text = "";
            kyuuyo_m.Text = "";

            yakusyoku_m.Text = "";
            keiyaku_m.Text = "";
            tomokubun_m.Text = "";

            kyuuzitsu_m.Text = "";
            //toukyuu_m.Text = "";

            syokusyu_m.Text = "";
            gakureki_m.Text = "";
            keiken_m.Text = "";
            nyuusyaold_m.Text = "";

            //項目情報 ラベル
            lbl_tiku.BackColor = Color.White;
            lbl_soshiki.BackColor = Color.White;
            lbl_genba.BackColor = Color.White;
            lbl_kyuuyo.BackColor = Color.White;
            lbl_yakusyoku.BackColor = Color.White;
            lbl_keiyaku.BackColor = Color.White;
            lbl_tomo.BackColor = Color.White;

            lbl_kyuuzitsu.BackColor = Color.White;
            lbl_shiyou.BackColor = Color.White;

            lbl_syokumu.BackColor = Color.White;
            lbl_gakureki.BackColor = Color.White;
            lbl_keiken.BackColor = Color.White;
            lbl_goukei.BackColor = Color.White;

            //税区分　変更後
            zeikubun.SelectedIndex = -1;
            syougai.SelectedIndex = -1;
            kahu.SelectedIndex = -1;
            kinrou.SelectedIndex = -1;
            gaikoku.SelectedIndex = -1;
            saigai.SelectedIndex = -1;
            //税区分　変更前
            zeikubun_m.Text = "";
            syougai_m.Text = "";
            kahu_m.Text = "";
            kinrou_m.Text = "";
            gaikoku_m.Text = "";
            saigai_m.Text = "";
            //税区分　ラベル
            label34.BackColor = Color.White;
            label35.BackColor = Color.White;
            label38.BackColor = Color.White;
            label39.BackColor = Color.White;
            label48.BackColor = Color.White;
            label49.BackColor = Color.White;

            //苗字変更
            sei.Text = "";
            seihuri.Text = "";
            sei_m.Text = "";
            seihuri_m.Text = "";
            lbl_sei.BackColor = Color.White;
            lbl_seihuri.BackColor = Color.White;

            //現郵便住所
            yuubin.Text = "";
            zyuusyo.Text = "";
            yuubin_m.Text = "";
            zyuusyo_m.Text = "";
            lbl_yuubin.BackColor = Color.White;
            lbl_zyuusyo.BackColor = Color.White;

            yuubin2.Text = "";
            zyuusyo2.Text = "";
            yuubin_m2.Text = "";
            zyuusyo_m2.Text = "";
            lbl_yuubin2.BackColor = Color.White;
            lbl_zyuusyo2.BackColor = Color.White;

            //時給・回数・週休・勤務時間等
            zikyuu.Value = 0;
            nikkyuu.Value = 0;
            kaisuu1.Value = 0;
            //kaisuu2.Value = 0;
            kyuuka.SelectedIndex = 0;
            kinmu.SelectedIndex = 0;

            zikyuu_m.Text = "";
            nikkyuu_m.Text = "";
            kaisuu1_m.Text = "";
            //kaisuu2_m.Text = "";
            kyuuka_m.Text = "";
            kinmu_m.Text = "";

            label4.BackColor = Color.White;
            label5.BackColor = Color.White;
            label28.BackColor = Color.White;
            //label29.BackColor = Color.White;
            label30.BackColor = Color.White;
            label31.BackColor = Color.White;


            //通勤項目
            tuukinteatekubun.SelectedIndex = -1;
            tuukinkubun.SelectedIndex = -1;
            katakyori.Value = 0;
            tuutanka.Text = "0";

            tuukinteatekubun_m.Text = "";
            tuukinkubun_m.Text = "";
            katakyori_m.Text = "0.0";
            tuutanka_m.Text = "0";

            lbl_tuukinteatekubun.BackColor = Color.White;
            lbl_tuukinkubun.BackColor = Color.White;
            lbl_katakyori.BackColor = Color.White;
            lbl_tuutanka.BackColor = Color.White;
            lbl_tuukinhi2.BackColor = Color.White;
            lbl_tuukinka2.BackColor = Color.White;

            tuukinhi2.Text = "0"; //通勤非
            tuukinka2.Text = "0"; //通勤

            tuukinhi2_m.Text = "0"; //通勤非
            tuukinka2_m.Text = "0"; //通勤

            honkyuu.Enabled = false;
            syokumu.Enabled = false;
            //menkyo.Enabled = false;

            //固定給
            honkyuu.Value = 0;
            syokumu.Value = 0;
            //tyousei.Value = 0;
            tokubetsu.Value = 0;
            yakuteate.Text = "0";
            //menkyo.Text = "0";
            menkyo.Text = "0";
            huyou.Text = "0";
            kizyun.Text = "0";
            syukkou.Value = 0;
            tuukinhi.Text = "0";
            tuukinka.Text = "0";
            touroku.Text = "0";
            tuushin.Text = "0";
            syaryou.Text = "0";
            shikyuu.Text = "0";
            tomonokai.Text = "0";
            //kotei1.Value = 0;
            //kotei2.Value = 0;
            //kouzyo.Text = "0";

            honkyuu_m.Text = "0";
            syokumu_m.Text = "0";
            //tyousei_m.Text = "0";
            tokubetsu_m.Text = "0";
            yakuteate_m.Text = "0";
            menkyo_m.Text = "0";
            huyou_m.Text = "0";
            kizyun_m.Text = "0";
            syukkou_m.Text = "0";
            tuukinhi_m.Text = "0";
            tuukinka_m.Text = "0";
            touroku_m.Text = "0";
            tuushin_m.Text = "0";
            syaryou_m.Text = "0";
            shikyuu_m.Text = "0";
            tomonokai_m.Text = "0";
            //kotei1_m.Text = "0";
            //kotei2_m.Text = "0";
            //kouzyo_m.Text = "0";

            lbl_honkyuu.BackColor = Color.White;
            lbl_syokumu.BackColor = Color.White;
            lbl_tokubetsu.BackColor = Color.White;
            lbl_yakuteate.BackColor = Color.White;
            //lbl_genbateate.BackColor = Color.White;
            lbl_menkyo.BackColor = Color.White;
            lbl_huyou.BackColor = Color.White;
            lbl_kizyun.BackColor = Color.White;
            lbl_syukkou.BackColor = Color.White;
            lbl_tuukinhi.BackColor = Color.White;
            lbl_tuukinka.BackColor = Color.White;
            lbl_touroku.BackColor = Color.White;
            lbl_tuushin.BackColor = Color.White;
            lbl_syaryou.BackColor = Color.White;
            lbl_shikyuu.BackColor = Color.White;
            lbl_tomonokai.BackColor = Color.White;
            //lbl_kotei1.BackColor = Color.White;
            //lbl_kotei2.BackColor = Color.White;
            //lbl_kouzyo.BackColor = Color.White;
            lbl_ritou.BackColor = Color.White;

            bikou1.Text = "";
            bikou2.Text = "";
            bikou3.Text = "";
            bikou4.Text = "";

            //家族情報クリア
            KazokuClear();

            huyoumae.Text = "0";
            huyouato.Text = "0";

            //資格情報クリア
            ShikakuClear();

            menkyomae.Text = "0";
            menkyoato.Text = "0";

            tourokumae.Text = "0";
            tourokuato.Text = "0";

            //登録ボタン
            button1.Enabled = true;

            //労働条件
            keiyakunengetsu.Value = null;
            koyoukubun.Text = null;
            koyoukaishibi.Value = null;
            koyousyuuryoubi.Value = null;
            koushinkubun.Text = null;
            syuugyoubasyo.Text = null;
            gyoumunaiyou.Text = null;
            dtps0.Value = zerodt;
            dtpe0.Value = zerodt;
            dtps1.Value = zerodt;
            dtpe1.Value = zerodt;
            dtps2.Value = zerodt;
            dtpe2.Value = zerodt;
            dtps3.Value = zerodt;
            dtpe3.Value = zerodt;
            dtps4.Value = zerodt;
            dtpe4.Value = zerodt;
            dtps5.Value = zerodt;
            dtpe5.Value = zerodt;
            zikangairoudou.Text = null;
            yakankinmu.Text = null;
            kyuujitsukaisuu.Text = null;
            kyuusyustu.Text = null;
            teinen.Text = null;
            syouyo.Text = null;
            taisyokukin.Text = null;
            kinmuH.Text = null;


            tokureason.Text = ""; //特別手当付与理由

            //固定控除
            //koteik1reason.Text = "";
            //koteik2reason.Text = "";
        }

        //職務技能給の合計・基準内賃金・支給額・控除額の計算
        private void GetKizyunnai()
        {
            int kizyunnai = 0;
            int shikyuusum = 0;
            //int kouzyosum = 0;

            //職務技能給合計と職務技能給
            GetSyokumu();

            if (honkyuu.Value.ToString() == "" || syokumu.Value.ToString() == "" ||  tokubetsu.Text == "" || yakuteate.Text == "" || menkyo.Text == "" || ritou.Text == "" || syukkou.Value.Equals(null) || touroku.Text == "" || tuushin.Text == "" || syaryou.Text == "") return;
            kizyunnai = Convert.ToInt32(honkyuu.Value) + Convert.ToInt32(syokumu.Value) + 
            Convert.ToInt32(tokubetsu.Value) + Convert.ToInt32(yakuteate.Text) + Convert.ToInt32(menkyo.Text) + //Convert.ToInt32(menkyo.Text) +
            Convert.ToInt32(ritou.Text) + Convert.ToInt32(touroku.Text) + Convert.ToInt32(tuushin.Text) + Convert.ToInt32(syukkou.Value) + Convert.ToInt32(syaryou.Text);
            kizyun.Text = kizyunnai.ToString();

            if (tuukinhi.Text == "" || tuukinka.Text == "" || huyou.Text == "") return;
            shikyuusum = kizyunnai + Convert.ToInt32(tuukinhi.Text.Replace(".0", "")) + Convert.ToInt32(tuukinka.Text.Replace(".0", "")) + Convert.ToInt32(huyou.Text);
            shikyuu.Text = shikyuusum.ToString();

            //if (tomonokai.Text == "" || kotei1.Text == "" || kotei2.Text == "") return;
            //kouzyosum = Convert.ToInt32(tomonokai.Text) + Convert.ToInt32(kotei1.Value) + Convert.ToInt32(kotei2.Value);
            //kouzyo.Text = kouzyosum.ToString();
        }

        //変更無の場合は非表示
        private void Disp()
        {
            //集計セル
            GetKizyunnai();
        }


        #region 変更前表示非表示と項目背景色の変更
        //地区CD
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            soshiki.Items.Clear();

            DataRow[] dr = soshikidt.Select("組織CD like '" + (tiku.SelectedIndex + 1).ToString() + "%'");

            foreach (DataRow drw in dr)
            {
                soshiki.Items.Add(drw["組織CD"].ToString() + "　" + drw["組織名"].ToString());
            }

            //TODO 2020/04/02 変更
            soshiki.SelectedIndex = -1;

            //表示・非表示
            if (tiku.SelectedItem?.ToString() == tiku_m.Text)
            {
                lbl_tiku.BackColor = Color.White;
            }
            else
            {
                lbl_tiku.BackColor = Color.LightGreen;
            }
        }

        //組織CD
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            genba.Items.Clear();
            genbadt.Clear();

            GetGenba();
            foreach (DataRow drw in genbadt.Rows)
            {
                genba.Items.Add(drw["現場CD"].ToString() + "　" + drw["現場名"].ToString());
            }

            genba.SelectedIndex = 0;

            //表示・非表示
            if (soshiki.SelectedItem?.ToString() == soshiki_m.Text)
            {
                lbl_soshiki.BackColor = Color.White;
                soshikiflg = false;
            }
            else
            {
                lbl_soshiki.BackColor = Color.LightGreen;
                soshikiflg = true;
            }

            RitouCalc();
            //TODO 発令整理20200330
            //GetHatsurei();
        }

        //現場CD
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (genba.SelectedItem?.ToString() == genba_m.Text)
            {
                lbl_genba.BackColor = Color.White;
                genbaflg = false;
            }
            else
            {
                lbl_genba.BackColor = Color.LightGreen;
                genbaflg = true;
            }

            RitouCalc();
        }

        private void RitouCalc()
        {
            string ss = "0";
            string kyuu = kyuuyo.SelectedItem?.ToString().Substring(0, 1);
            int yaku = Convert.ToInt32(yakusyoku.SelectedItem?.ToString().Substring(0, 4));
            string syoku = comboBoxSyoku.SelectedItem?.ToString();
            string soshi = soshiki.SelectedItem?.ToString().Split('　')[0];
            string gen = genba.SelectedItem?.ToString().Split('　')[0];

            if (kyuu == null || syoku == null || soshi == null || gen == null) return;

            //if (kyuu == "C" && yaku > 135 && syukkou.Value == 0) //係長未満の月給者で転勤手当無

            //TODO 技能実習生を除く場合の処理
            //if (kyuu == "C" && syukkou.Value == 0 && keiyaku.SelectedItem?.ToString().Substring(0,2) != "30")
            if (kyuu == "C" && syukkou.Value == 0) 
            {
                //DataTable dt = Com.GetDB(" select (select 離島手当 from dbo.K_職務給_職種 where 備考 = '" + syoku + "') + (select 地区 from dbo.r離島手当_地区 where code = '" + soshi.Substring(0,1) + "') as 離島手当 ");
                //ss = dt.Rows[0][0].ToString();

                ss = Com.RitouCalc(soshi, syoku);
            }

            ritou.Text = ss;

        }

        //給与支給区分の変更 すげー長い
        private void kyuuyo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (kyuuyo.SelectedIndex == -1) return;

            string kyuu = kyuuyo.SelectedItem.ToString().Substring(0, 1);

            string kyuu_m = "";
            if (kyuuyo_m.Text.Length > 0)
            {
                kyuu_m = kyuuyo_m.Text.ToString().Substring(0, 1);
            }

            //TODO 
            if (yakusyoku.SelectedItem == null)
            {
                yakusyoku.SelectedIndex = 0;
                //return;
            }


            int yaku = Convert.ToInt32(yakusyoku.SelectedItem.ToString().Substring(0, 4));

            //役職コンボボックスクリア
            yakusyoku.Items.Clear();

            //TODO 役員いらない
            //B 兼務役員
            if (kyuu == "A" || kyuu == "B")
            {
                #region いらない
                //社員区分設定
                //kyuuzitsukubun.Text = "00　役員";

                //等級・号俸の設定
                //toukyuu.Enabled = true;
                //toukyuu.SelectedItem = toukyuu_m.Text;

                //時給・日給の設定
                zikyuu.Enabled = false;
                nikkyuu.Enabled = false;

                kyuuka.Enabled = true;
                //とりあえず現行
                kyuuka.SelectedItem = kyuuka_m.Text;

                //役職
                yakusyoku.Items.Add("0050" + "　" + "常務取締役");
                yakusyoku.Items.Add("0020" + "　" + "代表取締役社長");
                yakusyoku.Items.Add("0045" + "　" + "取締役相談役");
                yakusyoku.SelectedIndex = 0;

                //職務技能給
                //

                #endregion
            }
            else if (kyuu == "C")
            {
                //社員区分設定
                //kyuuzitsukubun.Text = "01　月給者";

                //部長以下の本給設定
                if (yaku >= 110)
                {
                    //TODO 
                    DataTable hkdt = new DataTable();
                    hkdt = Com.GetDB("select * from dbo.HK_本給 where '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

                    foreach (DataRow row in hkdt.Rows)
                    {
                        if (keiyaku.SelectedItem?.ToString() != "")
                        {
                            honkyuu.Value = Convert.ToDecimal(honkyuu_m.Text);
                        }
                        else if (yaku == 110) //部長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_部長"].ToString());
                        }
                        else if (yaku == 120) //副部長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_副部長"].ToString());
                        }
                        else if (yaku == 130) //課長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_課長"].ToString());
                        }
                        else if (yaku == 135) //係長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_係長"].ToString());
                        }
                        else if (yaku == 140) //主任
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_主任"].ToString());
                        }
                        else if (yaku == 150) //副主任
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_副主任"].ToString());
                        }
                        else
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString());
                        }
                    }


                    

                    

                    //honkyuu.Value = 142000;
                    //honkyuu.Value = Convert.ToDecimal(hkdt.Rows[0][0].ToString());
                    syokumuginoupanel.Visible = true;
                }
                else
                {

                    //TODO 20200318コメントアウト
                    //係長以上は職務技能給根拠パネルいらない
                    //syokumuginoupanel.Visible = false;
                }


                //等級・号俸の設定
                //toukyuu.Enabled = false;
                //toukyuu.SelectedItem = toukyuu_m.Text;

                //現場手当
                //menkyo.Enabled = false;

                //時給・日給の設定
                zikyuu.Enabled = false;
                nikkyuu.Enabled = false;

                zikyuu.Text = "0";
                nikkyuu.Text = "0";

                kyuuka.Enabled = true;

                //とりあえず現行
                kyuuka.SelectedItem = kyuuka_m.Text;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                //役職
                yakusyoku.Items.Add("0180" + "　" + "係員");
                //yakusyoku.Items.Add("0150" + "　" + "補佐");
                yakusyoku.Items.Add("0150" + "　" + "副主任");
                yakusyoku.Items.Add("0140" + "　" + "主任");
                yakusyoku.Items.Add("0135" + "　" + "係長");
                yakusyoku.Items.Add("0132" + "　" + "技術係長");
                yakusyoku.Items.Add("0130" + "　" + "課長");
                yakusyoku.Items.Add("0122" + "　" + "技術課長");
                yakusyoku.Items.Add("0120" + "　" + "副部長");
                yakusyoku.Items.Add("0112" + "　" + "技術副部長");
                yakusyoku.Items.Add("0110" + "　" + "部長");
                yakusyoku.Items.Add("0102" + "　" + "技術部長");
                //yakusyoku.Items.Add("0100" + "　" + "部長");

                yakusyoku.Items.Add("0070" + "　" + "顧問");
                yakusyoku.Items.Add("0066" + "　" + "相談役");

                yakusyoku.Items.Add("0060" + "　" + "取締役部長");
                yakusyoku.Items.Add("0055" + "　" + "監査役");
                yakusyoku.Items.Add("0050" + "　" + "常務取締役");
                yakusyoku.Items.Add("0045" + "　" + "取締役相談役");
                yakusyoku.Items.Add("0020" + "　" + "代表取締役社長");
                yakusyoku.SelectedIndex = 0;

                //職種
                comboBoxSyoku.Enabled = true;

                if (kyuu_m == "E" || kyuu_m == "F")
                {
                    //学歴
                    comboBoxGaku.Enabled = true;
                    //社外経験
                    comboBoxKeiken.Enabled = true;
                }
                else
                {
                    //学歴
                    comboBoxGaku.Enabled = false;
                    //社外経験
                    comboBoxKeiken.Enabled = false;
                }
            }
            else if (kyuu == "D")
            {
                //社員区分設定
                //kyuuzitsukubun.Text = "D1　日給者";

                //
                honkyuu.Value = Convert.ToDecimal(Convert.ToInt32(nikkyuu.Value) * Getrday(kyuuka.SelectedItem.ToString()));


                if (kyuuyo_m.Text == "C1　月給者")
                {
                    syokumuginoupanel.Visible = false;

                    //職務技能給ゼロ
                    syokumu.Value = 0;

                    //免許ゼロ
                    menkyo.Text = "0";
                }
                else
                {
                    //とりあえず表示
                    syokumuginoupanel.Visible = true;
                }

                //試用期間
                //toukyuu.Enabled = false;
                shiyou.SelectedItem = "";

                //現場手当
                //menkyo.Enabled = false;


                //時給・日給の設定
                zikyuu.Enabled = false;
                nikkyuu.Enabled = true;

                zikyuu.Text = "0";

                kyuuka.Enabled = true;
                //とりあえず現行
                kyuuka.SelectedItem = kyuuka_m.Text;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                //役職
                yakusyoku.Items.Add("0180" + "　" + "係員");
                //yakusyoku.Items.Add("0150" + "　" + "補佐");　
                yakusyoku.Items.Add("0150" + "　" + "副主任"); //TODO 八重山だけ
                yakusyoku.SelectedIndex = 0;

                //職種
                comboBoxSyoku.Enabled = true;

                if (kyuu_m == "E" || kyuu_m == "F")
                {
                    //学歴
                    comboBoxGaku.Enabled = true;
                    //社外経験
                    comboBoxKeiken.Enabled = true;
                }
                else
                {
                    //学歴
                    comboBoxGaku.Enabled = false;
                    //社外経験
                    comboBoxKeiken.Enabled = false;
                }
            }
            else if (kyuu == "E")
            {
                //kyuuzitsukubun.Text = "03　パート";

                //TODO 20200318コメントアウト
                //syokumuginoupanel.Visible = false;

                //等級の設定
                shiyou.SelectedItem = "";

                //契約区分　空白に上書き！！ TODO 202506
                keiyaku.SelectedItem = "";

                //時給・日給の設定
                zikyuu.Enabled = true;
                nikkyuu.Enabled = false;
                nikkyuu.Text = "0";

                //職務技能給
                syokumu.Value = 0;

                

                //現場手当
                //menkyo.Enabled = true;

                kyuuka.Enabled = true;
                //とりあえず現行
                kyuuka.SelectedItem = kyuuka_m.Text;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                //役職
                yakusyoku.Items.Add("0180" + "　" + "係員");
                yakusyoku.Items.Add("0170" + "　" + "サブチーフ");
                yakusyoku.Items.Add("0160" + "　" + "チーフ");
                yakusyoku.SelectedIndex = 0;

                //職種
                comboBoxSyoku.SelectedIndex = -1;
                comboBoxSyoku.Enabled = false;

                //学歴
                comboBoxGaku.SelectedIndex = -1;
                comboBoxGaku.Enabled = false;

                //社外経験
                comboBoxKeiken.SelectedIndex = -1;
                comboBoxKeiken.Enabled = false;

                //基準外額
                kizyungai.Text = "0";
                hyoukakyuu.Text = "0";

            }
            else if (kyuu == "F")
            {
                //社員区分設定
                //kyuuzitsukubun.Text = "04　アルバイト";

                //等級の設定
                shiyou.SelectedItem = "";

                //現場手当
                //menkyo.Enabled = false;

                //職務技能給
                syokumu.Value = 0;

                //時給・日給の設定
                zikyuu.Enabled = true;
                nikkyuu.Enabled = false;
                nikkyuu.Text = "0";

                //付与無し固定にする
                kyuuka.SelectedIndex = kyuuka.FindString("9");
                kyuuka.Enabled = false;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("2　アルバイト加入");

                //役職
                yakusyoku.Items.Add("0180" + "　" + "係員");
                yakusyoku.SelectedIndex = 0;

                //職種
                comboBoxSyoku.SelectedIndex = -1;
                comboBoxSyoku.Enabled = false;

                //学歴
                comboBoxGaku.SelectedIndex = -1;
                comboBoxGaku.Enabled = false;

                //社外経験
                comboBoxKeiken.SelectedIndex = -1;
                comboBoxKeiken.Enabled = false;

                //基準外額
                kizyungai.Text = "0";
                hyoukakyuu.Text = "0";
            }
            else
            {
                MessageBox.Show("Error:社員区分で想定外　システム管理者へ連絡願います。");
            }

            //友の会費連動
            tomonokaiChange();

            //表示・非表示
            if (kyuuyo.SelectedItem?.ToString() == kyuuyo_m.Text)
            {
                //kyuuyo_m.ForeColor = kyuuyo_m.BackColor;
                lbl_kyuuyo.BackColor = Color.White;
            }
            else
            {
                //kyuuyo_m.ForeColor = Color.Black;
                lbl_kyuuyo.BackColor = Color.LightGreen;
            }

            //表示・非表示
            if (kyuuzitsukubun.Text == kyuuzitsu_m.Text)
            {
                //syain_m.ForeColor = syain_m.BackColor;
                lbl_kyuuzitsu.BackColor = Color.White;
            }
            else
            {
                //syain_m.ForeColor = Color.Black;
                lbl_kyuuzitsu.BackColor = Color.LightGreen;
            }

            //発令区分
            //TODO 発令整理20200330
            //GetHatsurei();

            Disp();

            DataDispShikaku();
            RitouCalc();
        }


        //役職
        private void yakusyoku_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (kyuuyo.SelectedIndex == -1) return;

            //表示・非表示
            if (yakusyoku.SelectedItem?.ToString() == yakusyoku_m.Text)
            {
                lbl_yakusyoku.BackColor = Color.White;
            }
            else
            {
                lbl_yakusyoku.BackColor = Color.LightGreen;

            }

            string str = kyuuyo.SelectedItem.ToString().Substring(0, 2);
            string yaku = yakusyoku.SelectedItem.ToString().Substring(0, 4);

            //副部長と部長選択場合の職種固定設定
            if (Convert.ToInt16(yaku) == 110 || Convert.ToInt16(yaku) == 120)
            {
                //職種固定
                comboBoxSyoku.SelectedItem = "総合職";
                comboBoxSyoku.Enabled = false;
            }
            else
            {
                comboBoxSyoku.Enabled = true;
            }

            //職務技能給根拠パネルの表示・非表示
            if (Convert.ToInt16(yaku) >= 110)
            {
                if (str == "C1")
                {

                    DataTable hkdt = new DataTable();
                    hkdt = Com.GetDB("select * from dbo.HK_本給 where '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

                    foreach (DataRow row in hkdt.Rows)
                    {
                        if (keiyaku.SelectedItem?.ToString() != "")
                        {
                            honkyuu.Value = Convert.ToDecimal(honkyuu_m.Text);
                        }
                        else if (Convert.ToInt16(yaku) == 110) //部長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_部長"].ToString());

                        }
                        else if (Convert.ToInt16(yaku) == 120) //副部長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_副部長"].ToString());
                        }
                        else if (Convert.ToInt16(yaku) == 130) //課長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_課長"].ToString());
                        }
                        else if (Convert.ToInt16(yaku) == 135) //係長
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_係長"].ToString());
                        }
                        else if (Convert.ToInt16(yaku) == 140) //主任
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_主任"].ToString());
                        }
                        else if (Convert.ToInt16(yaku) == 150) //副主任
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString()) + Convert.ToDecimal(row["役職給_副主任"].ToString());
                        }
                        else
                        {
                            honkyuu.Value = Convert.ToDecimal(row["本給"].ToString());
                        }
                    }
                }
                else
                {
                    //TODO 20200318コメントアウト
                    //syokumuginoupanel.Visible = false;
                }

                //if (keiyaku.SelectedItem?.ToString() == "10　一般契約社員" || keiyaku.SelectedItem?.ToString() == "20　単年契約社員")
                if (keiyaku.SelectedItem?.ToString() != "")
                {
                    //2023/10/30 最賃対応　単年契約は本給もさわれるように設定
                    honkyuu.Enabled = true;
                    syokumu.Enabled = true;
                }
                else
                { 
                    honkyuu.Enabled = false;
                    syokumu.Enabled = false;
                }


            }
            else
            {
                honkyuu.Enabled = true;
                syokumu.Enabled = true;
                //menkyo.Enabled = true;
                //TODO 20200318コメントアウト
                //syokumuginoupanel.Visible = false;

                honkyuu.Value = Convert.ToDecimal(honkyuu_m.Text);
            }

            //役職手当の設定
            if (yaku == "0140") //係員
            {
                yakuteate.Text = "20000";
            }
            else if (yaku == "0150")//副主任
            {
                yakuteate.Text = "5000";
            }
            else if (yaku == "0160")//チーフ
            {
                yakuteate.Text = "5000";
            }
            else if (yaku == "0170")//サブチーフ
            {
                yakuteate.Text = "3000";
            }
            else
            {
                yakuteate.Text = "0";
            }




            //等級の設定
            //if (str == "B1" || str == "C1" || str == "C2")
            //{
            //}

            //TODO 発令整理20200330
            //GetHatsurei();
            //}

            DataDispShikaku();
            RitouCalc();

        }

        //契約社員
        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (keiyaku.SelectedItem?.ToString() == keiyaku_m.Text)
            {
                //keiyaku_m.ForeColor = keiyaku_m.BackColor;
                lbl_keiyaku.BackColor = Color.White;
            }
            else
            {
                //keiyaku_m.ForeColor = Color.Black;
                lbl_keiyaku.BackColor = Color.LightGreen;
            }

            //職務技能給根拠パネルの表示・非表示
            //if (keiyaku.SelectedItem?.ToString() == "10　一般契約社員" || keiyaku.SelectedItem?.ToString() == "20　単年契約社員")
            if (keiyaku.SelectedItem?.ToString() != "")
            {
                honkyuu.Enabled = true;
                syokumu.Enabled = true;
            }
            else
            {
                string yaku = yakusyoku.SelectedItem?.ToString().Substring(0, 4);

                //職務技能給根拠パネルの表示・非表示
                if (Convert.ToInt16(yaku) > 135)
                {
                    honkyuu.Enabled = false;
                    syokumu.Enabled = false;
                }
                else
                {
                    honkyuu.Enabled = true;
                    syokumu.Enabled = true;
                }
            }
        }

        private void tomonokaiChange()
        {
            if (kyuuyo.SelectedIndex == -1) return;

            string str = kyuuyo.SelectedItem.ToString().Substring(0, 2);
            if (str == "A1")
            {
                tomonokai.Text = "0";
            }
            else if (str == "F1")
            {
                //アルバイト
                if (tomokubun.SelectedItem?.ToString() == "2　アルバイト加入")
                {
                    tomonokai.Text = "300";
                }
                else
                {
                    tomonokai.Text = "0";
                }
            }
            else
            {
                //アルバイト以外
                if (tomokubun.SelectedItem?.ToString() == "1　非加入")
                {
                    tomonokai.Text = "0";

                }
                else
                {
                    tomonokai.Text = "300";
                }
            }
        }

        //友の会区分
        private void tomokubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tomokubun.SelectedItem?.ToString() == tomokubun_m.Text)
            {
                //tomokubun_m.ForeColor = tomokubun_m.BackColor;
                lbl_tomo.BackColor = Color.White;
            }
            else
            {
                //tomokubun_m.ForeColor = Color.Black;
                lbl_tomo.BackColor = Color.LightGreen;
            }

            tomonokaiChange();
        }

        //休日区分kyuuzitsukubun
        private void syainkubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kyuuzitsukubun.SelectedItem?.ToString() == kyuuzitsukubun.Text)
            {
                //syain_m.ForeColor = syain_m.BackColor;
                lbl_kyuuzitsu.BackColor = Color.White;
            }
            else
            {
                //syain_m.ForeColor = Color.Black;
                lbl_kyuuzitsu.BackColor = Color.LightGreen;
            }
        }

        //試用期間
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            //等級　表示・非表示
            if (shiyou.SelectedItem?.ToString() == shiyou_m.Text)
            {
                lbl_shiyou.BackColor = Color.White;
            }
            else
            {
                lbl_shiyou.BackColor = Color.LightGreen;
            }
        }

        private void tuukinteatekubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tuukinteatekubun.SelectedItem?.ToString() == tuukinteatekubun_m.Text)
            {
                lbl_tuukinteatekubun.BackColor = Color.White;
            }
            else
            {
                lbl_tuukinteatekubun.BackColor = Color.LightGreen;
            }

            tuukinCalc();
        }


        //通勤手当区分
        private void tuukinkubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tuukinkubun.SelectedItem?.ToString() == tuukinkubun_m.Text)
            {
                lbl_tuukinkubun.BackColor = Color.White;
            }
            else
            {
                lbl_tuukinkubun.BackColor = Color.LightGreen;
            }

            tuukinCalc();

        }

        //通勤手当算出
        private void tuukinCalc()
        {
            decimal flg = 0;
            decimal kotei = 0;

            if (tuukinteatekubun.SelectedItem?.ToString() == "1 実費精算")
            {
                tuutanka.Text = "0";
                tuukinhi2.Text = "0";
                tuukinka2.Text = "0";

                return;
            }

            switch (tuukinkubun.SelectedItem?.ToString())
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
            decimal tanka = katakyori.Value * 30 * flg + kotei;

            //40overの場合の単価
            if (katakyori.Value > 40) tanka = 40 * 30 * flg + kotei;

            //概算通勤手当総額
            decimal dec = tanka * Convert.ToDecimal(Getrday(kyuuka.SelectedItem?.ToString()));


            //通勤1日単価
            tuutanka.Text = katakyori.Value < 1 ? "0" : Convert.ToInt32(tanka).ToString();

            if (katakyori.Value < 1)
            {
                tuukinhi2.Text = "0";
                tuukinka2.Text = "0";
            }
            else if (katakyori.Value < 2) //0
            {　
                tuukinhi2.Text = "0";
                tuukinka2.Text = Convert.ToInt32(dec).ToString();
            }
            else if (katakyori.Value < 10) //4200
            {
                tuukinhi2.Text = dec > 4200 ? "4200" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 4200 ? (Convert.ToInt32(dec) - 4200).ToString() : "0";
            }
            else if (katakyori.Value < 15)
            {
                tuukinhi2.Text = dec > 7100 ? "7100" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 7100 ? (Convert.ToInt32(dec) - 7100).ToString() : "0";
            }
            else if (katakyori.Value < 25)
            {
                tuukinhi2.Text = dec > 12900 ? "12900" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 12900 ? (Convert.ToInt32(dec) - 12900).ToString() : "0";
            }
            else if (katakyori.Value < 35)
            {
                tuukinhi2.Text = dec > 18700 ? "18700" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 18700 ? (Convert.ToInt32(dec) - 18700).ToString() : "0";
            }
            else if (katakyori.Value < 45)
            {
                tuukinhi2.Text = dec > 24400 ? "24400" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 24400 ? (Convert.ToInt32(dec) - 24400).ToString() : "0";
            }
            else if (katakyori.Value < 55)
            {
                tuukinhi2.Text = dec > 28000 ? "28000" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 28000 ? (Convert.ToInt32(dec) - 28000).ToString() : "0";
            }
            else
            {
                tuukinhi2.Text = dec > 31600 ? "31600" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = dec > 31600 ? (Convert.ToInt32(dec) - 31600).ToString() : "0";
            }

            //バス、モノレールは全額非課税に
            if (tuukinkubun.SelectedItem?.ToString() == "4 バス・モノレール")
            {
                tuukinhi2.Text = katakyori.Value < 1 ? "0" : Convert.ToInt32(dec).ToString();
                tuukinka2.Text = "0";
            }
        }

        //通勤距離
        private void katakyori_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (katakyori.Value.ToString() == katakyori_m.Text)
            {
                lbl_katakyori.BackColor = Color.White;
            }
            else
            {
                lbl_katakyori.BackColor = Color.LightGreen;
            }

            tuukinCalc();
           
        }


        //姓
        private void sei_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (sei.Text == sei_m.Text)
            {
                //sei_m.ForeColor = sei_m.BackColor;
                lbl_sei.BackColor = Color.White;
            }
            else
            {
                //sei_m.ForeColor = Color.Black;
                lbl_sei.BackColor = Color.LightGreen;
            }
        }

        //姓フリガナ
        private void seihuri_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (seihuri.Text == seihuri_m.Text)
            {
                //seihuri_m.ForeColor = seihuri_m.BackColor;
                lbl_seihuri.BackColor = Color.White;
            }
            else
            {
                //seihuri_m.ForeColor = Color.Black;
                lbl_seihuri.BackColor = Color.LightGreen;
            }
        }

        //郵便番号
        private void yuubin_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (yuubin.Text == yuubin_m.Text)
            {
                //yuubin_m.ForeColor = yuubin_m.BackColor;
                lbl_yuubin.BackColor = Color.White;
            }
            else
            {
                //yuubin_m.ForeColor = Color.Black;
                lbl_yuubin.BackColor = Color.LightGreen;
            }
        }

        //現住所
        private void zyuusyo_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (zyuusyo.Text == zyuusyo_m.Text)
            {
                //zyuusyo_m.ForeColor = zyuusyo_m.BackColor;
                lbl_zyuusyo.BackColor = Color.White;
            }
            else
            {
                //zyuusyo_m.ForeColor = Color.Black;
                lbl_zyuusyo.BackColor = Color.LightGreen;
            }
        }



        private void tokubetsu_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tokubetsu.Value.ToString() == tokubetsu_m.Text)
            {
                //tokubetsu_m.ForeColor = tokubetsu_m.BackColor;
                lbl_tokubetsu.BackColor = Color.White;
            }
            else
            {
                //tokubetsu_m.ForeColor = Color.Black;
                lbl_tokubetsu.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void yakuteate_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (yakuteate.Text == yakuteate_m.Text)
            {
                //yakuteate_m.ForeColor = yakuteate_m.BackColor;
                lbl_yakuteate.BackColor = Color.White;
            }
            else
            {
                //yakuteate_m.ForeColor = Color.Black;
                lbl_yakuteate.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void genbateate_ValueChanged(object sender, EventArgs e)
        {
            ////表示・非表示
            //if (menkyo.Text == genbateate_m.Text)
            //{
            //    //genbateate_m.ForeColor = genbateate_m.BackColor;
            //    lbl_genbateate.BackColor = Color.White;
            //}
            //else
            //{
            //    //genbateate_m.ForeColor = Color.Black;
            //    lbl_genbateate.BackColor = Color.LightGreen;
            //}

            //Disp();
        }

        private void huyou_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (huyou.Text == huyou_m.Text)
            {
                //huyou_m.ForeColor = huyou_m.BackColor;
                lbl_huyou.BackColor = Color.White;
            }
            else
            {
                //huyou_m.ForeColor = Color.Black;
                lbl_huyou.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void syukkou_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syukkou.Value.ToString() == syukkou_m.Text)
            {
                //syukkou_m.ForeColor = syukkou_m.BackColor;
                lbl_syukkou.BackColor = Color.White;
            }
            else
            {
                //syukkou_m.ForeColor = Color.Black;
                lbl_syukkou.BackColor = Color.LightGreen;
            }

            Disp();
            RitouCalc();
        }

        private void tuukinhi_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tuukinhi.Text == tuukinhi_m.Text)
            {
                lbl_tuukinhi.BackColor = Color.White;
            }
            else
            {
                lbl_tuukinhi.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void tuukinka_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tuukinka.Text == tuukinka_m.Text)
            {
                //tuukinka_m.ForeColor = tuukinka_m.BackColor;
                lbl_tuukinka.BackColor = Color.White;
            }
            else
            {
                //tuukinka_m.ForeColor = Color.Black;
                lbl_tuukinka.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void touroku_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (touroku.Text == touroku_m.Text)
            {
                lbl_touroku.BackColor = Color.White;
            }
            else
            {
                lbl_touroku.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void tomonokai_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tomonokai.Text == tomonokai_m.Text)
            {
                //tomonokai_m.ForeColor = tomonokai_m.BackColor;
                lbl_tomonokai.BackColor = Color.White;
            }
            else
            {
                //tomonokai_m.ForeColor = Color.Black;
                lbl_tomonokai.BackColor = Color.LightGreen;
            }

            Disp();
        }

        //private void kotei1_ValueChanged(object sender, EventArgs e)
        //{
        //    //表示・非表示
        //    if (kotei1.Value.ToString() == kotei1_m.Text)
        //    {
        //        //kotei1_m.ForeColor = kotei1_m.BackColor;
        //        lbl_kotei1.BackColor = Color.White;
        //    }
        //    else
        //    {
        //        //kotei1_m.ForeColor = Color.Black;
        //        lbl_kotei1.BackColor = Color.LightGreen;
        //    }

        //    Disp();
        //}

        //private void kotei2_ValueChanged(object sender, EventArgs e)
        //{
        //    //表示・非表示
        //    if (kotei2.Value.ToString() == kotei2_m.Text)
        //    {
        //        //kotei2_m.ForeColor = kotei2_m.BackColor;
        //        lbl_kotei2.BackColor = Color.White;
        //    }
        //    else
        //    {
        //        //kotei2_m.ForeColor = Color.Black;
        //        lbl_kotei2.BackColor = Color.LightGreen;
        //    }

        //    Disp();
        //}

        private void zeikubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (zeikubun.SelectedItem?.ToString() == zeikubun_m.Text)
            {
                //zeikubun_m.ForeColor = zeikubun_m.BackColor;
                label34.BackColor = Color.White;
            }
            else
            {
                //zeikubun_m.ForeColor = Color.Black;
                label34.BackColor = Color.LightGreen;
            }
        }

        private void syougai_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syougai.SelectedItem?.ToString() == syougai_m.Text)
            {
                //syougai_m.ForeColor = syougai_m.BackColor;
                label35.BackColor = Color.White;
            }
            else
            {
                //syougai_m.ForeColor = Color.Black;
                label35.BackColor = Color.LightGreen;
            }
        }

        private void kahu_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kahu.SelectedItem?.ToString() == kahu_m.Text)
            {
                //kahu_m.ForeColor = kahu_m.BackColor;
                label38.BackColor = Color.White;
            }
            else
            {
                //kahu_m.ForeColor = Color.Black;
                label38.BackColor = Color.LightGreen;
            }
        }

        private void kinrou_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kinrou.SelectedItem?.ToString() == kinrou_m.Text)
            {
                //kinrou_m.ForeColor = kinrou_m.BackColor;
                label39.BackColor = Color.White;
            }
            else
            {
                //kinrou_m.ForeColor = Color.Black;
                label39.BackColor = Color.LightGreen;
            }
        }

        private void gaikoku_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (gaikoku.SelectedItem?.ToString() == gaikoku_m.Text)
            {
                //gaikoku_m.ForeColor = gaikoku_m.BackColor;
                label48.BackColor = Color.White;
            }
            else
            {
                //gaikoku_m.ForeColor = Color.Black;
                label48.BackColor = Color.LightGreen;
            }
        }

        private void saigai_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (saigai.SelectedItem?.ToString() == saigai_m.Text)
            {
                //saigai_m.ForeColor = saigai_m.BackColor;
                label49.BackColor = Color.White;
            }
            else
            {
                //saigai_m.ForeColor = Color.Black;
                label49.BackColor = Color.LightGreen;
            }
        }


        private void nikkyuu_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kaisuu1_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kaisuu1.Value.ToString() == kaisuu1_m.Text)
            {
                //kaisuu1_m.ForeColor = kaisuu1_m.BackColor;
                label28.BackColor = Color.White;
            }
            else
            {
                //kaisuu1_m.ForeColor = Color.Black;
                label28.BackColor = Color.LightGreen;
            }
        }

        //private void kaisuu2_ValueChanged(object sender, EventArgs e)
        //{
        //    //表示・非表示
        //    if (kaisuu2.Value.ToString() == kaisuu2_m.Text)
        //    {
        //        //kaisuu2_m.ForeColor = kaisuu2_m.BackColor;
        //        label29.BackColor = Color.White;
        //    }
        //    else
        //    {
        //        //kaisuu2_m.ForeColor = Color.Black;
        //        label29.BackColor = Color.LightGreen;
        //    }
        //}

        private void kyuuka_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kyuuka.SelectedItem?.ToString() == kyuuka_m.Text)
            {
                //kyuuka_m.ForeColor = kyuuka_m.BackColor;
                label30.BackColor = Color.White;
            }
            else
            {
                //kyuuka_m.ForeColor = Color.Black;
                label30.BackColor = Color.LightGreen;
            }

            //日給者・パート・アルバイトの本給に暫定給を表示
            double hon = 0;
            if (zikyuu.Text != "0")
            {
                hon = Math.Round(Convert.ToInt32(zikyuu.Value) * Convert.ToInt32(kinmu.Text) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;
            }
            if (nikkyuu.Text != "0")
            {
                hon = Math.Round(Convert.ToInt32(nikkyuu.Value) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;
            }

            tuukinCalc();
        }

        private void kinmu_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kinmu.SelectedItem?.ToString() == kinmu_m.Text)
            {
                //kinmu_m.ForeColor = kinmu_m.BackColor;
                label31.BackColor = Color.White;
            }
            else
            {
                //kinmu_m.ForeColor = Color.Black;
                label31.BackColor = Color.LightGreen;
            }

            //日給者・パート・アルバイトの本給に暫定給を表示
            double hon = 0;
            if (zikyuu.Text != "0")
            {
                hon = Math.Round(Convert.ToInt32(zikyuu.Value) * Convert.ToInt32(kinmu.Text) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;
            }
            if (nikkyuu.Text != "0")
            {
                hon = Math.Round(Convert.ToInt32(nikkyuu.Value) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;
            }
        }

        private void kizyun_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (kizyun.Text == kizyun_m.Text)
            {
                //kizyun_m.ForeColor = kizyun_m.BackColor;
                lbl_kizyun.BackColor = Color.White;
            }
            else
            {
                //kizyun_m.ForeColor = Color.Black;
                lbl_kizyun.BackColor = Color.LightGreen;
            }
        }

        private void shikyuu_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (shikyuu.Text == shikyuu_m.Text)
            {
                //shikyuu_m.ForeColor = shikyuu_m.BackColor;
                lbl_shikyuu.BackColor = Color.White;
            }
            else
            {
                //shikyuu_m.ForeColor = Color.Black;
                lbl_shikyuu.BackColor = Color.LightGreen;
            }
        }

        //private void kouzyo_TextChanged(object sender, EventArgs e)
        //{
        //    //表示・非表示
        //    if (kouzyo.Text == kouzyo_m.Text)
        //    {
        //        //kouzyo_m.ForeColor = kouzyo_m.BackColor;
        //        lbl_kouzyo.BackColor = Color.White;
        //    }
        //    else
        //    {
        //        //kouzyo_m.ForeColor = Color.Black;
        //        lbl_kouzyo.BackColor = Color.LightGreen;
        //    }
        //}

        #endregion

        private void dispdgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //ヘッダは対象外
            if (dispdgv.CurrentCell != null)
            {
                //リセット
                DataReset();

                DataGridViewRow dgr = dispdgv.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;

                //Noを反映
                lbl_no.Text = drv[0].ToString();

                //異動日を反映
                idoudaynew.Value = Convert.ToDateTime(drv[1].ToString());

                //現行データを取得
                GetData(drv[2].ToString());

                //異動データを取得
                GetDataIdou(drv[0].ToString(), drv[2].ToString());

                //離島処理
                RitouCalc();

                //労働条件取得
                DataDispRoudou();
            }
        }

        private void shikakubtn_Click(object sender, EventArgs e)
        {
            //ボタンが新規登録で、既に登録されている資格コードが選択された場合はすでにあるよー表示にする
            if (shikakubtn.Text == "資格登録")
            {
                //資格コードのみにする
                string[] del = { "　" };
                string[] shikakucd = shikakutextb.Text.Split(del, StringSplitOptions.None);

                DataTable dt = new DataTable();
                dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "' and 資格コード = '" + shikakucd[0] + "' and 適用終了日 > '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

                if (dt.Rows.Count == 1)
                {
                    MessageBox.Show(shikakutextb.Text + nl + "もう登録されてます。");
                    return;
                }
            }

            if (shikakutextb.Text == "" || shikakusyutokubi.Value == DateTime.Today || shikakuno.Text == "" || (shikakukigenday.Value == DateTime.Today && radioButton1.Checked == true))
            {
                string msg = "";
                if (shikakutextb.Text == "") msg += "資格コードは必須です。" + nl;
                //TODO 変更しよー
                if (shikakusyutokubi.Value == DateTime.Today) msg += "この資格、今日とったー？" + nl;
                if (shikakuno.Text == "") msg += "資格取得番号は必須です。" + nl;
                if ((shikakukigenday.Value == DateTime.Today && radioButton1.Checked == true)) msg += "この資格、今日有効期限切れ？" + nl;
                MessageBox.Show(msg);
                return;
            }

            //登録or更新
            SetShikaku();

            //登録後のフォーカスに利用
            //TODO 一覧表示は登録順または追加は必ず最後の行に入らなければならない！
            if (shikakudgv.CurrentCell == null)
            {

                dgvRow = shikakudgv.Rows.Count;
            }
            else
            {
                dgvRow = shikakudgv.CurrentCell.RowIndex;
            }

            //一覧データ取得
            DataDispShikaku();

            //登録の場合
            if (shikakubtn.Text != "資格更新")
            {
                dgvRow = shikakudgv.Rows.Count - 1;
            }

            if (shikakudgv.Rows.Count > 0)
            {
                shikakudgv.CurrentCell = shikakudgv[1, dgvRow];
            }

            //クリア
            ShikakuClear();

            DataGridViewRow dgr = shikakudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            //TODO:0でOK？　下にも同じ処理ある
            GetShikakukoumokuData(drv[0].ToString());

            if (shikakubtn.Text == "資格更新")
            {
                MessageBox.Show("資格情報を更新しました。");

            }
            else
            {
                //TODO　登録更新した情報を表示
                MessageBox.Show("登録情報を登録しました。");
            }

            //全体登録更新
            TourokuKoushin();
        }

        //資格情報登録更新
        private void SetShikaku()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[資格データ登録更新]";

                    Cmd.Parameters.Add(new SqlParameter("会社コード", SqlDbType.VarChar)); Cmd.Parameters["会社コード"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("適用開始日", SqlDbType.Char)); Cmd.Parameters["適用開始日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格コード", SqlDbType.VarChar)); Cmd.Parameters["資格コード"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("個人識別ＩＤ", SqlDbType.VarChar)); Cmd.Parameters["個人識別ＩＤ"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("適用終了日", SqlDbType.Char)); Cmd.Parameters["適用終了日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格取得日", SqlDbType.VarChar)); Cmd.Parameters["資格取得日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格認定日", SqlDbType.VarChar)); Cmd.Parameters["資格認定日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格取得番号", SqlDbType.VarChar)); Cmd.Parameters["資格取得番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("手当支給区分", SqlDbType.Decimal)); Cmd.Parameters["手当支給区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("資格有効期限", SqlDbType.Char)); Cmd.Parameters["資格有効期限"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新日時", SqlDbType.DateTime)); Cmd.Parameters["更新日時"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新ユーザＩＤ", SqlDbType.Decimal)); Cmd.Parameters["更新ユーザＩＤ"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar)); Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["会社コード"].Value = "E0";
                    Cmd.Parameters["社員番号"].Value = syainno.Text;
                    Cmd.Parameters["適用開始日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

                    //資格コードのみにする
                    string[] del = { "　" };
                    string[] shikakucd = shikakutextb.Text.Split(del, StringSplitOptions.None);
                    Cmd.Parameters["資格コード"].Value = shikakucd[0];
                    Cmd.Parameters["個人識別ＩＤ"].Value = kozinshikibetsuid.Text;
                    Cmd.Parameters["適用終了日"].Value = "9999/12/31";
                    Cmd.Parameters["資格取得日"].Value = shikakusyutokubi.Value.ToString("yyyy/MM/dd");
                    Cmd.Parameters["資格認定日"].Value = shikakusyutokubi.Value.ToString("yyyy/MM/dd");
                    Cmd.Parameters["資格取得番号"].Value = shikakuno.Text;
                    Cmd.Parameters["手当支給区分"].Value = "0";
                    if (radioButton1.Checked)
                    { 
                        Cmd.Parameters["資格有効期限"].Value = shikakukigenday.Value.ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        Cmd.Parameters["資格有効期限"].Value = "";
                    }
                    Cmd.Parameters["更新日時"].Value = DateTime.Today;
                    Cmd.Parameters["更新ユーザＩＤ"].Value = "24"; //
                    Cmd.Parameters["更新者"].Value = Program.loginname;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void DataDispShikaku()
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "' and 適用終了日 > '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");
            shikakudgv.DataSource = dt;

            string k = kyuuyo.SelectedItem.ToString().Substring(0, 2); //給与支給区分
            string s = comboBoxSyoku.SelectedItem?.ToString(); //職種
            string y = yakusyoku.Text.Substring(0, 4); //役職
            string h = kyuuka.Text.Substring(0, 1); //休暇付与区分
            int d = Convert.ToInt16(kinmu.SelectedItem); //勤務時間

            if (dt.Rows.Count != 0)
            {
                //TODO カラム幅
                shikakudgv.Columns[0].Width = 100;
                shikakudgv.Columns[1].Width = 200;
                shikakudgv.Columns[2].Width = 100;
                shikakudgv.Columns[3].Width = 150;
                shikakudgv.Columns[4].Width = 100;
                shikakudgv.Columns[5].Width = 100;

                shikakudgv.Columns["社員番号"].Visible = false;
                shikakudgv.Columns["個人識別ＩＤ"].Visible = false;
                shikakudgv.Columns["規程額"].DefaultCellStyle.Format = "#,0";
                shikakudgv.Columns["規程額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //免許手当総額を取得
                DataTable dt2 = new DataTable();

                dt2 = Com.GetDB("select * from dbo.m免許手当合計規程額取得_異動('" + k + "', '" + s + "', '" + y + "', '" + h + "', '" + d + "') where 社員番号 = '" + syainno.Text + "'");

                //null対応
                if (dt2.Rows.Count == 0)
                {
                    menkyoato.Text = "0";
                }
                else
                {
                    if (dt2.Rows[0][0].ToString() == "")
                    {
                        menkyoato.Text = "0";
                    }
                    else
                    {
                        menkyoato.Text = dt2.Rows[0][0].ToString().Replace(".000000", "");
                    }
                }
            }

            //登録手当総額を取得
            DataTable dt3 = new DataTable();
            dt3 = Com.GetDB("select sum(手当対象額) from dbo.t登録手当一覧_異動('" + y + "') where 社員番号 = '" + syainno.Text + "'");

            //null対応
            if (dt3.Rows.Count == 0)
            {
                tourokuato.Text = "0";
            }
            else
            {
                if (dt3.Rows[0][0].ToString() == "")
                {
                    tourokuato.Text = "0";
                }
                else
                {
                    tourokuato.Text = dt3.Rows[0][0].ToString().Replace(".000000", "");
                }
                
            }

            //資格情報一覧
            shikakudgv.CurrentCell = null;
        }

        private void ShikakuClear()
        {
            shikakutextb.Text = "";
            shikakusyutokubi.Text = DateTime.Today.ToString();
            shikakuno.Text = "";
            shikakubtn.Text = "資格登録";

            //登録削除ボタンの無効化
            Tourokusakuzyobtn.Enabled = false;

            shikakubtn.BackColor = Color.Transparent;
            shikakubtn.ForeColor = Color.Black;

            //shikakucombo.Enabled = true;

            shikakukigenday.Text = DateTime.Today.ToString();
            radioButton2.Checked = true;

            //TODO 削除ボタンまだない
            //delshikaku.Visible = false;
        }

        private void GetShikakukoumokuData(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "' and 資格コード = '" + str + "' and 適用終了日 > '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

            shikakutextb.Text = dt.Rows[0][0].ToString() + '　' + dt.Rows[0][1].ToString();
            shikakusyutokubi.Text = dt.Rows[0][2].ToString();
            shikakuno.Text = dt.Rows[0][3].ToString();

            if (dt.Rows[0][9].ToString() == "必須")
            {
                radioButton1.Checked = true;
                shikakukigenday.Text = dt.Rows[0][4].ToString();
            }
            else
            {
                radioButton2.Checked = true;
                //TODO 
                //shikakukigenday.Text = ""; //資格有効期限
            }
        }

        private void shikakudgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //ヘッダは対象外
            if (shikakudgv.CurrentCell != null)
            {
                ShikakuClear();

                DataGridViewRow dgr = shikakudgv.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                GetShikakukoumokuData(drv[0].ToString());
            }

            //TODO 資格登録更新ボタン表示の変更
            shikakubtn.Text = "資格更新";
            shikakubtn.BackColor = Color.Blue;
            shikakubtn.ForeColor = Color.White;

            //登録削除ボタンの有効化
            Tourokusakuzyobtn.Enabled = true;


            //更新の時は免許の変更はできない
            //shikakucombo.Enabled = false;

            //TODO まだ作成されてない
            //delshikaku.Visible = true;
        }

        private void kazokubtn_Click(object sender, EventArgs e)
        {
            //必須項目チェック
            if (kazosei.Text.Trim() == "" || kazomei.Text.Trim() == "" || kazokanasei.Text.Trim() == "" || kazokanamei.Text.Trim() == "")
            {
                MessageBox.Show("名前必須です");
                return;
            }

            if (zokugara.Text == "")
            {
                MessageBox.Show("続柄必須です");
                return;
            }

            if (kazoseinengappinew.Text == "")
            {
                MessageBox.Show("生年月日必須です");
                return;
            }


            //登録or更新
            SetKazoku();

            if (kazokubtn.Text == "家族更新")
            {
                //更新時の処理
                dgvRow = kazokudgv.CurrentCell.RowIndex;

                //一覧データ取得
                DataDispKazoku();

                kazokudgv.CurrentCell = kazokudgv[1, dgvRow];

                DataGridViewRow dgr = kazokudgv.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;

                //TODO:0でOK？　下にも同じ処理ある
                GetKazokukoumokuData(drv[0].ToString());

                MessageBox.Show("家族情報を更新しました。");

            }
            else
            {
                //一覧データ取得
                DataDispKazoku();

                //選択無
                kazokudgv.CurrentCell = null;

                //項目クリア
                KazokuClear();

                MessageBox.Show("家族情報を登録しました。");
            }

            //全体登録更新
            TourokuKoushin();
            //DataInsertUpdate();
        }

        //家族情報登録更新
        private void SetKazoku()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;

                    //①追加　　　　家族識別IDがはいってない
                    //②更新と追加　適用終了日が'9999/12/31'データの適用開始日が異動日より過去
                    //③更新　　　　適用終了日が'9999/12/31'データの適用開始日が異動日である
                    //④例外　　　　適用終了日が'9999/12/31'データの適用開始日が異動日より未来

                    DataTable flgdata = Com.GetDB("select 適用開始日 from QUATRO.dbo.SJMTKAZOKU where 会社コード = 'E0' and 社員番号 = '" + syainno.Text + "' and 適用終了日 = '9999/12/31' and 家族識別ＩＤ = '" + this.kazokuid.Text + "'");

                    //家族IDが入ってないor適用開始日が異動日より若い
                    if (this.kazokuid.Text == "" || Convert.ToDateTime(flgdata.Rows[0][0]) < Convert.ToDateTime(idoudaynew.Value))
                    {
                        if (this.kazokuid.Text != "")
                        {
                            //②の更新
                            DataTable update = Com.GetDB("update QUATRO.dbo.SJMTKAZOKU set 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "', ユーザ任意フィールド１ = '02_異動入力_更新' where 会社コード = 'E0' and 社員番号 = '" + syainno.Text + "' and 適用終了日 = '9999/12/31' and 家族識別ＩＤ = '" + this.kazokuid.Text + "'");
                        }
                        //①追加　　　　家族識別IDがはいってない
                        //②の追加


                        Cmd.CommandText = "[dbo].[i異動家族データインサート]";
                        
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("個人識別ＩＤ", SqlDbType.VarChar)); Cmd.Parameters["個人識別ＩＤ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("家族識別ＩＤ", SqlDbType.VarChar)); Cmd.Parameters["家族識別ＩＤ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("姓", SqlDbType.VarChar)); Cmd.Parameters["姓"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("名", SqlDbType.VarChar)); Cmd.Parameters["名"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("カナ姓", SqlDbType.VarChar)); Cmd.Parameters["カナ姓"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("カナ名", SqlDbType.VarChar)); Cmd.Parameters["カナ名"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("生年月日", SqlDbType.Char)); Cmd.Parameters["生年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("続柄区分", SqlDbType.VarChar)); Cmd.Parameters["続柄区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("同居区分", SqlDbType.Decimal)); Cmd.Parameters["同居区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("配偶者", SqlDbType.Decimal)); Cmd.Parameters["配偶者"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("税扶養区分", SqlDbType.Decimal)); Cmd.Parameters["税扶養区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("源泉控除対象配偶者", SqlDbType.Decimal)); Cmd.Parameters["源泉控除対象配偶者"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("健保加入区分", SqlDbType.Decimal)); Cmd.Parameters["健保加入区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("世帯主区分", SqlDbType.Decimal)); Cmd.Parameters["世帯主区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("障害者区分", SqlDbType.Decimal)); Cmd.Parameters["障害者区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("特別障害者区分", SqlDbType.Decimal)); Cmd.Parameters["特別障害者区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("異動日", SqlDbType.VarChar)); Cmd.Parameters["異動日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("etc", SqlDbType.VarChar)); Cmd.Parameters["etc"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("直系尊属区分", SqlDbType.Decimal)); Cmd.Parameters["直系尊属区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("特定親族区分", SqlDbType.Decimal)); Cmd.Parameters["特定親族区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                        Cmd.Parameters["社員番号"].Value = syainno.Text;
                        Cmd.Parameters["個人識別ＩＤ"].Value = kozinshikibetsuid.Text;


                        if (this.kazokuid.Text == "")
                        { 
                            DataTable kazokuid = new DataTable();
                            kazokuid = Com.GetDB("select max(家族識別ＩＤ) from QUATRO.dbo.SJMTKAZOKU where 会社コード = 'E0' and 社員番号 = '" + syainno.Text + "'");
                            if (kazokuid.Rows[0][0] == DBNull.Value)
                            {
                                Cmd.Parameters["家族識別ＩＤ"].Value = 1;
                            }
                            else
                            {
                                Cmd.Parameters["家族識別ＩＤ"].Value = Convert.ToInt16(kazokuid.Rows[0][0]) + 1;
                            }
                        }
                        else
                        {
                            Cmd.Parameters["家族識別ＩＤ"].Value = kazokuid.Text;
                        }

                        Cmd.Parameters["姓"].Value = kazosei.Text;
                        Cmd.Parameters["名"].Value = kazomei.Text;
                        Cmd.Parameters["カナ姓"].Value = kazokanasei.Text;
                        Cmd.Parameters["カナ名"].Value = kazokanamei.Text;
                        Cmd.Parameters["生年月日"].Value = Convert.ToDateTime(kazoseinengappinew.Text).ToString("yyyy/MM/dd");
                        Cmd.Parameters["続柄区分"].Value = zokugara.Text.Substring(0, 2);
                        Cmd.Parameters["同居区分"].Value = doukyokubun.Text.Substring(0, 1);

                        if (zokugara.Text == "00　夫" || zokugara.Text == "01　妻")
                        {
                            Cmd.Parameters["配偶者"].Value = 1; //妻or夫の場合
                        }
                        else
                        {
                            Cmd.Parameters["配偶者"].Value = 0; //妻or夫の場合
                        }

                        Cmd.Parameters["税扶養区分"].Value = huyoukubun.Text.Substring(0, 1);
                        Cmd.Parameters["源泉控除対象配偶者"].Value = gensenkubun.Text.Substring(0, 1);
                        Cmd.Parameters["健保加入区分"].Value = kenpokanyuu.Text.Substring(0, 1);
                        Cmd.Parameters["世帯主区分"].Value = setainushi.Text.Substring(0, 1);

                        if (syougaikubun.Text == "0　該当しない")
                        {
                            Cmd.Parameters["障害者区分"].Value = 0;
                            Cmd.Parameters["特別障害者区分"].Value = 0; //障害区分を
                        }
                        else if (syougaikubun.Text == "2　特別")
                        {
                            Cmd.Parameters["障害者区分"].Value = 0;
                            Cmd.Parameters["特別障害者区分"].Value = 1;
                        }
                        else
                        {
                            Cmd.Parameters["障害者区分"].Value = 1;
                            Cmd.Parameters["特別障害者区分"].Value = 0;
                        }

                        Cmd.Parameters["異動日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

                        if (this.kazokuid.Text != "")
                        {
                            Cmd.Parameters["etc"].Value = "03_異動入力_更新インサート";
                        }
                        else
                        {
                            Cmd.Parameters["etc"].Value = "01_異動入力_新規インサート";
                        }

                        string kcode = zokugara.Text.Substring(0, 2);
                        if (kcode == "30" || kcode == "31" || kcode == "32" || kcode == "33" || kcode == "40" || kcode == "41")
                        {
                            Cmd.Parameters["直系尊属区分"].Value = 1;
                        }
                        else
                        {
                            Cmd.Parameters["直系尊属区分"].Value = 0;
                        }

                        //20251021 追加
                        Cmd.Parameters["特定親族区分"].Value = 0;

                        using (dr = Cmd.ExecuteReader())
                        {
                            int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        }
                    }
                    else if (Convert.ToDateTime(flgdata.Rows[0][0]) == Convert.ToDateTime(idoudaynew.Value))
                    {
                        //③更新　　適用開始日が異動日である
                        Cmd.CommandText = "[dbo].[i異動家族データアップデート]";

                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("個人識別ＩＤ", SqlDbType.VarChar)); Cmd.Parameters["個人識別ＩＤ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("家族識別ＩＤ", SqlDbType.VarChar)); Cmd.Parameters["家族識別ＩＤ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("姓", SqlDbType.VarChar)); Cmd.Parameters["姓"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("名", SqlDbType.VarChar)); Cmd.Parameters["名"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("カナ姓", SqlDbType.VarChar)); Cmd.Parameters["カナ姓"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("カナ名", SqlDbType.VarChar)); Cmd.Parameters["カナ名"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("生年月日", SqlDbType.Char)); Cmd.Parameters["生年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("続柄区分", SqlDbType.VarChar)); Cmd.Parameters["続柄区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("同居区分", SqlDbType.Decimal)); Cmd.Parameters["同居区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("配偶者", SqlDbType.Decimal)); Cmd.Parameters["配偶者"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("税扶養区分", SqlDbType.Decimal)); Cmd.Parameters["税扶養区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("源泉控除対象配偶者", SqlDbType.Decimal)); Cmd.Parameters["源泉控除対象配偶者"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("健保加入区分", SqlDbType.Decimal)); Cmd.Parameters["健保加入区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("世帯主区分", SqlDbType.Decimal)); Cmd.Parameters["世帯主区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("障害者区分", SqlDbType.Decimal)); Cmd.Parameters["障害者区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("特別障害者区分", SqlDbType.Decimal)); Cmd.Parameters["特別障害者区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("異動日", SqlDbType.VarChar)); Cmd.Parameters["異動日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("etc", SqlDbType.VarChar)); Cmd.Parameters["etc"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("直系尊属区分", SqlDbType.Decimal)); Cmd.Parameters["直系尊属区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("特定親族区分", SqlDbType.Decimal)); Cmd.Parameters["特定親族区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                        Cmd.Parameters["社員番号"].Value = syainno.Text;
                        Cmd.Parameters["個人識別ＩＤ"].Value = kozinshikibetsuid.Text;
                        Cmd.Parameters["家族識別ＩＤ"].Value = kazokuid.Text;

                        Cmd.Parameters["姓"].Value = kazosei.Text;
                        Cmd.Parameters["名"].Value = kazomei.Text;
                        Cmd.Parameters["カナ姓"].Value = kazokanasei.Text;
                        Cmd.Parameters["カナ名"].Value = kazokanamei.Text;
                        Cmd.Parameters["生年月日"].Value = Convert.ToDateTime(kazoseinengappinew.Text).ToString("yyyy/MM/dd");
                        Cmd.Parameters["続柄区分"].Value = zokugara.Text.Substring(0, 2);
                        Cmd.Parameters["同居区分"].Value = doukyokubun.Text.Substring(0, 1);

                        if (zokugara.Text == "00　夫" || zokugara.Text == "01　妻")
                        {
                            Cmd.Parameters["配偶者"].Value = 1; //妻or夫の場合
                        }
                        else
                        {
                            Cmd.Parameters["配偶者"].Value = 0; //妻or夫の場合
                        }

                        Cmd.Parameters["税扶養区分"].Value = huyoukubun.Text.Substring(0, 1);
                        Cmd.Parameters["源泉控除対象配偶者"].Value = gensenkubun.Text.Substring(0, 1);
                        Cmd.Parameters["健保加入区分"].Value = kenpokanyuu.Text.Substring(0, 1);
                        Cmd.Parameters["世帯主区分"].Value = setainushi.Text.Substring(0, 1);

                        if (syougaikubun.Text == "0　該当しない")
                        {
                            Cmd.Parameters["障害者区分"].Value = 0;
                            Cmd.Parameters["特別障害者区分"].Value = 0; //障害区分を
                        }
                        else if (syougaikubun.Text == "2　特別")
                        {
                            Cmd.Parameters["障害者区分"].Value = 0;
                            Cmd.Parameters["特別障害者区分"].Value = 1;
                        }
                        else
                        {
                            Cmd.Parameters["障害者区分"].Value = 1;
                            Cmd.Parameters["特別障害者区分"].Value = 0;
                        }

                        Cmd.Parameters["異動日"].Value = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");
                        Cmd.Parameters["etc"].Value = "04_異動入力_更新アップデート";

                        string kcode = zokugara.Text.Substring(0, 2);
                        if (kcode == "30" || kcode == "31" || kcode == "32" || kcode == "33" || kcode == "40" || kcode == "41")
                        {
                            Cmd.Parameters["直系尊属区分"].Value = 1;
                        }
                        else
                        {
                            Cmd.Parameters["直系尊属区分"].Value = 0;
                        }

                        //TODO 20251021 追加
                        Cmd.Parameters["特定親族区分"].Value = 0;

                        using (dr = Cmd.ExecuteReader())
                        {
                            int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        }
                    }
                    else
                    {
                        MessageBox.Show("未来日付で更新データがあるみたいですけど。。");
                        return;
                    }

                }
            }
        }

        private void DataDispKazoku()
        {
            if (kyuuyo.SelectedItem == null) return;

            DataTable dt = new DataTable();
            dt = Com.GetDB("select* from dbo.k家族情報取得_異動('" + syainno.Text + "', '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' ) order by 家族識別ＩＤ");

            kazokudgv.DataSource = dt;

            //decimal mae = 0;
            decimal ato = 0;
            foreach (DataRow row in dt.Rows)
            {
                //mae += Convert.ToDecimal(row["手当額(前)"]);
                ato += Convert.ToDecimal(row["手当額(変更後)"]);
            }

            //string kyuuyomae = kyuuyo_m.Text.Substring(0, 1);
            string kyuuyoato = kyuuyo.SelectedItem.ToString().Substring(0, 1);

            //if (kyuuyomae == "C")
            //{
            //    huyoumae.Text = mae.ToString();
            //}
            //else
            //{
            //    huyoumae.Text = "0";
            //}

            //TODO changeイベント発動のため
            huyouato.Text = "1";

            if (kyuuyoato == "C")
            {

                huyouato.Text = ato.ToString();
            }
            else
            {
                huyouato.Text = "0";
            }

            //kazokudgv.CurrentCell = null;



        }

        private void KazokuClear()
        {
            kazomei.Text = "";
            kazosei.Text = sei_m.Text;
            kazokanamei.Text = "";
            kazokanasei.Text = seihuri_m.Text;

            warekicb.SelectedIndex = -1;
            kazoseinengappinew.Value = null;

            zokugara.SelectedIndex = -1;

            doukyokubun.SelectedIndex = 0;
            huyoukubun.SelectedIndex = 0;
            kenpokanyuu.SelectedIndex = 0;
            huyoukubun.SelectedIndex = 0;
            kenpokanyuu.SelectedIndex = 0;
            setainushi.SelectedIndex = 0;
            syougaikubun.SelectedIndex = 0;
            gensenkubun.SelectedIndex = 0;

            kazokubtn.Text = "家族登録";
            kazokubtn.BackColor = Color.Transparent;
            kazokubtn.ForeColor = Color.Black;

            kazokuid.Text = "";

            //削除ボタン
            delkazoku.Visible = false;


        }


        private void GetKazokukoumokuData(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from k家族情報取得_異動個別('" + syainno.Text + "', '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "', '" + str + "')");

            kazosei.Text = dt.Rows[0][1].ToString();
            kazomei.Text = dt.Rows[0][2].ToString();
            kazokanasei.Text = dt.Rows[0][3].ToString();
            kazokanamei.Text = dt.Rows[0][4].ToString();
            if (dt.Rows[0][5].ToString() == "")
            {
                kazoseinengappinew.Value = null;
            }
            else
            {
                kazoseinengappinew.Value = Convert.ToDateTime(dt.Rows[0][5].ToString());
            }

            //kazoseinengappinew.Value = dt.Rows[0][5].ToString();
            zokugara.SelectedIndex = zokugara.FindString(dt.Rows[0][6].ToString());
            doukyokubun.SelectedIndex = doukyokubun.FindString(dt.Rows[0][7].ToString());
            //配偶者[0][8]
            huyoukubun.SelectedIndex = huyoukubun.FindString(dt.Rows[0][8].ToString());
            gensenkubun.SelectedIndex = gensenkubun.FindString(dt.Rows[0][9].ToString());
            kenpokanyuu.SelectedIndex = kenpokanyuu.FindString(dt.Rows[0][10].ToString());
            //setainushi.SelectedIndex = setainushi.FindString(dt.Rows[0][11].ToString());

            if (dt.Rows[0][12].ToString() == "1")
            {
                //特別障碍者の場合
                syougaikubun.SelectedIndex = syougaikubun.FindString("2");

            }
            else
            {
                syougaikubun.SelectedIndex = syougaikubun.FindString(dt.Rows[0][11].ToString());
            }

            kazokuid.Text = str;

            status.Text = dt.Rows[0][13].ToString();
        }



        private void kazokudgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //ヘッダは対象外
            if (kazokudgv.CurrentCell != null)
            {
                KazokuClear();

                DataGridViewRow dgr = kazokudgv.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                //0は家族識別ID
                GetKazokukoumokuData(drv[0].ToString());
            }

            kazokubtn.Text = "家族更新";
            kazokubtn.BackColor = Color.Blue;
            kazokubtn.ForeColor = Color.White;

            if (status.Text == "新規追加")
            {
                delkazoku.Visible = true;
            }
        }

        private void yuubinzyuusyo_Click(object sender, EventArgs e)
        {
            //Form2に送るテキスト
            //何故異動日を送るのかは不明
            string sendText = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectAdress.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            yuubin.Text = receiveText[0];
            zyuusyo.Text = receiveText[1];
        }

        private void kazokunew_Click(object sender, EventArgs e)
        {
            KazokuClear();
        }

        private void kazokanasei_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角→半角　ひらがな→カタカナへ
            string result = Microsoft.VisualBasic.Strings.StrConv(Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Katakana), VbStrConv.Narrow).Trim();

            //カナと空白以外はエラー
            Regex regex = new Regex(@"^[ｦ-ﾝﾞﾟ ]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "カタカナ以外は入力しないでください");
                e.Cancel = true;
            }

            ((TextBox)sender).Text = result;
        }

        private void kazokanamei_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角→半角　ひらがな→カタカナへ
            string result = Microsoft.VisualBasic.Strings.StrConv(Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Katakana), VbStrConv.Narrow).Trim();

            //カナと空白以外はエラー
            Regex regex = new Regex(@"^[ｦ-ﾝﾞﾟ ]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "カタカナ以外は入力しないでください");
                e.Cancel = true;
            }

            ((TextBox)sender).Text = result;
        }

        private void warekicb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (warekicb.SelectedItem == null) return;

            int mm = 1;
            int dd = 1;

            if (kazoseinengappinew.Text != "")
            {
                mm = Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("MM"));
                dd = Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("dd"));
            }


            if (warekicb.SelectedItem.ToString().Substring(0, 2) == "昭和")
            {
                //昭和1桁対応
                if (warekicb.SelectedItem.ToString().Substring(3, 1) == "年")
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 1)) + 1925;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }
                else
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 2)) + 1925;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }

                //int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 2)) + 1925;
                //kazoseinengappinew.Value = new DateTime(i, mm, dd);
                return;
            }

            if (warekicb.SelectedItem.ToString() == "平成元年")
            {
                kazoseinengappinew.Value = new DateTime(1989, mm, dd);
                return;
            }


            if (warekicb.SelectedItem.ToString().Substring(0, 2) == "平成")
            {
                //平成1桁対応
                if (warekicb.SelectedItem.ToString().Substring(3, 1) == "年")
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 1)) + 1988;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }
                else
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }
            }

            if (warekicb.SelectedItem.ToString() == "令和元年")
            {
                kazoseinengappinew.Value = new DateTime(2019, mm, dd);
                return;
            }


            if (warekicb.SelectedItem.ToString().Substring(0, 2) == "令和")
            {
                //平成1桁対応
                if (warekicb.SelectedItem.ToString().Substring(3, 1) == "年")
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 1)) + 2018;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }
                else
                {
                    int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 2)) + 2018;
                    kazoseinengappinew.Value = new DateTime(i, mm, dd);
                }
            }
        }

        private void kazoseinengappinew_ValueChanged(object sender, EventArgs e)
        {
            //エラー対策
            if (kazoseinengappinew.Text == "") return;

            //和暦追加
            if (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) < 1989)
            {
                warekicb.SelectedIndex = warekicb.FindString("昭和" + (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) - 1925).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) == 1989)
            {
                warekicb.SelectedIndex = warekicb.FindString("平成元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) > 1989)
            {
                warekicb.SelectedIndex = warekicb.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) == 2019)
            {
                warekicb.SelectedIndex = warekicb.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) > 2019)
            {
                warekicb.SelectedIndex = warekicb.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(kazoseinengappinew.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        private void Genpyou_Load(object sender, EventArgs e)
        {
            //画面表示の時に選択無にする
            dispdgv.CurrentCell = null;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            shikakudgv.CurrentCell = null;
            kazokudgv.CurrentCell = null;

            ShikakuClear();
            KazokuClear();
        }



        private void delkazoku_Click(object sender, EventArgs e)
        {
            //TODO 新規追加の場合のみ表示
            DataGridViewRow dgr = kazokudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            DialogResult result = MessageBox.Show("ホントに削除していいっすか？" + nl + drv[2].ToString() + "　" + drv[5].ToString(),
                            "警告",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation,
                            MessageBoxDefaultButton.Button2);

            //何が選択されたか調べる
            if (result == DialogResult.Yes)
            {
                //「はい」が選択された時
                DataTable dtdel = new DataTable();
                dtdel = Com.GetDB("delete from QUATRO.dbo.SJMTKAZOKU where 会社コード = 'E0' and 社員番号 = '" + syainno.Text + "' and 家族識別ＩＤ = '" + drv[0].ToString() + "'");

                //一覧データ取得
                DataDispKazoku();

                //選択無
                kazokudgv.CurrentCell = null;

                //項目クリア
                KazokuClear();

                MessageBox.Show("消し去りましたー");
            }
            else if (result == DialogResult.No)
            {
                //「いいえ」が選択された時
            }
        }

        private void shikakunew_Click(object sender, EventArgs e)
        {
            ShikakuClear();
        }

        private void menkyogaku_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\入社入力及び異動入力操作方法.xlsx"); return;
        }

        //職務技能給　変更前
        private void Syokumu_Sum_m_Calc()
        {
            int yaku = Convert.ToInt32(yakusyoku_m.Text.Substring(0, 4));

            //職種
            if (syokusyu_m.Text == "" || syokusyu_m.Text == "-" || yaku < 110)
            {
                syokumukyuu_m.Text = "0";
            }
            else
            {
                DataRow[] dr = Syoku.Select("備考 = '" + syokusyu_m.Text + "'");
                syokumukyuu_m.Text = dr[0][1].ToString();
            }

            //学歴
            if (gakureki_m.Text == "" || gakureki_m.Text == "【基準外】" || gakureki_m.Text == "-")
            {
                gakurekikyuu_m.Text = "0";
            }
            else
            {
                DataRow[] dr = Gaku.Select("備考 = '" + gakureki_m.Text + "'");
                gakurekikyuu_m.Text = dr[0][1].ToString();
            }

            //社外経験
            if (keiken_m.Text == "" || keiken_m.Text == "【基準外】" || keiken_m.Text == "-")
            {
                keikenkyuu_m.Text = "0";
            }
            else
            {
                DataRow[] dr = Keiken.Select("備考 = '" + keiken_m.Text + "'");
                keikenkyuu_m.Text = dr[0][1].ToString();
            }

            //入社年齢
            if (nyuusyaold_m.Text == "" || nyuusyaold_m.Text == "-" || keiken_m.Text == "【基準外】")
            {
                nennreikyuu_m.Text = "0";
            }
            else
            {
                //if (kyuuzitsu_m.Text == "01　月給者" || kyuuzitsu_m.Text == "02　日給者")
                if (kyuuyo_m.Text == "C1　月給者" || kyuuyo_m.Text == "D1　日給者")
                {
                    DataRow[] dr = Nen.Select("年齢 = '" + nyuusyaold_m.Text + "'");
                    nennreikyuu_m.Text = dr[0][1].ToString();
                }
                else
                {
                    nennreikyuu_m.Text = "0";
                }
            }

            //合計額
            //色変更の都合上、変更前から
            syokumu_sum_m.Text = (Convert.ToDecimal(syokumukyuu_m.Text) + Convert.ToDecimal(nennreikyuu_m.Text) + Convert.ToDecimal(keikenkyuu_m.Text) + Convert.ToDecimal(gakurekikyuu_m.Text) + Convert.ToDecimal(kizyungai_m.Text) + Convert.ToDecimal(hyoukakyuu_m.Text)).ToString();

        }




        //職種
        private void syokusyu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (yakusyoku.SelectedIndex == -1) return;

            //表示・非表示
            if (comboBoxSyoku.SelectedItem?.ToString() == syokusyu_m.Text)
            {
                lbl_syokusyu.BackColor = Color.White;
            }
            else
            {
                lbl_syokusyu.BackColor = Color.LightGreen;
            }

            int yaku = Convert.ToInt32(yakusyoku.SelectedItem.ToString().Substring(0, 4));

            //職種
            if (comboBoxSyoku.SelectedItem == null || comboBoxSyoku.SelectedItem.ToString() == "" ||  yaku < 110)
            {
                syokumukyuu.Text = "0";
            }
            else
            {
                DataRow[] dr = Syoku.Select("備考 = '" + comboBoxSyoku.SelectedItem.ToString() + "'");
                syokumukyuu.Text = dr[0][1].ToString();
            }

            GetSyokumu();
            DataDispShikaku();
            RitouCalc();
        }

        //最終学歴
        private void comboBoxGaku_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (comboBoxGaku.SelectedItem?.ToString() == gakureki_m.Text)
            {
                lbl_gakureki.BackColor = Color.White;
            }
            else
            {
                lbl_gakureki.BackColor = Color.LightGreen;
            }

            //学歴
            if (comboBoxGaku.SelectedItem == null || comboBoxGaku.SelectedItem.ToString() == "" || comboBoxGaku.SelectedItem.ToString() == "【基準外】")
            {
                gakurekikyuu.Text = "0";
            }
            else
            {
                DataRow[] dr = Gaku.Select("備考 = '" + comboBoxGaku.SelectedItem.ToString() + "'");
                gakurekikyuu.Text = dr[0][1].ToString();
            }

            GetSyokumu();
        }

        //社外経験
        private void comboBoxKeiken_SelectedIndexChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (comboBoxKeiken.SelectedItem?.ToString() == keiken_m.Text)
            {
                lbl_keiken.BackColor = Color.White;
            }
            else
            {
                lbl_keiken.BackColor = Color.LightGreen;
            }

            //社外経験
            if (comboBoxKeiken.SelectedItem == null || comboBoxKeiken.SelectedItem.ToString() == "" || comboBoxKeiken.SelectedItem.ToString() == "【基準外】")
            {
                keikenkyuu.Text = "0";
            }
            else
            {
                DataRow[] dr = Keiken.Select("備考 = '" + comboBoxKeiken.SelectedItem.ToString() + "'");
                keikenkyuu.Text = dr[0][1].ToString();
            }

            GetSyokumu();
        }

        private void syokumu_sum_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syokumu_sum.Text == syokumu_sum_m.Text)
            {
                lbl_goukei.BackColor = Color.White;
            }
            else
            {
                lbl_goukei.BackColor = Color.LightGreen;
            }


            string kyuu = kyuuyo.SelectedItem.ToString().Substring(0, 1);
            string kyuu_m = "";
            if (kyuuyo_m.Text.Length > 0)
            {
                kyuu_m = kyuuyo_m.Text.ToString().Substring(0, 1);
            }

            string keiy = keiyaku.SelectedItem?.ToString();


            //TODO 
            if (yakusyoku.SelectedItem == null) return;

            int yaku = Convert.ToInt32(yakusyoku.SelectedItem.ToString().Substring(0, 4));

            //月給者で係長未満
            //プラス一般契約社員ではない

            if (kyuu == "C" && yaku >= 110 && keiy == "")
            {
                syokumu.Value = Convert.ToDecimal(syokumu_sum.Text);
            }



        }

        //①
        //通勤非
        private void tuukinhi_m_TextChanged(object sender, EventArgs e)
        {
            tuukinhi2_m.Text = tuukinhi_m.Text;
        }

        //通勤課
        private void tuukinka_m_TextChanged(object sender, EventArgs e)
        {
            tuukinka2_m.Text = tuukinka_m.Text;
        }

        //免許手当
        private void menkyo_m_TextChanged(object sender, EventArgs e)
        {
            menkyomae.Text = menkyo_m.Text;
        }

        //登録手当
        private void touroku_m_TextChanged(object sender, EventArgs e)
        {
            tourokumae.Text = touroku_m.Text;
        }

        //扶養手当
        private void huyou_m_TextChanged(object sender, EventArgs e)
        {
            huyoumae.Text = huyou_m.Text;
        }




        //②
        private void tuukinhi2_m_TextChanged(object sender, EventArgs e)
        {
            tuukinhi2.Text = tuukinhi2_m.Text;
        }

        private void tuukinka2_m_TextChanged(object sender, EventArgs e)
        {
            tuukinka2.Text = tuukinka2_m.Text;
        }

        private void menkyomae_TextChanged(object sender, EventArgs e)
        {
            menkyoato.Text = menkyomae.Text;
        }

        private void tourokumae_TextChanged(object sender, EventArgs e)
        {
            tourokuato.Text = tourokumae.Text;
        }

        private void huyoumae_TextChanged(object sender, EventArgs e)
        {
            huyouato.Text = huyoumae.Text;
        }



        private void menkyoato_TextChanged(object sender, EventArgs e)
        {
            menkyo.Text = menkyoato.Text;
        }

        private void tourokuato_TextChanged(object sender, EventArgs e)
        {
            touroku.Text = tourokuato.Text;
        }

        private void huyouato_TextChanged(object sender, EventArgs e)
        {
            huyou.Text = huyouato.Text;
        }






        private void ymlist_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetIdouData();
            DataReset();
        }

        private void yuubinzyuusyo_Click_1(object sender, EventArgs e)
        {
            //Form2に送るテキスト
            string sendText = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectAdress.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            yuubin.Text = receiveText[0];
            zyuusyo.Text = receiveText[1];
        }

        private void nyuusyaold_TextChanged(object sender, EventArgs e)
        {
            GetNenreiKyuu();
        }

        private void GetNenreiKyuu()
        {
            if (nyuusyaold.Text == "" || nyuusyaold.Text == "-") return;

            //入社年齢
            if (nyuusyaold.Text == "" || nyuusyaold.Text == "-" || comboBoxKeiken.SelectedItem?.ToString() == "【基準外】")
            {
                nennreikyuu.Text = "0";
            }
            else
            {
                if (kyuuyo.Text == "C1　月給者")
                {
                    DataRow[] dr = Nen.Select("年齢 = '" + nyuusyaold.Text + "'");
                    nennreikyuu.Text = dr[0][1].ToString();
                }
                else
                {
                    nennreikyuu.Text = "0";
                }
            }

            GetSyokumu();
        }










        


        private void GetSyokumu()
        {
            //TODO
            if (kyuuyo.SelectedItem == null) return;
            if (yakusyoku.SelectedItem == null) return;

            string kyuu = kyuuyo.SelectedItem.ToString().Substring(0, 1);

            int yaku = Convert.ToInt32(yakusyoku.SelectedItem.ToString().Substring(0, 4));

            syokumu_sum.Text = (Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) + Convert.ToDecimal(gakurekikyuu.Text) + Convert.ToDecimal(kizyungai.Text) + Convert.ToDecimal(hyoukakyuu.Text)).ToString();

            //TODO
            //単年契約
            //if (keiyaku.SelectedItem?.ToString() == "10　一般契約社員" || keiyaku.SelectedItem?.ToString() == "20　単年契約社員") return;
            if (keiyaku.SelectedItem?.ToString() != "") return;

            //TODO 
            //月給者で係長未満
            if (kyuu == "C" && yaku >= 110)
            {
                syokumu.Value = Convert.ToDecimal(syokumu_sum.Text);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                shikakukigenday.Visible = true;
            }
            else
            {
                shikakukigenday.Visible = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (syainno.Text == "") return;

            string msg = "No：" + lbl_no.Text + nl;
            msg += "異動日：" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + nl;
            msg += "氏名：" + name.Text + nl;
            msg += "組織名：" + soshiki_m.Text + nl;
            msg += "現場名：" + genba_m.Text + nl;

            DialogResult result = MessageBox.Show("本当に消去してもいいですか？" + nl + msg,
                                    "警告",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Warning,
                                    MessageBoxDefaultButton.Button2);
            
            //何が選択されたか調べる
            if (result == DialogResult.Yes)
            {
                //「はい」が選択された時
                DataTable dtdel = new DataTable();
                //string sql = "delete from dbo.異動データNo where 社員番号 = '" + syainno.Text + "' and 異動年月日 = '" + idouday.Value.ToString("yyyy/MM/dd") + "' ";
                string sql = "delete from dbo.i異動データ where No = '" + lbl_no.Text + "'";
                dtdel = Com.GetDB(sql);

                //①通勤元データの適用開始日が異動日のレコードがないか確認
                DataTable dtcheck = new DataTable();
                string sql2 = "select * from dbo.t通勤手当元データ where 社員番号 = '" + syainno.Text + "' and 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'";
                dtcheck = Com.GetDB(sql2);
                if (dtcheck.Rows.Count > 0)
                {
                    //①のデータがあった場合、①のレコードは削除。
                    DataTable dtdel2 = new DataTable();
                    dtdel2 = Com.GetDB("delete from dbo.t通勤手当元データ where 社員番号 = '" + syainno.Text + "' and 適用開始日 = '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "'");

                    //①のデータがあった場合で、適用終了日が異動日の前日になっているデータを、9999/12/31に更新する。
                    DataTable dtupdate = new DataTable();
                    dtupdate = Com.GetDB("update dbo.t通勤手当元データ set 適用終了日 = '9999/12/31' where 社員番号 = '" + syainno.Text + "' and 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "'");
                }

                //TODO 家族と資格も本来戻し処理が必要だが、パターンが多いため、とりあえずi異動削除履歴テーブルに情報を残す。
                DeleteInfo();

                //異動一覧取得
                GetIdouData();

                //入力フォームクリア
                DataReset();
                ShikakuClear();
                KazokuClear();

                shikakudgv.DataSource = "";
                kazokudgv.DataSource = "";


                MessageBox.Show("消し去りましたー");
            }
            else if (result == DialogResult.No)
            {
                //「いいえ」が選択された時
            }
        }

        private double Getrday(string syuur)
        {
            //休暇付与区分

            double rday = 0;

            if (string.IsNullOrEmpty(kyuuzitsu_m.Text) || kyuuzitsu_m.Text == "-") return rday;


            //休日区分によって変更
            if (kyuuzitsu_m.Text.Split('　')[1] == "10")
            {
                switch (syuur)
                {
                    case "0　5日以上": rday = 21.5; break;
                    case "1　4日": rday = 17.5; break;
                    case "2　3日": rday = 13.5; break;
                    case "3　2日": rday = 9; break;
                    case "4　1日": rday = 4.5; break;
                    case "9　付与なし": rday = 21.5; break;
                    default: break;
                }
            }
            else
            {
                switch (syuur)
                {
                    case "0　5日以上": rday = 20.5; break;
                    case "1　4日": rday = 16.5; break;
                    case "2　3日": rday = 12.5; break;
                    case "3　2日": rday = 8; break;
                    case "4　1日": rday = 3.5; break;
                    case "9　付与なし": rday = 20.5; break;
                    default: break;
                }
            }

            return rday;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //TODO 新規追加の場合のみ表示
            DataGridViewRow dgr = shikakudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            if (drv[0].ToString().Substring(0,1) != "T")
            {
                MessageBox.Show("資格は削除できません。できるのは登録だけす。");
                return;
            }

            DialogResult result = MessageBox.Show("ホントに削除していいっすか？" + nl + drv[1].ToString() + "　" + drv[3].ToString(),
                            "警告",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation,
                            MessageBoxDefaultButton.Button2);

            //何が選択されたか調べる
            if (result == DialogResult.Yes)
            {
                //「はい」が選択された時
                DataTable dtdel = new DataTable();
                dtdel = Com.GetDB("update QUATRO.dbo.SJMTSHIKAK set 適用終了日 = '" + Convert.ToDateTime(idoudaynew.Value).AddDays(-1).ToString("yyyy/MM/dd") + "' where 会社コード = 'E0' and 社員番号 = '" + syainno.Text + "' and 資格コード = '" + drv[0].ToString() + "' and 適用終了日 = '9999/12/31'");

                //一覧データ取得
                DataDispShikaku();

                //選択無
                shikakudgv.CurrentCell = null;

                //項目クリア
                ShikakuClear();

                MessageBox.Show("消し去りましたー");
            }
            else if (result == DialogResult.No)
            {
                //「いいえ」が選択された時
            }
        }



        //private void menkyo_ValueChanged(object sender, EventArgs e)
        //{
        //    //表示・非表示
        //    if (menkyo.Text == menkyo_m.Text)
        //    {
        //        lbl_menkyo.BackColor = Color.White;
        //    }
        //    else
        //    {
        //        lbl_menkyo.BackColor = Color.LightGreen;
        //    }

        //    Disp();
        //}

        private void honkyuu_ValueChanged_1(object sender, EventArgs e)
        {
            //表示・非表示
            if (honkyuu.Value == Convert.ToDecimal(honkyuu_m.Text))
            {
                lbl_honkyuu.BackColor = Color.White;
            }
            else
            {
                lbl_honkyuu.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void syokumu_ValueChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syokumu.Value == Convert.ToDecimal(syokumu_m.Text))
            {
                lbl_syokumu.BackColor = Color.White;
            }
            else
            {
                lbl_syokumu.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void menkyo_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (menkyo.Text == menkyo_m.Text)
            {
                lbl_menkyo.BackColor = Color.White;
            }
            else
            {
                lbl_menkyo.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void yakuteate_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (yakuteate.Text == yakuteate_m.Text)
            {
                lbl_yakuteate.BackColor = Color.White;
            }
            else
            {
                lbl_yakuteate.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void ritou_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (ritou.Text == ritou_m.Text)
            {
                lbl_ritou.BackColor = Color.White;
            }
            else
            {
                lbl_ritou.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void tuushin_m_TextChanged(object sender, EventArgs e)
        {
            tuushin.Text = tuushin_m.Text;
        }

        private void tuushin_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (tuushin.Text == tuushin_m.Text)
            {
                lbl_tuushin.BackColor = Color.White;
            }
            else
            {
                lbl_tuushin.BackColor = Color.LightGreen;
            }

            Disp();
        }


        //③
        private void tuukinhi2_TextChanged(object sender, EventArgs e)
        {
            //色変更
            if (tuukinhi2.Text == tuukinhi2_m.Text)
            {
                lbl_tuukinhi2.BackColor = Color.White;
            }
            else
            {
                lbl_tuukinhi2.BackColor = Color.LightGreen;
            }


            tuukinhi.Text = tuukinhi2.Text.Replace(".0", "");
        }

        private void tuukinka2_TextChanged(object sender, EventArgs e)
        {
            //色変更
            if (tuukinka2.Text == tuukinka2_m.Text)
            {
                lbl_tuukinka2.BackColor = Color.White;
            }
            else
            {
                lbl_tuukinka2.BackColor = Color.LightGreen;
            }


            tuukinka.Text = tuukinka2.Text.Replace(".0", "");
        }


        private void tuutanka_TextChanged(object sender, EventArgs e)
        {
            //色変更
            if (tuutanka.Text == tuutanka_m.Text)
            {
                lbl_tuutanka.BackColor = Color.White;
            }
            else
            {
                lbl_tuutanka.BackColor = Color.LightGreen;
            }
        }

        private void idouday_ValueChanged(object sender, EventArgs e)
        {
            CheckNoDay(syainno.Text, lbl_no.Text);
        }

        private void seihuri_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角→半角　ひらがな→カタカナへ
            string result = Microsoft.VisualBasic.Strings.StrConv(Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Katakana), VbStrConv.Narrow).Trim();

            //カナと空白以外はエラー
            Regex regex = new Regex(@"^[ｦ-ﾝﾞﾟ ]+$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "カタカナ以外は入力しないでください");
                e.Cancel = true;
            }

            ((TextBox)sender).Text = result;
        }

        private void zyuusyo2_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (zyuusyo2.Text == zyuusyo_m2.Text)
            {
                //zyuusyo_m.ForeColor = zyuusyo_m.BackColor;
                lbl_zyuusyo2.BackColor = Color.White;
            }
            else
            {
                //zyuusyo_m.ForeColor = Color.Black;
                lbl_zyuusyo2.BackColor = Color.LightGreen;
            }
        }

        private void yuubin2_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (yuubin2.Text == yuubin_m2.Text)
            {
                //yuubin_m.ForeColor = yuubin_m.BackColor;
                lbl_yuubin2.BackColor = Color.White;
            }
            else
            {
                //yuubin_m.ForeColor = Color.Black;
                lbl_yuubin2.BackColor = Color.LightGreen;
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {

            //Form2に送るテキスト
            string sendText = Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectAdress.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            yuubin2.Text = receiveText[0];
            zyuusyo2.Text = receiveText[1];
        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        private void shikakuselect_Click(object sender, EventArgs e)
        {
            //Form2に送るテキスト
            string sendText = ""; //特になし

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectShikaku.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            ShikakuClear();

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            shikakutextb.Text = receiveText[0];


            if (receiveText[1] == "必須")
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (syainno.Text == "") return;
            //TODO 従業員検索に同じ処理があります。。
            //TODO 異動にも同じ処理があります。。

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button9.Enabled = false;

            UpdateRoudou();

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button9.Enabled = true;

            MessageBox.Show("更新しました。");
        }

        private void UpdateRoudou()
        {
            //TODO 従業員検索に同じ処理があります。。

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[r労働条件更新]";
                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("契約年月", SqlDbType.VarChar)); Cmd.Parameters["契約年月"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用区分", SqlDbType.Char)); Cmd.Parameters["雇用区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用開始日", SqlDbType.VarChar)); Cmd.Parameters["雇用開始日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("雇用終了日", SqlDbType.VarChar)); Cmd.Parameters["雇用終了日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("更新区分", SqlDbType.Char)); Cmd.Parameters["更新区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("就業場所", SqlDbType.VarChar)); Cmd.Parameters["就業場所"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("業務内容", SqlDbType.VarChar)); Cmd.Parameters["業務内容"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定始業", SqlDbType.Time)); Cmd.Parameters["定始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定終業", SqlDbType.Time)); Cmd.Parameters["定終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定休憩", SqlDbType.VarChar)); Cmd.Parameters["定休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1始業", SqlDbType.Time)); Cmd.Parameters["S1始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1終業", SqlDbType.Time)); Cmd.Parameters["S1終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S1休憩", SqlDbType.VarChar)); Cmd.Parameters["S1休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2始業", SqlDbType.Time)); Cmd.Parameters["S2始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2終業", SqlDbType.Time)); Cmd.Parameters["S2終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S2休憩", SqlDbType.VarChar)); Cmd.Parameters["S2休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3始業", SqlDbType.Time)); Cmd.Parameters["S3始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3終業", SqlDbType.Time)); Cmd.Parameters["S3終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S3休憩", SqlDbType.VarChar)); Cmd.Parameters["S3休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4始業", SqlDbType.Time)); Cmd.Parameters["S4始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4終業", SqlDbType.Time)); Cmd.Parameters["S4終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S4休憩", SqlDbType.VarChar)); Cmd.Parameters["S4休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5始業", SqlDbType.Time)); Cmd.Parameters["S5始業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5終業", SqlDbType.Time)); Cmd.Parameters["S5終業"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("S5休憩", SqlDbType.VarChar)); Cmd.Parameters["S5休憩"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("時間外労働区分", SqlDbType.Char)); Cmd.Parameters["時間外労働区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("夜間勤務区分", SqlDbType.Char)); Cmd.Parameters["夜間勤務区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("休日回数", SqlDbType.VarChar)); Cmd.Parameters["休日回数"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("休出有無", SqlDbType.Char)); Cmd.Parameters["休出有無"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("定年区分", SqlDbType.Char)); Cmd.Parameters["定年区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("賞与区分", SqlDbType.Char)); Cmd.Parameters["賞与区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("退職金区分", SqlDbType.Char)); Cmd.Parameters["退職金区分"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainno.Text;
                    Cmd.Parameters["契約年月"].Value = keiyakunengetsu.Value;
                    Cmd.Parameters["雇用区分"].Value = koyoukubun.Text;
                    Cmd.Parameters["雇用開始日"].Value = koyoukaishibi.Value;
                    Cmd.Parameters["雇用終了日"].Value = koyousyuuryoubi.Value;
                    Cmd.Parameters["更新区分"].Value = koushinkubun.Text;
                    Cmd.Parameters["就業場所"].Value = syuugyoubasyo.Text;
                    Cmd.Parameters["業務内容"].Value = gyoumunaiyou.Text;
                    Cmd.Parameters["定始業"].Value = dtps0.Text;
                    Cmd.Parameters["定終業"].Value = dtpe0.Text;
                    Cmd.Parameters["定休憩"].Value = kyuukei0.Text;
                    Cmd.Parameters["S1始業"].Value = dtps1.Text;
                    Cmd.Parameters["S1終業"].Value = dtpe1.Text;
                    Cmd.Parameters["S1休憩"].Value = kyuukei1.Text;
                    Cmd.Parameters["S2始業"].Value = dtps2.Text;
                    Cmd.Parameters["S2終業"].Value = dtpe2.Text;
                    Cmd.Parameters["S2休憩"].Value = kyuukei2.Text;
                    Cmd.Parameters["S3始業"].Value = dtps3.Text;
                    Cmd.Parameters["S3終業"].Value = dtpe3.Text;
                    Cmd.Parameters["S3休憩"].Value = kyuukei3.Text;
                    Cmd.Parameters["S4始業"].Value = dtps4.Text;
                    Cmd.Parameters["S4終業"].Value = dtpe4.Text;
                    Cmd.Parameters["S4休憩"].Value = kyuukei4.Text;
                    Cmd.Parameters["S5始業"].Value = dtps5.Text;
                    Cmd.Parameters["S5終業"].Value = dtpe5.Text;
                    Cmd.Parameters["S5休憩"].Value = kyuukei5.Text;
                    Cmd.Parameters["時間外労働区分"].Value = zikangairoudou.Text;
                    Cmd.Parameters["夜間勤務区分"].Value = yakankinmu.Text;
                    Cmd.Parameters["休日回数"].Value = kyuujitsukaisuu.Text;
                    Cmd.Parameters["休出有無"].Value = kyuusyustu.Text;
                    Cmd.Parameters["定年区分"].Value = teinen.Text;
                    Cmd.Parameters["賞与区分"].Value = syouyo.Text;
                    Cmd.Parameters["退職金区分"].Value = taisyokukin.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        //MessageBox.Show("更新しました。");
                    }
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (syainno.Text == "") return;
            if (BeforeCheck()) return;

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button10.Enabled = false;

            //更新して出力
            UpdateRoudou();

            //対象データ取得
            DataTable dt = new DataTable();

            //TODO 異動データleft join
            dt = Com.GetDB("select * from r労働条件取得 where 社員番号 = '" + syainno.Text + "'");

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            if (kyuuyo.Text.Substring(0, 2) == "E1" || kyuuyo.Text.Substring(0, 2) == "F1")
            {
                //パート・アルバイト
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\11_労働条件ファミリー.xlsx");
            }
            else
            {
                //TODO 日給者・役員・契約社員の対応が未
                //月給者
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\11_労働条件一般.xlsx");
            }

            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ls[j, i + 1].Value = dt.Rows[i][j].ToString();
                }
            }

            //苗字変更があった場合
            if (sei.Text != name.Text.Split('　')[0])
            { 
                ls[0, 2].Value = sei.Text + name.Text.Split('　')[0]; //名前
            }

            //ls[1, 2].Value = kyuuka.SelectedItem.ToString(); //週労働数
            //ls[2, 2].Value = syainno.Text; //社員番号
            //ls[3, 2].Value = "-"; //契約年月
            //ls[4, 2].Value = "-"; //雇用区分
            //ls[5, 2].Value = "-"; //雇用開始日
            //ls[6, 2].Value = "-"; //雇用終了日
            //ls[7, 2].Value = "-"; //更新区分

            //元現場と変更現場が異なり、かつ就業場所にデータ無の場合
            if (genba.SelectedItem?.ToString() != genba_m.Text && dt.Rows[0][8].ToString() == "")
            { 
                ls[8, 2].Value = genba.SelectedItem?.ToString();  //就業場所 ※現場名
            }



            //ls[9, 2].Value = "-"; //業務内容
            //ls[10, 2].Value = "-"; //定時
            //ls[11, 2].Value = "-"; //シフト1
            //ls[12, 2].Value = "-"; //シフト2
            //ls[13, 2].Value = "-"; //シフト3
            //ls[14, 2].Value = "-"; //シフト4
            //ls[15, 2].Value = "-"; //シフト5
            //ls[16, 2].Value = "-"; //時間外労働区分
            //ls[17, 2].Value = "-"; //夜間勤務区分
            //ls[18, 2].Value = "-"; //休日回数
            //ls[19, 2].Value = "-"; //休出有無
            //ls[20, 2].Value = "-"; //定年区分
            //ls[21, 2].Value = "-"; //賞与区分
            //ls[22, 2].Value = "-"; //退職金区分
            //ls[23, 2].Value = "-"; //etc1
            //ls[24, 2].Value = "-"; //etc2
            //ls[25, 2].Value = "-"; //etc3

            ls[26, 1].Value = Convert.ToInt32(honkyuu.Value); //本給
            ls[27, 1].Value = Convert.ToInt32(syokumu.Value); //職務技能給

            ls[28, 1].Value = Convert.ToInt32(syokumukyuu.Text); //職務給
            ls[29, 1].Value = Convert.ToInt32(gakurekikyuu.Text); //学歴給
            ls[30, 1].Value = Convert.ToInt32(kizyungai.Text); //基準外(年齢経験学歴)
            ls[31, 1].Value = Convert.ToInt32(nennreikyuu.Text); //年齢給
            ls[32, 1].Value = Convert.ToInt32(keikenkyuu.Text); //経験給
            ls[33, 1].Value = Convert.ToInt32(hyoukakyuu.Text); //評価額

            ls[34, 1].Value = Convert.ToInt32(yakuteate.Text); //役職手当
            ls[35, 1].Value = Convert.ToInt32(huyou.Text); //扶養手当
            ls[36, 1].Value = Convert.ToInt32(menkyo.Text); //免許手当
            ls[37, 1].Value = Convert.ToInt32(ritou.Text); //離島手当
            ls[38, 1].Value = Convert.ToInt32(tokubetsu.Value); //特別手当
            ls[39, 1].Value = Convert.ToInt32(syukkou.Value); //転勤手当
            ls[40, 1].Value = Convert.ToInt32(touroku.Text); //登録手当
            ls[41, 1].Value = Convert.ToInt32(tuushin.Text); //通信手当
            ls[42, 1].Value = Convert.ToInt32(syaryou.Text); //車両手当

            string tomo = "";
            if (tomokubun.SelectedItem?.ToString().Split('　')[0] == "")
            {
                tomo = "0";
            }
            else
            {
                tomo = tomokubun.SelectedItem?.ToString().Split('　')[0];
            }

            ls[43, 1].Value = tomo.ToString(); //友の会区分

            //ls[44, 2].Value = ""; //厚年
            //ls[45, 2].Value = ""; //健保
            //ls[46, 2].Value = ""; //雇保

            ls[47, 1].Value = Convert.ToInt32(zikyuu.Value); //時給

            string localPass = @"C:\ODIS\ROUDOU\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");
            System.Diagnostics.Process.Start(exlName + ".xlsx");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            //button10.Enabled = true;
        }

        private bool BeforeCheck()
        {
            //TODO 従業員検索に同じ処理があります。。
            //必須項目が入ってない場合は出力できない処理
            string msg = "";
            if (keiyakunengetsu.Value.Equals(DBNull.Value)) msg += "・作成日" + Com.nl;
            if (koyoukubun.SelectedItem?.ToString() == "") msg += "・雇用区分" + Com.nl;
            //if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり" && koyoukaishibi.Value.Equals(DBNull.Value)) msg += "・雇用開始日" + Com.nl;
            if (koyoukaishibi.Value.Equals(DBNull.Value)) msg += "・雇用開始日" + Com.nl;
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり" && koyousyuuryoubi.Value.Equals(DBNull.Value)) msg += "・雇用終了日" + Com.nl;
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり" && koushinkubun.SelectedItem?.ToString() == "") msg += "・更新区分" + Com.nl;
            if (gyoumunaiyou.Text == "") msg += "・業務内容" + Com.nl;
            if ((dtpe0.Value - dtps0.Value).Hours == 0 && (dtpe1.Value - dtps1.Value).Hours == 0) msg += "・勤務時間" + Com.nl;
            if (zikangairoudou.SelectedItem?.ToString() == "") msg += "・時間外労働" + Com.nl;
            if (yakankinmu.SelectedItem?.ToString() == "") msg += "・夜間勤務" + Com.nl;
            if (kyuujitsukaisuu.SelectedItem?.ToString() == "") msg += "・休日回数" + Com.nl;
            if (kyuusyustu.SelectedItem?.ToString() == "") msg += "・休出有無" + Com.nl;
            if (teinen.SelectedItem?.ToString() == "") msg += "・定年区分" + Com.nl;
            if (syouyo.SelectedItem?.ToString() == "") msg += "・賞与区分" + Com.nl;
            if (taisyokukin.SelectedItem?.ToString() == "") msg += "・退職金区分" + Com.nl;

            if (!koyoukaishibi.Value.Equals(DBNull.Value) && !koyousyuuryoubi.Value.Equals(DBNull.Value))
            {
                if (Convert.ToDateTime(koyoukaishibi.Value) > Convert.ToDateTime(koyousyuuryoubi.Value)) msg += "契約期間がおかしー" + Com.nl;
            }

            if (msg != "")
            {
                msg = "下記項目は必須です。" + Com.nl + msg;
                MessageBox.Show(msg);
                return true;
            }
            else
            {
                return false;
            }
        }

        private void koyoukubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり")
            {
                //koyoukaishibi.Enabled = true;
                koyousyuuryoubi.Enabled = true;
                koushinkubun.Enabled = true;
            }
            else
            {
                //koyoukaishibi.Value = null;
                //koyoukaishibi.Enabled = false;

                koyousyuuryoubi.Value = null;
                koyousyuuryoubi.Enabled = false;

                koushinkubun.Enabled = false;
                koushinkubun.SelectedIndex = 0;
            }
        }

        private void dtp0_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe0.Value - dtps0.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps0.Value.Year, dtps0.Value.Month, dtps0.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps0.Value;

                TimeSpan ts2 = dtpe0.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei0.Text = "120分";
                    }
                    else
                    {
                        kyuukei0.Text = "60分";
                    }
                }
                else
                {
                    kyuukei0.Text = "45分";
                }
            }
            else
            {
                kyuukei0.Text = "";
            }

            //勤務時間追加
            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh0.Text = h.ToString();
            }
            else
            {
                kinmuh0.Text = "";
            }
        }

        private void dtp1_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe1.Value - dtps1.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps1.Value.Year, dtps1.Value.Month, dtps1.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps1.Value;

                TimeSpan ts2 = dtpe1.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei1.Text = "120分";
                    }
                    else
                    {
                        kyuukei1.Text = "60分";
                    }
                }
                else
                {
                    kyuukei1.Text = "45分";
                }
            }
            else
            {
                kyuukei1.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh1.Text = h.ToString();
            }
            else
            {
                kinmuh1.Text = "";
            }
        }

        private void dtp2_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe2.Value - dtps2.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;


            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps2.Value.Year, dtps2.Value.Month, dtps2.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps2.Value;

                TimeSpan ts2 = dtpe2.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei2.Text = "120分";
                    }
                    else
                    {
                        kyuukei2.Text = "60分";
                    }
                }
                else
                {
                    kyuukei2.Text = "45分";
                }
            }
            else
            {
                kyuukei2.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh2.Text = h.ToString();
            }
            else
            {
                kinmuh2.Text = "";
            }
        }

        private void dtp3_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe3.Value - dtps3.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps3.Value.Year, dtps3.Value.Month, dtps3.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps3.Value;

                TimeSpan ts2 = dtpe3.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei3.Text = "120分";
                    }
                    else
                    {
                        kyuukei3.Text = "60分";
                    }
                }
                else
                {
                    kyuukei3.Text = "45分";
                }
            }
            else
            {
                kyuukei3.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh3.Text = h.ToString();
            }
            else
            {
                kinmuh3.Text = "";
            }
        }

        private void dtp4_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe4.Value - dtps4.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps4.Value.Year, dtps4.Value.Month, dtps4.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps4.Value;

                TimeSpan ts2 = dtpe4.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei4.Text = "120分";
                    }
                    else
                    {
                        kyuukei4.Text = "60分";
                    }
                }
                else
                {
                    kyuukei4.Text = "45分";
                }
            }
            else
            {
                kyuukei4.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh4.Text = h.ToString();
            }
            else
            {
                kinmuh4.Text = "";
            }
        }

        private void dtp5_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dtpe5.Value - dtps5.Value;
            double h = ts.Hours;
            int mm = ts.Minutes;

            //日跨ぎ
            if (ts.Hours < 0)
            {
                DateTime datetime_set = new DateTime(dtps5.Value.Year, dtps5.Value.Month, dtps5.Value.Day, 00, 00, 00); //年, 月, 日, 時間, 分, 秒
                ts = datetime_set - dtps5.Value;

                TimeSpan ts2 = dtpe5.Value - datetime_set.AddDays(-1);
                h = (ts + ts2).Hours;
                mm = (ts + ts2).Minutes;
            }

            if (h >= 6)
            {
                if (h >= 8)
                {
                    if (h >= 14)
                    {
                        kyuukei5.Text = "120分";
                    }
                    else
                    {
                        kyuukei5.Text = "60分";
                    }
                }
                else
                {
                    kyuukei5.Text = "45分";
                }
            }
            else
            {
                kyuukei5.Text = "";
            }

            if (mm >= 30) h = h + 0.5;
            if (h != 0)
            {
                kinmuh5.Text = h.ToString();
            }
            else
            {
                kinmuh5.Text = "";
            }
        }

        private void DataDispRoudou()
        {
            //労働条件情報
            DataTable roudoudt = new DataTable();
            roudoudt = Com.GetDB("select * from dbo.r労働条件 where 社員番号 = '" + syainno.Text + "'");


            if (roudoudt.Rows.Count == 0)
            {
                //TODO 通常データ無はありえない。　途中導入のため。。
                //労働条件テーブルへインサート
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
                            DataTable dt = new DataTable();
                            Cmd.CommandText = "insert into dbo.r労働条件(社員番号) VALUES('" + syainno.Text + "')";
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

                dtps0.Value = zerodt; dtpe0.Value = zerodt;
                dtps1.Value = zerodt; dtpe1.Value = zerodt;
                dtps2.Value = zerodt; dtpe2.Value = zerodt;
                dtps3.Value = zerodt; dtpe3.Value = zerodt;
                dtps4.Value = zerodt; dtpe4.Value = zerodt;
                dtps5.Value = zerodt; dtpe5.Value = zerodt;
            }
            else
            {


                if (roudoudt.Rows[0]["契約年月"].Equals(DBNull.Value) || roudoudt.Rows[0]["契約年月"].Equals(""))
                {
                    keiyakunengetsu.Value = null;
                }
                else
                {
                    keiyakunengetsu.Value = Convert.ToDateTime(roudoudt.Rows[0]["契約年月"].ToString());
                }

                koyoukubun.SelectedIndex = koyoukubun.FindString(roudoudt.Rows[0]["雇用区分"].ToString());

                if (roudoudt.Rows[0]["雇用開始日"].Equals(DBNull.Value) || roudoudt.Rows[0]["雇用開始日"].Equals(""))
                {
                    koyoukaishibi.Value = null;
                }
                else
                {
                    koyoukaishibi.Value = Convert.ToDateTime(roudoudt.Rows[0]["雇用開始日"].ToString());
                }

                if (roudoudt.Rows[0]["雇用終了日"].Equals(DBNull.Value) || roudoudt.Rows[0]["雇用終了日"].Equals(""))
                {
                    koyousyuuryoubi.Value = null;
                }
                else
                {
                    koyousyuuryoubi.Value = Convert.ToDateTime(roudoudt.Rows[0]["雇用終了日"].ToString());
                }

                koushinkubun.SelectedIndex = koushinkubun.FindString(roudoudt.Rows[0]["更新区分"].ToString());

                syuugyoubasyo.Text = roudoudt.Rows[0]["就業場所"].ToString();
                gyoumunaiyou.Text = roudoudt.Rows[0]["業務内容"].ToString();

                dtps0.Value = roudoudt.Rows[0]["定始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["定始業"].ToString());
                dtpe0.Value = roudoudt.Rows[0]["定終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["定終業"].ToString());
                dtps1.Value = roudoudt.Rows[0]["S1始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S1始業"].ToString());
                dtpe1.Value = roudoudt.Rows[0]["S1終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S1終業"].ToString());
                dtps2.Value = roudoudt.Rows[0]["S2始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S2始業"].ToString());
                dtpe2.Value = roudoudt.Rows[0]["S2終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S2終業"].ToString());
                dtps3.Value = roudoudt.Rows[0]["S3始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S3始業"].ToString());
                dtpe3.Value = roudoudt.Rows[0]["S3終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S3終業"].ToString());
                dtps4.Value = roudoudt.Rows[0]["S4始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S4始業"].ToString());
                dtpe4.Value = roudoudt.Rows[0]["S4終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S4終業"].ToString());
                dtps5.Value = roudoudt.Rows[0]["S5始業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S5始業"].ToString());
                dtpe5.Value = roudoudt.Rows[0]["S5終業"].Equals(DBNull.Value) ? zerodt : Convert.ToDateTime(roudoudt.Rows[0]["S5終業"].ToString());

                zikangairoudou.SelectedIndex = zikangairoudou.FindString(roudoudt.Rows[0]["時間外労働区分"].ToString());
                yakankinmu.SelectedIndex = yakankinmu.FindString(roudoudt.Rows[0]["夜間勤務区分"].ToString());

                kinmuH.Text = "【" + kinmu.SelectedItem.ToString() + "時間】";

                //TODO 週労働数との連動処理
                syuuroucopy.Text = kyuuka.SelectedItem.ToString();

                switch (kyuuka.SelectedItem.ToString())
                {
                    case "0　5日以上":

                        if (kyuuyo.Text.Substring(0, 2) == "E1")
                        {
                            //パート
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("2");
                        }
                        else if (kyuuyo.Text.Substring(0, 2) == "F1")
                        {
                            //アルバイト
                            //TODO 設定無で
                        }
                        else if (kyuuyo.Text.Substring(0, 2) == "C1")
                        {
                            //正社員
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("1");
                        }
                        else if (kyuuyo.Text.Substring(0, 2) == "B1")
                        {
                            //兼務役員
                            kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("1");
                        }

                        break;
                    case "1　4日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("3");
                        break;
                    case "2　3日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("4");
                        break;
                    case "3　2日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("5");
                        break;
                    case "4　1日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("6");
                        break;
                    case "9　付与なし":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("6");
                        break;
                default:
                        //役員
                        kyuujitsukaisuu.SelectedIndex = -1;
                        break;

                }

                if (roudoudt.Rows[0]["休日回数"].ToString() != "")
                {
                    kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString(roudoudt.Rows[0]["休日回数"].ToString());
                }

                kyuusyustu.SelectedIndex = kyuusyustu.FindString(roudoudt.Rows[0]["休出有無"].ToString());
                teinen.SelectedIndex = teinen.FindString(roudoudt.Rows[0]["定年区分"].ToString());
                syouyo.SelectedIndex = syouyo.FindString(roudoudt.Rows[0]["賞与区分"].ToString());
                taisyokukin.SelectedIndex = taisyokukin.FindString(roudoudt.Rows[0]["退職金区分"].ToString());
            }


            //選択無だった場合に初期データとして

            //期末日設定
            //DateTime s49ki = new System.DateTime(2020, 04, 01, 0, 0, 0, 0);
            //DateTime s50ki = new System.DateTime(2021, 04, 01, 0, 0, 0, 0);
            //DateTime s51ki = new System.DateTime(2022, 04, 01, 0, 0, 0, 0);
            //DateTime s52ki = new System.DateTime(2023, 04, 01, 0, 0, 0, 0);
            //DateTime s53ki = new System.DateTime(2024, 04, 01, 0, 0, 0, 0);
            //DateTime s54ki = new System.DateTime(2025, 04, 01, 0, 0, 0, 0);
            //DateTime s55ki = new System.DateTime(2026, 04, 01, 0, 0, 0, 0);

            //DateTime kimatsuday = new System.DateTime(2021, 03, 31, 0, 0, 0, 0);
            //if (saiyoudate.Value < s50ki)
            //{
            //    //DateTime kimatsuday = new System.DateTime(2020, 03, 31, 0, 0, 0, 0);
            //}
            //else if (saiyoudate.Value < s51ki)
            //{
            //    kimatsuday = new System.DateTime(2022, 03, 31, 0, 0, 0, 0);
            //}
            //else if (saiyoudate.Value < s52ki)
            //{
            //     kimatsuday = new System.DateTime(2023, 03, 31, 0, 0, 0, 0);
            //}
            //else if (saiyoudate.Value < s53ki)
            //{
            //    kimatsuday = new System.DateTime(2024, 03, 31, 0, 0, 0, 0);
            //}
            //else if (saiyoudate.Value < s54ki)
            //{
            //    kimatsuday = new System.DateTime(2025, 03, 31, 0, 0, 0, 0);
            //}
            //else if (saiyoudate.Value < s55ki)
            //{
            //    kimatsuday = new System.DateTime(2026, 03, 31, 0, 0, 0, 0);
            //}

            //共通設定
            //if (keiyakunengetsu.Value.Equals(DBNull.Value)) keiyakunengetsu.Value = saiyoudate.Value; //作成日　なし

            if (kyuuyo.Text.Substring(0, 2) == "F1")
            {
                //アルバイトの場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                //if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                //if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = saiyoudate.Value.AddMonths(6); //雇用区分 入社日から6ヶ月後
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 3; //更新区分　契約の更新はしない
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 2; //休日勤務　なし
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 2; //時間外労働　なし
            }
            else if (kyuuyo.Text.Substring(0, 2) == "E1")
            {
                //パートの場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                //if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                //if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = kimatsuday; //雇用区分 期末日
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 3; //更新区分　契約の更新はしない
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 2; //休日勤務　なし
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 2; //時間外労働　なし

            }
            else if (kyuuyo.Text.Substring(0, 2) == "D1")
            {
                //日給者の場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 2; //雇用区分 定めなし
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 1; //更新区分　自動更新
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり
            }
            else if (kyuuyo.Text.Substring(0, 2) == "C1")
            {
                //月給者の場合
                //　TODO契約社員の場合
                //　TODO60オーバーの場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 2; //雇用区分 定めなし
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 1; //更新区分　自動更新
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 2; //定年　あり
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 1; //賞与　あり
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 1; //退職金　あり
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり
            }
            else
            {
                //兼務役員の場合
                //役員の場合
            }


        }

        private void lbl_syaryou_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syaryou.Text == syaryou_m.Text)
            {
                lbl_syaryou.BackColor = Color.White;
            }
            else
            {
                lbl_syaryou.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void syaryou_m_TextChanged(object sender, EventArgs e)
        {
            syaryou.Text = syaryou_m.Text;
        }

        private void tuushin_m_Click(object sender, EventArgs e)
        {

        }

        private void syaryou_TextChanged(object sender, EventArgs e)
        {
            //表示・非表示
            if (syaryou.Text == syaryou_m.Text)
            {
                lbl_syaryou.BackColor = Color.White;
            }
            else
            {
                lbl_syaryou.BackColor = Color.LightGreen;
            }

            Disp();
        }

        private void dispdgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void idoudaynew_ValueChanged(object sender, EventArgs e)
        {
            if (idoudaynew.Value == DBNull.Value) return;

            //時給と日給の最賃設定
            DataTable tbsai = Com.GetDB("select * from dbo.HK_本給 where '" + Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

            saichin = Convert.ToDecimal(tbsai.Rows[0]["最賃"].ToString());

            //zikyuu.Minimum = Convert.ToInt32(tbsai.Rows[0]["最賃"].ToString());
            //nikkyuu.Minimum = Convert.ToInt32(tbsai.Rows[0]["最賃"].ToString()) * 8;
        }

        private void zikyuu_ValueChanged(object sender, EventArgs e)
        {
            if (zikyuu_m.Text == "") return;

            //表示・非表示
            if (zikyuu.Value == Convert.ToDecimal(zikyuu_m.Text))
            {
                label4.BackColor = Color.White;
            }
            else
            {
                label4.BackColor = Color.LightGreen;
            }


            double hon = 0;

            if (zikyuu.Text != "0")
            {
                //日給者・パート・アルバイトの本給に暫定給を表示
                hon = Math.Round(Convert.ToInt32(zikyuu.Value) * Convert.ToInt32(kinmu.Text) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;

                //最賃きってたら、最賃に変更
                if (zikyuu.Value < saichin) zikyuu.Value = saichin;
            }
        }

        private void nikkyuu_ValueChanged(object sender, EventArgs e)
        {
            if (nikkyuu_m.Text == "") return;

            //表示・非表示
            if (nikkyuu.Value == Convert.ToDecimal(nikkyuu_m.Text))
            {
                label5.BackColor = Color.White;
            }
            else
            {
                label5.BackColor = Color.LightGreen;
            }

            double hon = 0;

            if (nikkyuu.Text != "0")
            {           
                //日給者・パート・アルバイトの本給に暫定給を表示
                hon = Math.Round(Convert.ToInt32(nikkyuu.Value) * Getrday(kyuuka.SelectedItem.ToString()));
                honkyuu.Value = Convert.ToDecimal(hon);
                honkyuu.ForeColor = Color.Red;

                //最賃きってたら、最賃に変更
                //TODO 一旦コメントアウト20250820
                //if (nikkyuu.Value < saichin * 8) nikkyuu.Value = saichin*8;
            }
        }
    }
}
