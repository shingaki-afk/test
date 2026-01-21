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
    public partial class Nyuusya : Form
    {
        private string nl = Environment.NewLine;
        private DataTable soshi = new DataTable();
        private DataTable yuubinno = new DataTable();


        //選択行
        private int dgvRow = 0;

        //職種
        private DataTable Syoku = new DataTable();

        //年齢
        private DataTable Nen = new DataTable();

        //社外経験
        DataTable Keiken = new DataTable();

        //学歴
        DataTable Gaku = new DataTable();

        //本給　TODO
        //private string honkyuu_ = "142,000";

        //TODO test
        int hatsurei_ex = 0;

        //null対応
        private DateTime zerodt = new System.DateTime(2022, 12, 31, 0, 0, 0, 0);

        public Nyuusya()
        {


            if (Convert.ToInt16(Program.access) == 1)
            {
                MessageBox.Show("入力権限がありません。");
                Com.InHistory("30_入社入力権限無", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            shikakudgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            kazokudgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //選択モードを行単位での選択のみにする
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            shikakudgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            kazokudgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //ソート不可対応
            foreach (DataGridViewColumn c in dataGridView1.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            foreach (DataGridViewColumn c in shikakudgv.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            foreach (DataGridViewColumn c in kazokudgv.Columns)
                c.SortMode = DataGridViewColumnSortMode.Programmatic;

            //初期設定　値
            IniSet();

            //初期設定　表示非表示
            IniSet2();

            //基本給設定
            GetDataKihonkyuu();

            GetData();

            ToolTip();

            yuubinno = Com.GetDB("select * from dbo.郵便番号");

            Com.InHistory("30_入社入力", "", "");

            //タブ非表示！！
            this.tabControl1.TabPages.Remove(this.roudoutab);

        }

        private void ToolTip()
        {   
            //ToolTipの設定を行う
            //ToolTipが表示されるまでの時間
            toolTip1.InitialDelay = 200;
            //ToolTipが表示されている時に、別のToolTipを表示するまでの時間
            toolTip1.ReshowDelay = 1000;
            //ToolTipを表示する時間
            toolTip1.AutoPopDelay = 30000;
            //フォームがアクティブでない時でもToolTipを表示する
            toolTip1.ShowAlways = true;

            //ToolTip1.SetToolTip(Button1, "このボタンは、\nButton1です");
            //toolTip1.SetToolTip(syokumulbl, "【職務給】\n15,000円　01 現業\n15,000円　02 客室\n10,000円　03 施設\n         0円　04 警備\n30,000円　05 エンジ\n  5,000円　06 サービス\n30,000円　07 営業\n14,000円　08 管理事務");
        }

        //
        private void IniSet()
        {
            //TODO 毎年追加しなければならない
            //ymlist.Items.Add("2020/04");
            //ymlist.Items.Add("2020/05");
            //ymlist.Items.Add("2020/06");
            //ymlist.Items.Add("2020/07");
            //ymlist.Items.Add("2020/08");
            //ymlist.Items.Add("2020/09");
            //ymlist.Items.Add("2020/10");
            //ymlist.Items.Add("2020/11");
            //ymlist.Items.Add("2020/12");
            //ymlist.Items.Add("2021/01");
            //ymlist.Items.Add("2021/02");
            //ymlist.Items.Add("2021/03");
            //ymlist.Items.Add("2021/04");
            //ymlist.Items.Add("2021/05");
            //ymlist.Items.Add("2021/06");
            //ymlist.Items.Add("2021/07");
            //ymlist.Items.Add("2021/08");
            //ymlist.Items.Add("2021/09");
            //ymlist.Items.Add("2021/10");
            //ymlist.Items.Add("2021/11");
            //ymlist.Items.Add("2021/12");
            //ymlist.Items.Add("2022/01");
            //ymlist.Items.Add("2022/02");

            //TODO 3月処理おわったらコメントアウトする。
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


            //TODO 21年4月 一時的に変更
            //ymlist.SelectedIndex = 0;
            ymlist.SelectedIndex = ymlist.FindString(saiyoudate.Value.ToString("yyyy/MM"));

            honkyuu.Text = "0";
            syokumukyuu.Text = "0";
            tokuteate.Text = "0";
            yakuteate.Text = "0";
            menkyoteate.Text = "0";
            huyou.Text = "0";
            ritou.Text = "0";
            kizyunnnai.Text = "0";

            tenkin.Text = "0";
            tuukinhi.Text = "0";
            tuukinka.Text = "0";
            tourokuteate.Text = "0";
            tuushinteate.Text = "0";
            shikyuu.Text = "0";

            for (int i = 10; i < 64; i++)
            {
                warekibirth.Items.Add("昭和" + i + "年");
            }

            warekibirth.Items.Add("平成元年");

            for (int i = 2; i < 31; i++)
            {
                warekibirth.Items.Add("平成" + i + "年");
            }

            //家族の方
            for (int i = 1; i < 64; i++)
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


            wareki.Items.Add("");
            wareki2.Items.Add("");
            wareki3.Items.Add("");
            wareki4.Items.Add("");


            for (int i = 20; i < 31; i++)
            {
                wareki.Items.Add("平成" + i + "年");
            }

            wareki.Items.Add("令和元年(平成31年)");

            for (int i = 2; i < 30; i++)
            {
                wareki.Items.Add("令和" + i + "年(平成" + (i + 30) + "年)");
            }



            for (int i = 20; i < 31; i++)
            {
                wareki2.Items.Add("平成" + i + "年");
            }

            wareki2.Items.Add("令和元年(平成31年)");

            for (int i = 2; i < 30; i++)
            {
                wareki2.Items.Add("令和" + i + "年(平成" + (i + 30) + "年)");
            }



            for (int i = 20; i < 31; i++)
            {
                wareki3.Items.Add("平成" + i + "年");
            }

            wareki3.Items.Add("令和元年(平成31年)");

            for (int i = 2; i < 30; i++)
            {
                wareki3.Items.Add("令和" + i + "年(平成" + (i + 30) + "年)");
            }



            for (int i = 20; i < 31; i++)
            {
                wareki4.Items.Add("平成" + i + "年");
            }

            wareki4.Items.Add("令和元年(平成31年)");

            for (int i = 2; i < 30; i++)
            {
                wareki4.Items.Add("令和" + i + "年(平成" + (i + 30) + "年)");
            }


            //状況
            status.Items.Add("");
            status.Items.Add("01　不採用");
            status.Items.Add("02　取消");
            //status.Items.Add("03　取消");

            //入社のきっかけ
            kikkake.Items.Add("");
            kikkake.Items.Add("1　縁故紹介");
            kikkake.Items.Add("2　求人誌");
            kikkake.Items.Add("3　ハローワーク");
            kikkake.Items.Add("4　折込チラシ");
            kikkake.Items.Add("5　新聞広告");
            kikkake.Items.Add("6　その他");

            //発令区分
            hatsurei.Items.Add("0900　パート契約");
            hatsurei.Items.Add("1000　アルバイト契約");
            hatsurei.Items.Add("1100　日給者契約");　//2022年04対応
            hatsurei.Items.Add("0001　正社員採用");

            //性別
            seibetsu.Items.Add("1　男性");
            seibetsu.Items.Add("2　女性");

            //地区コード
            tiku.Items.Add("1　本社");
            tiku.Items.Add("2　那覇");
            tiku.Items.Add("3　八重山");
            tiku.Items.Add("4　北部");
            tiku.Items.Add("5　広域");
            tiku.Items.Add("6　宮古島");
            tiku.Items.Add("7　久米島");

            //国籍
            kokuseki.Items.Add("");

            //友の会区分
            tomokubun.Items.Add("");
            tomokubun.Items.Add("1　非加入");
            tomokubun.Items.Add("2　アルバイト加入");

            //税表区分
            zeikubun.Items.Add("1" + "　" + "甲");
            zeikubun.Items.Add("2" + "　" + "乙");
            zeikubun.Items.Add("3" + "　" + "非居住");
            zeikubun.SelectedIndex = 0;

            //障害区分
            syougai.Items.Add("");
            syougai.Items.Add("1" + "　" + "普通");
            syougai.Items.Add("2" + "　" + "特別");

            //寡フ区分
            kahu.Items.Add("");
            kahu.Items.Add("1" + "　" + "寡フ");
            kahu.Items.Add("2" + "　" + "ひとり親");

            //勤労　外国人　災害
            gakusei.Items.Add("");
            gakusei.Items.Add("1" + "　" + "○");
            saigai.Items.Add("");
            saigai.Items.Add("1" + "　" + "○");
            gaikoku.Items.Add("");
            gaikoku.Items.Add("1" + "　" + "○");

            yakusyoku.Text = "0180　係員";

            //契約社員設定
            keiyaku.Items.Add("");
            keiyaku.Items.Add("10" + "　" + "一般契約社員");
            keiyaku.Items.Add("20" + "　" + "単年契約社員");
            keiyaku.Items.Add("30" + "　" + "技能実習生");
            keiyaku.Items.Add("31" + "　" + "特技能実習生");

            //休暇付与区分
            kyuuka.Items.Add("0" + "　" + "５日以上");
            kyuuka.Items.Add("1" + "　" + "４日");
            kyuuka.Items.Add("2" + "　" + "３日");
            kyuuka.Items.Add("3" + "　" + "２日");
            kyuuka.Items.Add("4" + "　" + "１日");
            kyuuka.Items.Add("9" + "　" + "付与なし");
            kyuuka.SelectedIndex = 0;

            //勤務時間
            kinmu.Items.Add("8");
            kinmu.Items.Add("7");
            kinmu.Items.Add("6");
            kinmu.Items.Add("5");
            kinmu.Items.Add("4");
            kinmu.Items.Add("3");
            kinmu.Items.Add("2");
            kinmu.Items.Add("1");
            kinmu.SelectedIndex = 0;

            //通勤手段区分
            tuukinkubun.Items.Add("1 車");
            tuukinkubun.Items.Add("2 バイク");
            //tuukinkubun.Items.Add("3 徒歩・自転車");
            tuukinkubun.Items.Add("4 バス・モノレール");
            tuukinkubun.Items.Add("5 送迎(会社)");
            tuukinkubun.Items.Add("6 送迎(知人・親族)");
            tuukinkubun.Items.Add("7 業務車両");
            tuukinkubun.Items.Add("8 徒歩");
            tuukinkubun.Items.Add("9 自転車");
            //tuukinkubun.Items.Add("8 車(実費精算)");
            tuukinkubun.SelectedIndex = -1;

            //通勤手当区分
            tuukinteatekubun.Items.Add("");
            tuukinteatekubun.Items.Add("1 実費精算");
            ////tuukinteatekubun.Items.Add("3 支給対象外");
            tuukinteatekubun.SelectedIndex = 0;

            ////資格
            //DataTable shikakudt = new DataTable();
            //shikakudt = Com.GetDB("select 管理コード +'　' + 摘要 as 資格 from QUATRO.dbo.QCMTCODED c where c.情報キー = 'SJMT095' and c.適用終了日 = '9999/12/31'");

            //foreach (DataRow drw in shikakudt.Rows)
            //{
            //    shikakucombo.Items.Add(drw["資格"].ToString());
            //}


            //家族タブ
            //続柄区分
            DataTable zokugaradt = new DataTable();
            zokugaradt = Com.GetDB("select 管理コード + '　' + 摘要 as 続柄 from QUATRO.dbo.QCMTCODED where 情報キー = 'SJMT030' and 適用終了日 = '9999/12/31' and 管理コード <> '0'");
            foreach (DataRow drw in zokugaradt.Rows)
            {
                zokugara.Items.Add(drw["続柄"].ToString());
            }

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

            //転勤内容
            comboBoxTenkin.Items.Add("");
            comboBoxTenkin.Items.Add("1　本島⇒離島");
            comboBoxTenkin.Items.Add("2　離島⇒本島");
            comboBoxTenkin.Items.Add("3　離島⇒離島");

            //休日区分
            kyuuzitsukubun.Items.Add("10　年間最低数");
            kyuuzitsukubun.Items.Add("20　土日祝");
            syougaikubun.SelectedIndex = 0;

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            DataTable hatsu = new DataTable();
            DataTable yaku = new DataTable();
            DataTable tou = new DataTable();
            DataTable gou = new DataTable();

            DataTable koku = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        //組織一覧
                        //string sql = "select distinct a.組織CD, a.組織名 from dbo.社員基本情報 a left join dbo.担当テーブル b on a.組織CD = b.組織CD and a.現場CD = b.現場CD where 在籍区分 <> '9' and 担当区分 ";
                        string sql = "select distinct a.組織CD, a.組織名 from dbo.担当テーブル a where 定員数 > 0　and ";

                        if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
                        {
                            sql = sql + "(担当区分  in ('03_施設', '04_エンジ') or (担当区分 in ('14_宮古島','15_久米島') and 担当事務 in ('03_施設', '04_エンジ', '03_警備')))";
                        }
                        else if (Program.loginname == "金城　智之")
                        {
                            sql = sql + "担当区分  like '%%' ";
                        }
                        //TODO 2503大濱さん宮古島応援のため
                        else if (Program.loginname == "大浜　綾希子")
                        {
                            sql = sql + "(担当区分  in ('01_現業') or (担当区分 in ('15_久米島') and 担当事務 = '01_現業')) ";
                        }
                        else
                        {
                            sql = sql + "担当区分  like '%" + Program.loginbusyo + "%'";
                        }

                        Cmd.CommandText = sql;

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(soshi);

                        //国籍一覧
                        //TODO : 優先
                        kokuseki.Items.Add("NP" + "　" + "ネパール");
                        kokuseki.Items.Add("VN" + "　" + "ベトナム");
                        kokuseki.Items.Add("PH" + "　" + "フィリピン");
                        kokuseki.Items.Add("ID" + "　" + "インドネシア");

                        Cmd.CommandText = "select 管理コード, 摘要 from QUATRO.dbo.QCMTCODED where 適用終了日 = '9999/12/31' AND 情報キー = 'KJMT008' order by 管理コード";
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(koku);

                        foreach (DataRow drw in koku.Rows)
                        {
                            if (drw["管理コード"].ToString() == "NP" || 
                                drw["管理コード"].ToString() == "VN" || 
                                drw["管理コード"].ToString() == "PH" || 
                                drw["管理コード"].ToString() == "ID")
                            {
                                //スルー
                            }
                            else
                            {
                                kokuseki.Items.Add(drw["管理コード"].ToString() + "　" + drw["摘要"].ToString());
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

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

            //厚生年金
            kouseicb.Items.Add("");
            kouseicb.Items.Add("1 有");
            kouseicb.Items.Add("2 無");

            //健康保険
            kenkoucb.Items.Add("");
            kenkoucb.Items.Add("1 有");
            kenkoucb.Items.Add("2 無");

            //雇用保険
            koyoucb.Items.Add("");
            koyoucb.Items.Add("1 有");
            koyoucb.Items.Add("2 無");

            //休日回数 
            kyuujitsukaisuu.Items.Add("");
            kyuujitsukaisuu.Items.Add("1 年間107日　月変形週労40時間  ※2月が29日の暦日の場合は108日");
            kyuujitsukaisuu.Items.Add("2 1ヶ月につき  4日～10日  (週5以上勤務)"); //週5
            kyuujitsukaisuu.Items.Add("3 1ヶ月につき 12日～15日  (週4勤務)"); //週4
            kyuujitsukaisuu.Items.Add("4 1ヶ月につき 16日～20日  (週3勤務)"); //週3
            kyuujitsukaisuu.Items.Add("5 1ヶ月につき 20日～25日  (週2勤務)"); //週2
            kyuujitsukaisuu.Items.Add("6 1ヶ月につき 23日～27日  (週1勤務)"); //週1
        }

        private void IniSet2()
        {
            //資格タブと家族タブ
            tabControl1.TabPages.Remove(this.kazokutab);
            tabControl1.TabPages.Remove(this.shikakutab);

            //職種学歴経験パネル非表示
            seisyapanel.Visible = false;

            //日給者の金額非表示対応
            syokumukyuu.Visible = true;
            gakurekikyuu.Visible = true;
            keikenkyuu.Visible = true;
            nennreikyuu.Visible = true;


            //縁故パネル非表示
            enkopanel.Visible = false;

            //生年月日
            birthnew.Value = null;

            tabControl1.Visible = false;

            mycarkigenpanel.Visible = false;

            syuturyoku.Visible = false;

            //label95.Visible = false;
            error.Visible = false;

            //label96.Visible = false;
            status.Visible = false;
        }

        private void GetData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();
            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        string sql = "select * from dbo.n入社データ一覧取得 where ";

                        if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
                        {
                            sql = sql + "(担当区分  in ('03_施設', '04_エンジ') or (担当区分 in ('14_宮古島','15_久米島') and 担当事務 in ('03_施設', '04_エンジ', '03_警備')))";
                        }
                        else if (Program.loginname == "金城　智之")
                        {
                            //TODO
                            sql = sql + "担当区分  like '%%' ";
                        }
                        //TODO 2503大濱さん宮古島応援のため
                        else if (Program.loginname == "大浜　綾希子")
                        {
                            sql = sql + "(担当区分  in ('01_現業') or (担当区分 in ('15_久米島') and 担当事務 = '01_現業')) ";
                        }
                        else if (Program.loginname == "佐久間　みどり")
                        {
                            sql = sql + "(担当区分  in ('02_客室','14_宮古島')) ";
                        }
                        else
                        {
                            sql = sql + "担当区分 like '%" + Program.loginbusyo + "%'";
                        }

                        //年月絞り込み
                        sql = sql + " and 入社年月日 like '" + ymlist.SelectedItem + "%'";

                        Cmd.CommandText = sql + " order by 社員番号";
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

        //TODO
        private void Nyuusya_Load(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
        }

        //入社書類項目
        private void SyoruiClear()
        {
            hoyuurbmu.Checked = true;
            huyourbmu.Checked = true;

            c01.Checked = false;
            c02.Checked = false;
            c03.Checked = false;
            c04.Checked = false;
            c05.Checked = false;
            c06.Checked = false;
            c07.Checked = false;
            c08.Checked = false;
            c09.Checked = false;
            c10.Checked = false;
            c11.Checked = false;
            c12.Checked = false;

            c13.Checked = false;
            c14.Checked = false;
            c15.Checked = false;

            c16.Checked = false;
            //c17.Checked = false;
            c18.Checked = false;

            c19.Checked = false;
            c20.Checked = false;

            tuukinteatekubun.SelectedIndex = 0; //通勤手当区分
            tuukinkubun.SelectedIndex = -1; //通勤手段区分
            katakyori.Value = 0; //片道通勤距離
            //kataryoukin.Value = 0; //片道料金

            menkyonew.Value = null; //免許
            syakennew.Value = null; //車検
            zibainew.Value = null; //自賠責
            ninninew.Value = null; //任意

            kikkake.SelectedIndex = -1;
            enko_no.Text = "";
            enko_name.Text = "";
            enko_soshiki.Text = "";
            enko_genba.Text = "";

            yuubin.Text = "";  //現郵便番号
            zyuusyo.Text = ""; //現住所
            zyuusyocheck.Checked = false; //住民登録別フラグ
            keitai.Text = "";  //携帯

            //有効無効
            c13.Enabled = false;
            c14.Enabled = false;
            c15.Enabled = false;

            c19.Enabled = false;
            c20.Enabled = false;

            bikou1.Text = "";
            bikou2.Text = "";
            bikou3.Text = "";
        }

        private void AllClear()
        {
            //入社書類項目
            SyoruiClear();

            syainno.Text = "";　//社員番号
            saiyoudate.Text = DateTime.Today.ToString(); //入社年月日
            hatsurei.SelectedIndex = -1; //発令区分

            seihuri.Text = ""; //姓フリ
            meihuri.Text = ""; //名フリ
            sei.Text = ""; // 姓
            mei.Text = ""; //名

            tiku.SelectedIndex = -1; //地区
            soshiki.SelectedIndex = -1; //所属
            genba.SelectedIndex = -1; //現場

            kyuuka.SelectedIndex = -1; //休暇付与
            kinmu.SelectedIndex = -1; //勤務時間

            status.SelectedIndex = -1; //状況
            error.Text = ""; //警告とエラー

            warekibirth.SelectedIndex = -1;
            birthnew.Value = null;　//生年月日
            old.Text = "0"; //年齢 TODO ゼロから""に変更したい
            seibetsu.SelectedIndex = -1; //性別

            zikyuu.Value = zikyuu.Minimum; //時給
            zikyuu.Visible = true;

            nikkyuu.Value = nikkyuu.Minimum; //日給
            nikkyuu.Visible = true;

            kaisuu1.Text = ""; //回数1
            //kaisuu2.Text = ""; //回数2

            keiyaku.SelectedIndex = -1; //契約社員
            keiyaku.Visible = true;

            kyuuzitsukubun.SelectedIndex = 0; //休日区分 デフォルト設定

            tomokubun.SelectedIndex = -1; //友の会区分
            kokuseki.SelectedIndex = -1; //国籍

            zeikubun.SelectedIndex = 0; //税区分 デフォルト設定
            syougai.SelectedIndex = -1; //障害区分
            kahu.SelectedIndex = -1;　//寡フ
            gakusei.SelectedIndex = -1; //学生
            saigai.SelectedIndex = -1; //災害
            gaikoku.SelectedIndex = -1; //外国



            //入社最初は全員係員
            //yakusyoku.Text = ""; //役職
            kyuuyo.Text = ""; //給与区分
            //.Text = ""; //社員区分⇒休日区分
            shiyoukikan.Text = ""; //試用期間

            comboBoxSyokusyu.SelectedIndex = -1; //職種
            comboBoxGaku.SelectedIndex = -1; //学歴
            comboBoxKeiken.SelectedIndex = -1; //社外経験

            honkyuu.Text = "0"; //本給
            syokumu.Text = "0"; //職務技能給
            tokuteate.Text = "0"; //特別手当
            yakuteate.Text = "0"; //役職手当
            menkyoteate.Text = "0"; //免許手当
            huyou.Text = "0"; //扶養手当
            ritou.Text = "0"; //離島手当
            kizyunnnai.Text = "0"; //基準内賃金
            tenkin.Text = "0"; //転勤手当
            tuukinhi2.Value = 0; //通勤非
            tuukinka2.Value = 0; //通勤
            tourokuteate.Text = "0"; //登録手当
            tuushinteate.Text = "0"; //通信手当
            shikyuu.Text = "0"; //支給合計

            tomonokai.Text = ""; //友の会
            //kouzyogoukei.Text = ""; //控除合計

            tokureason.Text = ""; //特別手当付与理由


            comboBoxTenkin.SelectedIndex = -1; //赴任元・先

            syokumukyuu.Text = "0";
            gakurekikyuu.Text = "0";
            keikenkyuu.Text = "0";
            nennreikyuu.Text = "0";

            shikyuu.Text = "0";
            tuukinhi.Text = "0";
            tuukinka.Text = "0";

            //家族情報クリア
            KazokuClear();
            //扶養手当額
            huyougaku.Text = "0";

            //資格情報クリア
            ShikakuClear();
            //免許手当額
            menkyogaku.Text = "0";
            //登録手当額
            tourokuteate.Text = "0";

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

            kouseicb.Text = null;
            kenkoucb.Text = null;
            koyoucb.Text = null;

            //緊急連絡先リセット
            honkeitai.Text = null;
            honkotei.Text = null;

            kaz1name.Text = null;
            kaz1kana.Text = null;
            kaz1gara.Text = null;
            kaz1no.Text = null;

            kaz2name.Text = null;
            kaz2kana.Text = null;
            kaz2gara.Text = null;
            kaz2no.Text = null;
        }


        private void KazokuClear()
        {
            //TODO　データグリッドダブルクリック時に消されるとこまる
            //kazokudgv.DataSource = "";
            kazomei.Text = "";
            kazosei.Text = sei.Text;
            kazokanamei.Text = "";
            kazokanasei.Text = seihuri.Text;

            kazokuid.Text = "";

            warekicb.SelectedIndex = -1;
            //kazoseinengappi.Text = DateTime.Today.ToString();
            kazoseinengappinew.Value = null;
            zokugara.SelectedIndex = -1;
            doukyokubun.SelectedIndex = 1;
            huyoukubun.SelectedIndex = 0;
            gensenkubun.SelectedIndex = 0;
            kenpokanyuu.SelectedIndex = 0;

            setainushi.SelectedIndex = 0;
            syougaikubun.SelectedIndex = 0;

            kazokubtn.Text = "家族登録";
            kazokubtn.BackColor = Color.Transparent;
            kazokubtn.ForeColor = Color.Black;

            delkazoku.Visible = false;
        }

        private void ShikakuClear()
        {
            //shikakucombo.SelectedIndex = -1;
            shikakutextb.Text = "";
            shikakusyutokubi.Text = DateTime.Today.ToString();
            shikakuno.Text = "";
            shikakubtn.Text = "資格登録";
            shikakukigenday.Text = DateTime.Today.ToString();
            shikakubtn.BackColor = Color.Transparent;
            shikakubtn.ForeColor = Color.Black;

            //shikakuselect.Enabled = true;

            delshikaku.Visible = false;

        }





        private void tiku_SelectedIndexChanged(object sender, EventArgs e)
        {
            soshiki.Items.Clear();

            if (tiku.SelectedIndex == -1) return;
            DataRow[] dr = soshi.Select("組織CD like '" + (tiku.SelectedItem).ToString().Substring(0,1) + "%'");

            foreach (DataRow drw in dr)
            {
                soshiki.Items.Add(drw["組織CD"].ToString() + "　" + drw["組織名"].ToString());
            }
        }

        private void soshiki_SelectedIndexChanged(object sender, EventArgs e)
        {
            genba.Items.Clear();
            gen.Clear();

            //DataRow[] dr = soshiki.Select("組織CD like '" + (comboBox2.SelectedIndex + 1).ToString() + "%'");

            GetGenba();
            foreach (DataRow drw in gen.Rows)
            {
                genba.Items.Add(drw["現場CD"].ToString() + "　" + drw["現場名"].ToString());
            }


            if (tenkin.Text == "")
            {

            }
            else if (Convert.ToDecimal(tenkin.Text) > 0)
            {
                ritou.Text = "0";
                return;
            }
            if (soshiki.SelectedItem == null || genba.SelectedItem == null || comboBoxSyokusyu.SelectedItem == null) return;
            if (kyuuyo.Text.Substring(0, 2) != "C1") return;
            ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
        }

        private DataTable gen = new DataTable();
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

                        //担当区分
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(gen);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string msg = "";

            //発令区分
            if (hatsurei.Text == "") msg += "発令区分は必須です。" + nl;
            if (sei.Text == "" || mei.Text == "" || seihuri.Text == "" || meihuri.Text == "") msg += "名前は必須です。" + nl;
            if (genba.Text == "") msg += "地区名・組織名・現場名は必須です。" + nl;

            if (kinmu.Text == "") msg += "基本勤務時間は必須です。" + nl;
            if (kyuuka.Text == "") msg += "週労数は必須です。" + nl;

            if (Convert.ToInt32(tokuteate.Text) > 0 && tokureason.Text.Length == 0) msg += "特別手当支給場合は、特別手当理由は必須です。※別途稟議決裁も必要です。" + nl;

            if (msg != "")
            {
                MessageBox.Show(msg);
                return;
            }

            string no = DataInsertUpdate();

            //登録後のフォーカスに利用
            //TODO 一覧表示は登録順または追加は必ず最後の行に入らなければならない！
            if (dataGridView1.CurrentCell == null)
            {

                dgvRow = dataGridView1.Rows.Count;
            }
            else
            {
                dgvRow = dataGridView1.CurrentCell.RowIndex;
            }

            ymlist.SelectedIndex = ymlist.FindString(saiyoudate.Value.ToString("yyyy/MM")); 

            //一覧データ取得
            GetData();

            if (no != "")
            {
                dgvRow = dataGridView1.Rows.Count - 1;
            }

            dataGridView1.CurrentCell = dataGridView1[1, dgvRow];

            AllClear();
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            DataDisp(drv[1].ToString());

            if (no == "")
            {
                MessageBox.Show("更新しました。");

            }
            else
            {
                MessageBox.Show("登録しました。 社員番号は" + syainno.Text + "です。");
            }
        }

        private string DataInsertUpdate()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            SqlDataAdapter da2;
            SqlDataAdapter da3;
            SqlDataAdapter da4;
            SqlDataAdapter da5;
            string no = "";

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    Cn.Open();
                    using (Cmd = Cn.CreateCommand())
                    {
                        //Null対応
                        string menkyo = menkyonew.Text == "" ? "Null" : "'" + menkyonew.Value + "'";
                        string syaken = syakennew.Text == "" ? "Null" : "'" + syakennew.Value + "'";
                        string zibai = zibainew.Text == "" ? "Null" : "'" + zibainew.Value + "'";
                        string ninni = ninninew.Text == "" ? "Null" : "'" + ninninew.Value + "'";


                        if (syainno.Text == "")
                        {
                            DataTable dt = new DataTable();
                            Cmd.CommandText = "select top 1 社員番号 from dbo.新入社社員番号 where 取得日 is null and 社員番号 like '2" + saiyoudate.Value.ToString("yy") + "%' order by 社員番号";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(dt);

                            foreach (DataRow drw in dt.Rows)
                            {
                                no = drw["社員番号"].ToString();
                            }

                            DataTable dt2 = new DataTable();
                            Cmd.CommandText = "UPDATE dbo.新入社社員番号 SET 取得日 = GETDATE(), 取得者 = '" + Program.loginname + "' WHERE 社員番号 = '" + no + "'";
                            da2 = new SqlDataAdapter(Cmd);
                            da2.Fill(dt2);

                            //通勤管理テーブルへインサート
                            DataTable dt3 = new DataTable();
                            Cmd.CommandText = "insert into dbo.t通勤管理テーブル(社員番号, 氏名, 管理No) VALUES('" + no + "', '" + sei.Text + "　" + mei.Text + "', '1')";
                            da3 = new SqlDataAdapter(Cmd);
                            da3.Fill(dt3);

                            //通勤管理テーブルへインサート
                            DataTable dt4 = new DataTable();
                            Cmd.CommandText = "insert into dbo.t通勤手当元データ(社員番号, 適用開始日, 適用終了日) VALUES('" + no + "', '" + saiyoudate.Value.ToString("yyyy/MM/dd") + "', '9999/12/31')";
                            da4 = new SqlDataAdapter(Cmd);
                            da4.Fill(dt4);

                            //労働条件テーブルへインサート
                            DataTable dt5 = new DataTable();
                            Cmd.CommandText = "insert into dbo.r労働条件(社員番号) VALUES('" + no + "')";
                            da5 = new SqlDataAdapter(Cmd);
                            da5.Fill(dt5);
                        }
                        else
                        {
                            //緊急連絡先の更新
                            SqlDataReader drkin;

                            Cmd.CommandType = CommandType.StoredProcedure;
                            Cmd.CommandText = "[dbo].[k緊急連絡先更新]";

                            Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("携帯番号", SqlDbType.VarChar)); Cmd.Parameters["携帯番号"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("固定電話", SqlDbType.VarChar)); Cmd.Parameters["固定電話"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("氏名1", SqlDbType.VarChar)); Cmd.Parameters["氏名1"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("カナ名1", SqlDbType.VarChar)); Cmd.Parameters["カナ名1"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("続柄1", SqlDbType.VarChar)); Cmd.Parameters["続柄1"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("電話番号1", SqlDbType.VarChar)); Cmd.Parameters["電話番号1"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("氏名2", SqlDbType.VarChar)); Cmd.Parameters["氏名2"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("カナ名2", SqlDbType.VarChar)); Cmd.Parameters["カナ名2"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("続柄2", SqlDbType.VarChar)); Cmd.Parameters["続柄2"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("電話番号2", SqlDbType.VarChar)); Cmd.Parameters["電話番号2"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("最終更新日", SqlDbType.DateTime)); Cmd.Parameters["最終更新日"].Direction = ParameterDirection.Input;
                            Cmd.Parameters.Add(new SqlParameter("最終更新者", SqlDbType.VarChar)); Cmd.Parameters["最終更新者"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                            Cmd.Parameters["社員番号"].Value = syainno.Text;

                            Cmd.Parameters["携帯番号"].Value = honkeitai.Text;
                            Cmd.Parameters["固定電話"].Value = honkotei.Text;

                            Cmd.Parameters["氏名1"].Value = kaz1name.Text;
                            Cmd.Parameters["カナ名1"].Value = kaz1kana.Text;
                            Cmd.Parameters["続柄1"].Value = kaz1gara.Text;
                            Cmd.Parameters["電話番号1"].Value = kaz1no.Text;

                            Cmd.Parameters["氏名2"].Value = kaz2name.Text;
                            Cmd.Parameters["カナ名2"].Value = kaz2kana.Text;
                            Cmd.Parameters["続柄2"].Value = kaz2gara.Text;
                            Cmd.Parameters["電話番号2"].Value = kaz2no.Text;

                            Cmd.Parameters["最終更新日"].Value = DateTime.Now;
                            Cmd.Parameters["最終更新者"].Value = Program.loginname;

                            using (drkin = Cmd.ExecuteReader())
                            {
                                int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                            }


                            Cmd.CommandType = CommandType.Text;
                            Cmd.CommandText = "";

                            //通勤管理テーブル更新
                            DataTable dt4 = new DataTable();
                            string sql4 = "update dbo.t通勤管理テーブル set ";
                            sql4 += "通勤方法 = '" + tuukinkubun.Text + "',免許証 = " + menkyo + ",車検証 = " + syaken + ", 自賠責 = " + zibai + ", 任意保険 = " + ninni + ", ";
                            //sql4 += "通勤手当区分 = '" + tuukinteatekubun.Text + "', 片道距離 = '" + katakyori.Value + "',片道料金 = '', ";
                            sql4 += "メーカー = '', 車名 = '', 色 = '', 車両番号 = '', ";
                            sql4 += "備考 = '' where 社員番号 = '" + syainno.Text + "'";
                            Cmd.CommandText = sql4;
                            da4 = new SqlDataAdapter(Cmd);
                            da4.Fill(dt4);

                            //エラー対応
                            //string kkyori = "0";
                            //if (katakyori.Value != "")
                            //{
                            //    kkyori = katakyori.Value;
                            //}

                            string ttan = "0";
                            if (tuutanka.Text != "")
                            {
                                ttan = tuutanka.Text;
                            }

                            //通勤手当元データ更新
                            DataTable dt5 = new DataTable();
                            string sql5 = "update dbo.t通勤手当元データ set ";
                            sql5 += "通勤方法 = '" + tuukinkubun.Text + "',";
                            sql5 += "通勤手当区分 = '" + tuukinteatekubun.Text + "', 片道距離 = '" + katakyori.Value + "', 通勤1日単価 = '" + ttan + "',　適用開始日 = '" + saiyoudate.Value.ToString("yyyy/MM/dd") + "', 適用終了日 = '9999/12/31', ";
                            sql5 += "備考 = '' where 社員番号 = '" + syainno.Text + "'";
                            Cmd.CommandText = sql5;
                            da5 = new SqlDataAdapter(Cmd);
                            da5.Fill(dt5);
                        }

                        //Cmd = Cn.CreateCommand();
                        Cmd.CommandType = CommandType.StoredProcedure;

                        if (syainno.Text == "")
                        {
                            Cmd.CommandText = "[dbo].[n入社データ登録]";
                        }
                        else
                        {
                            Cmd.CommandText = "[dbo].[n入社データ更新]";
                        }

                        Cmd.Parameters.Clear();

                        Cmd.Parameters.Add(new SqlParameter("採用年月日", SqlDbType.Char)); Cmd.Parameters["採用年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("発令区分", SqlDbType.Char)); Cmd.Parameters["発令区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("姓フリ", SqlDbType.VarChar)); Cmd.Parameters["姓フリ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("名フリ", SqlDbType.VarChar)); Cmd.Parameters["名フリ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("姓", SqlDbType.VarChar)); Cmd.Parameters["姓"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("名", SqlDbType.VarChar)); Cmd.Parameters["名"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("生年月日", SqlDbType.Char)); Cmd.Parameters["生年月日"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("性別", SqlDbType.Char)); Cmd.Parameters["性別"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("携帯番号", SqlDbType.VarChar)); Cmd.Parameters["携帯番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現郵便番号", SqlDbType.VarChar)); Cmd.Parameters["現郵便番号"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現住所", SqlDbType.VarChar)); Cmd.Parameters["現住所"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("住所フラグ", SqlDbType.Char)); Cmd.Parameters["住所フラグ"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("地区", SqlDbType.Char)); Cmd.Parameters["地区"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("所属組織", SqlDbType.Char)); Cmd.Parameters["所属組織"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("現場名", SqlDbType.Char)); Cmd.Parameters["現場名"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("役職", SqlDbType.Char)); Cmd.Parameters["役職"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("給与支給区分", SqlDbType.Char)); Cmd.Parameters["給与支給区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社員区分", SqlDbType.Char)); Cmd.Parameters["社員区分"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("休日区分", SqlDbType.Char)); Cmd.Parameters["休日区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("試用期間", SqlDbType.Char)); Cmd.Parameters["試用期間"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("契約社員", SqlDbType.Char)); Cmd.Parameters["契約社員"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("友の会区分", SqlDbType.Char)); Cmd.Parameters["友の会区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("国籍", SqlDbType.Char)); Cmd.Parameters["国籍"].Direction = ParameterDirection.Input;

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
                        Cmd.Parameters.Add(new SqlParameter("特別手当", SqlDbType.Decimal)); Cmd.Parameters["特別手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("役職手当", SqlDbType.Decimal)); Cmd.Parameters["役職手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("免許手当", SqlDbType.Decimal)); Cmd.Parameters["免許手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("扶養手当", SqlDbType.Decimal)); Cmd.Parameters["扶養手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("離島手当", SqlDbType.Decimal)); Cmd.Parameters["離島手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("転勤手当", SqlDbType.Decimal)); Cmd.Parameters["転勤手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当非", SqlDbType.Decimal)); Cmd.Parameters["通勤手当非"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手当課", SqlDbType.Decimal)); Cmd.Parameters["通勤手当課"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("登録手当", SqlDbType.Decimal)); Cmd.Parameters["登録手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通信手当", SqlDbType.Decimal)); Cmd.Parameters["通信手当"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("友の会", SqlDbType.Decimal)); Cmd.Parameters["友の会"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("固定1", SqlDbType.Decimal)); Cmd.Parameters["固定1"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("固定2", SqlDbType.Decimal)); Cmd.Parameters["固定2"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("特別理由", SqlDbType.VarChar)); Cmd.Parameters["特別理由"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("赴任元先", SqlDbType.VarChar)); Cmd.Parameters["赴任元先"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("職種", SqlDbType.VarChar)); Cmd.Parameters["職種"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("最終学歴", SqlDbType.VarChar)); Cmd.Parameters["最終学歴"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("社外経験", SqlDbType.VarChar)); Cmd.Parameters["社外経験"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("通勤手当区分", SqlDbType.VarChar)); Cmd.Parameters["通勤手当区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("通勤手段区分", SqlDbType.VarChar)); Cmd.Parameters["通勤手段区分"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("片道通勤距離", SqlDbType.Decimal)); Cmd.Parameters["片道通勤距離"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("片道料金", SqlDbType.Decimal)); Cmd.Parameters["片道料金"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("01", SqlDbType.Char)); Cmd.Parameters["01"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("02", SqlDbType.Char)); Cmd.Parameters["02"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("03", SqlDbType.Char)); Cmd.Parameters["03"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("04", SqlDbType.Char)); Cmd.Parameters["04"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("05", SqlDbType.Char)); Cmd.Parameters["05"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("06", SqlDbType.Char)); Cmd.Parameters["06"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("07", SqlDbType.Char)); Cmd.Parameters["07"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("08", SqlDbType.Char)); Cmd.Parameters["08"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("09", SqlDbType.Char)); Cmd.Parameters["09"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("10", SqlDbType.Char)); Cmd.Parameters["10"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("11", SqlDbType.Char)); Cmd.Parameters["11"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("12", SqlDbType.Char)); Cmd.Parameters["12"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("13", SqlDbType.Char)); Cmd.Parameters["13"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("14", SqlDbType.Char)); Cmd.Parameters["14"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("15", SqlDbType.Char)); Cmd.Parameters["15"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("16", SqlDbType.Char)); Cmd.Parameters["16"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("17", SqlDbType.Char)); Cmd.Parameters["17"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("18", SqlDbType.Char)); Cmd.Parameters["18"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("19", SqlDbType.Char)); Cmd.Parameters["19"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("20", SqlDbType.Char)); Cmd.Parameters["20"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("21", SqlDbType.Char)); Cmd.Parameters["21"].Direction = ParameterDirection.Input;
                        //Cmd.Parameters.Add(new SqlParameter("22", SqlDbType.Char)); Cmd.Parameters["22"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("入社きっかけ", SqlDbType.Char)); Cmd.Parameters["入社きっかけ"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("縁故社員番号", SqlDbType.Char)); Cmd.Parameters["縁故社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("保有資格有無", SqlDbType.Char)); Cmd.Parameters["保有資格有無"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("扶養家族有無", SqlDbType.Char)); Cmd.Parameters["扶養家族有無"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("警告とエラー", SqlDbType.Char)); Cmd.Parameters["警告とエラー"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("削除理由", SqlDbType.Char)); Cmd.Parameters["削除理由"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考1", SqlDbType.Char)); Cmd.Parameters["備考1"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考2", SqlDbType.Char)); Cmd.Parameters["備考2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考3", SqlDbType.Char)); Cmd.Parameters["備考3"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("備考4", SqlDbType.VarChar)); Cmd.Parameters["備考4"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考5", SqlDbType.VarChar)); Cmd.Parameters["備考5"].Direction = ParameterDirection.Input;
                        Cmd.Parameters.Add(new SqlParameter("備考6", SqlDbType.VarChar)); Cmd.Parameters["備考6"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["採用年月日"].Value = saiyoudate.Value.ToString("yyyy/MM/dd");
                        Cmd.Parameters["発令区分"].Value = hatsurei.Text;
                        Cmd.Parameters["社員番号"].Value = syainno.Text == "" ? no : syainno.Text;

                        Cmd.Parameters["姓フリ"].Value = seihuri.Text;
                        Cmd.Parameters["名フリ"].Value = meihuri.Text;
                        Cmd.Parameters["姓"].Value = sei.Text;
                        Cmd.Parameters["名"].Value = mei.Text;

                        //Cmd.Parameters["生年月日"].Value = birthnew.Value.ToString("yyyy/MM/dd");
                        if (birthnew.Text == "")
                        {
                            Cmd.Parameters["生年月日"].Value = "";
                        }
                        else
                        {
                            Cmd.Parameters["生年月日"].Value = Convert.ToDateTime(birthnew.Text).ToString("yyyy/MM/dd");
                        }

                        Cmd.Parameters["性別"].Value = seibetsu.Text;

                        Cmd.Parameters["現郵便番号"].Value = yuubin.Text;
                        Cmd.Parameters["現住所"].Value = zyuusyo.Text;
                        Cmd.Parameters["住所フラグ"].Value = zyuusyocheck.Checked ? 1 : 0;
                        Cmd.Parameters["携帯番号"].Value = keitai.Text;

                        Cmd.Parameters["地区"].Value = tiku.Text;
                        Cmd.Parameters["所属組織"].Value = soshiki.Text;
                        Cmd.Parameters["現場名"].Value = genba.Text;

                        Cmd.Parameters["役職"].Value = yakusyoku.Text;
                        Cmd.Parameters["給与支給区分"].Value = kyuuyo.Text;
                        Cmd.Parameters["社員区分"].Value = "";
                        Cmd.Parameters["試用期間"].Value = shiyoukikan.Text;

                        Cmd.Parameters["契約社員"].Value = keiyaku.Text;
                        Cmd.Parameters["友の会区分"].Value = tomokubun.Text;
                        Cmd.Parameters["国籍"].Value = kokuseki.Text;

                        Cmd.Parameters["税表区分"].Value = zeikubun.Text;
                        Cmd.Parameters["本人障害"].Value = syougai.Text;
                        Cmd.Parameters["寡フ"].Value = kahu.Text;
                        Cmd.Parameters["勤労学生"].Value = gakusei.Text;
                        Cmd.Parameters["災害"].Value = saigai.Text;
                        Cmd.Parameters["外国人"].Value = gaikoku.Text;

                        Cmd.Parameters["休暇付与区分"].Value = kyuuka.Text;
                        Cmd.Parameters["基本勤務時間"].Value = kinmu.Text == "" ? 0 : Convert.ToDecimal(kinmu.Text);

                        if (hatsurei.Text != "1100　日給者契約" && hatsurei.Text != "0001　正社員採用")
                        {
                            Cmd.Parameters["時給"].Value = Convert.ToDecimal(zikyuu.Value);
                        }
                        else
                        {
                            Cmd.Parameters["時給"].Value = 0;
                        }

                        //Cmd.Parameters["時給"].Value = hatsurei.Text != "0001　正社員採用" ? Convert.ToDecimal(zikyuu.Value) : 0;
                        Cmd.Parameters["日給"].Value = hatsurei.Text == "1100　日給者契約" ? Convert.ToDecimal(nikkyuu.Value) : 0;
                        Cmd.Parameters["回数1"].Value = kaisuu1.Text == "" ? 0 : Convert.ToDecimal(kaisuu1.Text);
                        Cmd.Parameters["回数2"].Value = 0;

                        Cmd.Parameters["本給"].Value = honkyuu.Text == "" ? 0 : Convert.ToDecimal(honkyuu.Text);
                        Cmd.Parameters["職務技能給"].Value = syokumu.Text == "" ? 0 : Convert.ToDecimal(syokumu.Text);
                        Cmd.Parameters["特別手当"].Value = tokuteate.Text == "" ? 0 : Convert.ToDecimal(tokuteate.Text);
                        Cmd.Parameters["役職手当"].Value = yakuteate.Text == "" ? 0 : Convert.ToDecimal(yakuteate.Text);
                        Cmd.Parameters["免許手当"].Value = menkyoteate.Text == "" ? 0 : Convert.ToDecimal(menkyoteate.Text);
                        Cmd.Parameters["扶養手当"].Value = huyou.Text == "" ? 0 : Convert.ToDecimal(huyou.Text);
                        Cmd.Parameters["離島手当"].Value = ritou.Text == "" ? 0 : Convert.ToDecimal(ritou.Text);
                        Cmd.Parameters["転勤手当"].Value = tenkin.Text == "" ? 0 : Convert.ToDecimal(tenkin.Text);
                        Cmd.Parameters["通勤手当非"].Value = tuukinhi2.Text == "" ? 0 : Convert.ToDecimal(tuukinhi2.Text);
                        Cmd.Parameters["通勤手当課"].Value = tuukinka2.Text == "" ? 0 : Convert.ToDecimal(tuukinka2.Text);
                        Cmd.Parameters["登録手当"].Value = tourokuteate.Text == "" ? 0 : Convert.ToDecimal(tourokuteate.Text);
                        Cmd.Parameters["通信手当"].Value = 0;
                        Cmd.Parameters["友の会"].Value = tomonokai.Text == "" ? 0 : Convert.ToDecimal(tomonokai.Text);
                        Cmd.Parameters["固定1"].Value = 0;
                        Cmd.Parameters["固定2"].Value = 0;

                        Cmd.Parameters["特別理由"].Value = tokureason.Text;
                        Cmd.Parameters["赴任元先"].Value = comboBoxTenkin.Text;

                        if (hatsurei.Text != "1100　日給者契約" && hatsurei.Text != "0001　正社員採用")
                        {
                            Cmd.Parameters["職種"].Value = "";
                            Cmd.Parameters["最終学歴"].Value = "";
                            Cmd.Parameters["社外経験"].Value = "";
                        }
                        else
                        {
                            Cmd.Parameters["職種"].Value = comboBoxSyokusyu.Text;
                            Cmd.Parameters["最終学歴"].Value = comboBoxGaku.Text;
                            Cmd.Parameters["社外経験"].Value = comboBoxKeiken.Text;
                        }

                        //Cmd.Parameters["職種"].Value = hatsurei.Text == "0001　正社員採用" ? comboBoxSyoku.Text : "";
                        //Cmd.Parameters["最終学歴"].Value = hatsurei.Text == "0001　正社員採用" ? comboBoxGaku.Text : "";
                        //Cmd.Parameters["社外経験"].Value = hatsurei.Text == "0001　正社員採用" ? comboBoxKeiken.Text : "";

                        Cmd.Parameters["通勤手当区分"].Value = tuukinteatekubun.Text;
                        Cmd.Parameters["通勤手段区分"].Value = tuukinkubun.Text;
                        Cmd.Parameters["片道通勤距離"].Value = katakyori.Value;
                        Cmd.Parameters["片道料金"].Value = tuutanka.Text == "" ? 0 : Convert.ToDecimal(tuutanka.Text);

                        Cmd.Parameters["01"].Value = c01.Checked ? 1 : 0;
                        Cmd.Parameters["02"].Value = c02.Checked ? 1 : 0;
                        Cmd.Parameters["03"].Value = c03.Checked ? 1 : 0;
                        Cmd.Parameters["04"].Value = c04.Checked ? 1 : 0;
                        Cmd.Parameters["05"].Value = c05.Checked ? 1 : 0;
                        Cmd.Parameters["06"].Value = c06.Checked ? 1 : 0;
                        Cmd.Parameters["07"].Value = c07.Checked ? 1 : 0;
                        Cmd.Parameters["08"].Value = c08.Checked ? 1 : 0;
                        Cmd.Parameters["09"].Value = c09.Checked ? 1 : 0;
                        Cmd.Parameters["10"].Value = c10.Checked ? 1 : 0;
                        Cmd.Parameters["11"].Value = c11.Checked ? 1 : 0;
                        Cmd.Parameters["12"].Value = c12.Checked ? 1 : 0;
                        Cmd.Parameters["13"].Value = c13.Checked ? 1 : 0;
                        Cmd.Parameters["14"].Value = c14.Checked ? 1 : 0;
                        Cmd.Parameters["15"].Value = c15.Checked ? 1 : 0;
                        Cmd.Parameters["16"].Value = c16.Checked ? 1 : 0;
                        Cmd.Parameters["17"].Value = "";
                        Cmd.Parameters["18"].Value = c18.Checked ? 1 : 0;
                        Cmd.Parameters["19"].Value = c19.Checked ? 1 : 0;
                        Cmd.Parameters["20"].Value = c20.Checked ? 1 : 0;

                        Cmd.Parameters["入社きっかけ"].Value = kikkake.Text;
                        Cmd.Parameters["縁故社員番号"].Value = enko_no.Text;

                        Cmd.Parameters["保有資格有無"].Value = hoyuurbmu.Checked ? 0 : 1;
                        Cmd.Parameters["扶養家族有無"].Value = huyourbmu.Checked ? 0 : 1;

                        Cmd.Parameters["警告とエラー"].Value = error.Text;
                        Cmd.Parameters["削除理由"].Value = status.Text;
                        Cmd.Parameters["備考1"].Value = bikou1.Text;
                        Cmd.Parameters["備考2"].Value = bikou2.Text;
                        Cmd.Parameters["備考3"].Value = bikou3.Text;
                        Cmd.Parameters["備考4"].Value = 0; //固定控除1
                        Cmd.Parameters["備考5"].Value = 0; //固定控除2

                        Cmd.Parameters["備考6"].Value = kyuuzitsukubun.Text;


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

        private void hatsureitrig()
        {
            if (hatsurei.SelectedIndex == -1) return;

            string str = hatsurei.SelectedItem.ToString().Substring(0, 4);

            if (str == "0900") //パート契約
            {
                kyuuyo.Text = "E1　パート";
                //kyuuzitsukubun.Text = "03　パート";

                //時給・日給の設定
                zikyuu.Visible = true;
                zikyuu.Value = zikyuu.Minimum;

                nikkyuu.Visible = false;

                //職種学歴経験パネル
                seisyapanel.Visible = false;

                //試用期間
                shiyoukikan.Text = "";

                //勤務時間
                kinmu.Enabled = true;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                tomonokai.Text = "300";

                //契約社員
                keiyaku.SelectedIndex = -1;
                keiyaku.Visible = false;

            }
            else if (str == "1000") //アルバイト契約
            {
                //社員区分設定
                kyuuyo.Text = "F1　アルバイト";
                //kyuuzitsukubun.Text = "04　アルバイト";

                //時給・日給の設定
                zikyuu.Visible = true;
                nikkyuu.Visible = false;
                zikyuu.Value = zikyuu.Minimum;

                //職種学歴経験パネル
                seisyapanel.Visible = false;

                //試用期間
                shiyoukikan.Text = "";

                //勤務時間
                kinmu.Enabled = true;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("2　アルバイト加入");

                tomonokai.Text = "0";

                //契約社員
                keiyaku.SelectedIndex = -1;
                keiyaku.Visible = false;

            }
            else if (str == "1100") //日給者契約
            {
                //社員区分設定
                kyuuyo.Text = "D1　日給者";
                //kyuuzitsukubun.Text = "02　日給者";
                //時給・日給の設定
                zikyuu.Visible = false;
                nikkyuu.Visible = true;

                if (nikkyuu.Value == 0)
                {
                    nikkyuu.Value = nikkyuu.Minimum;
                }

                //職種学歴経験パネル
                seisyapanel.Visible = true;

                //金額は非表示
                syokumukyuu.Visible = false;
                gakurekikyuu.Visible = false;
                keikenkyuu.Visible = false;
                nennreikyuu.Visible = false;

                //試用期間
                //shiyoukikan.Text = "11 有";

                //勤務時間
                //kinmu.SelectedIndex = 0;
                //kinmu.Enabled = false;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                tomonokai.Text = "300";

                //契約社員
                keiyaku.Visible = true;
            }
            else if (str == "0001") //正社員採用
            {
                //社員区分設定
                kyuuyo.Text = "C1　月給者";
                //kyuuzitsukubun.Text = "01　月給者";
                //時給・日給の設定
                zikyuu.Visible = false;
                nikkyuu.Visible = false;
                //nikkyuu.Value = nikkyuu.Minimum;

                //職種学歴経験パネル
                seisyapanel.Visible = true;

                //金額は非表示
                syokumukyuu.Visible = true;
                gakurekikyuu.Visible = true;
                keikenkyuu.Visible = true;
                nennreikyuu.Visible = true;

                //試用期間
                shiyoukikan.Text = "11 有";

                //勤務時間
                kinmu.SelectedIndex = 0;
                kinmu.Enabled = false;

                //友の会区分
                tomokubun.Items.Clear();
                tomokubun.Items.Add("");
                tomokubun.Items.Add("1　非加入");

                tomonokai.Text = "300";

                //契約社員
                keiyaku.Visible = true;
            }
            else
            {
                MessageBox.Show("Error:社員区分で想定外　システム管理者へ連絡願います。");
            }
        }

        private void hatsurei_SelectedIndexChanged(object sender, EventArgs e)
        {
            hatsureitrig();

            ErrorCheck();

            
            //TODO 202203 いまいちわからん
            if (hatsurei.SelectedItem?.ToString() == "0001　正社員採用") return;

            //TODO 離島と出向の対応
            if (tenkin.Text == "")
            {

            }
            else if (Convert.ToDecimal(tenkin.Text) > 0)
            {
                ritou.Text = "0";
                return;
            }
            if (soshiki.SelectedItem == null || genba.SelectedItem == null || comboBoxSyokusyu.SelectedItem == null) return;
            if (kyuuyo.Text.Substring(0, 2) != "C1") return;
            ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //ヘッダは対象外
            if (dataGridView1.CurrentCell != null)
            {
                if (button2.Text == "更新" && hatsurei.SelectedIndex != hatsurei_ex)
                {

                    DialogResult result = MessageBox.Show("○○の変更内容を保存しますか？",
                                                "質問",
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Exclamation,
                                                MessageBoxDefaultButton.Button2);

                    //何が選択されたか調べる
                    if (result == DialogResult.Yes)
                    {
                        //「はい」が選択された時
                        // TODO 何処理？
                        Console.WriteLine("更新処理して選択変更");
                    }
                    else if (result == DialogResult.No)
                    {
                        //「いいえ」が選択された時
                        // TODO 何処理？
                        Console.WriteLine("選択変更のみ");
                    }

                }

                AllClear();

                //発令区分による項目設定
                hatsureitrig();

                DataGridViewRow dgr = dataGridView1.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                DataDisp(drv[1].ToString());

                //資格情報一覧
                shikakudgv.CurrentCell = null;
                //家族情報一覧
                kazokudgv.CurrentCell = null;

                //エラーチェック
                ErrorCheck();
            }


        }

        private void DataDisp(string str)
        {

            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.n入社データ where 社員番号 = '" + str + "'");

            //DataTable dt2 = new DataTable();
            //dt2 = Com.GetDB("select * from dbo.通勤管理テーブル where 社員番号 = '" + str + "' and 管理No = '1'");

            foreach (DataRow row in dt.Rows)
            {
                syainno.Text = row["社員番号"].ToString();
                saiyoudate.Text = row["採用年月日"].ToString();
                hatsurei.SelectedIndex = hatsurei_ex = hatsurei.FindString(row["発令区分"].ToString());

                seihuri.Text = kazokanasei.Text = row["姓フリ"].ToString();
                meihuri.Text = row["名フリ"].ToString();
                sei.Text = kazosei.Text = row["姓"].ToString();
                mei.Text = row["名"].ToString();

                if (row["生年月日"].ToString() == "          ")
                {
                    birthnew.Value = null;
                }
                else
                {
                    birthnew.Value = Convert.ToDateTime(row["生年月日"]);
                }

                seibetsu.SelectedIndex = seibetsu.FindString(row["性別"].ToString());

                yakusyoku.Text = row["役職"].ToString() == "0180" ? "0180　係員" : "";

                if (row["給与支給区分"].ToString() == "E1")
                {
                    kyuuyo.Text = "E1　パート";
                }
                else if (row["給与支給区分"].ToString() == "F1")
                {
                    kyuuyo.Text = "F1　アルバイト";
                }
                else if (row["給与支給区分"].ToString() == "D1")
                {
                    kyuuyo.Text = "D1　日給者";
                }
                else if (row["給与支給区分"].ToString() == "C1")
                {
                    kyuuyo.Text = "C1　月給者";
                }
                else
                {
                    kyuuyo.Text = "";
                }

                //if (row["社員区分"].ToString() == "03")
                //{
                //    kyuuzitsukubun.Text = "03　パート";
                //}
                //else if (row["社員区分"].ToString() == "04")
                //{
                //    kyuuzitsukubun.Text = "04　アルバイト";
                //}
                //else if (row["社員区分"].ToString() == "02")
                //{
                //    kyuuzitsukubun.Text = "02　日給者";
                //}
                //else if (row["社員区分"].ToString() == "01")
                //{
                //    kyuuzitsukubun.Text = "01　月給者";
                //}
                //else
                //{
                //    kyuuzitsukubun.Text = "";
                //}

                //if (row["備考6"].ToString() == "10")
                //{
                //    kyuuzitsukubun.Text = "10　年間最低数";
                //    kyuuzitsukubun.SelectedIndex = kyuuzitsukubun.FindString(row["備考6"].ToString());
                //}
                //else if (row["備考6"].ToString() == "20")
                //{
                //    kyuuzitsukubun.Text = "20　土日祝";
                //}

                //休日区分
                kyuuzitsukubun.SelectedIndex = kyuuzitsukubun.FindString(row["備考6"].ToString());


                shiyoukikan.Text = row["試用期間"].ToString() == "11" ? "11 有" : "";

                keiyaku.SelectedIndex = keiyaku.FindString(row["契約社員"].ToString());
                tomokubun.SelectedIndex = tomokubun.FindString(row["友の会区分"].ToString());
                kokuseki.SelectedIndex = kokuseki.FindString(row["国籍"].ToString());

                tiku.SelectedIndex = tiku.FindString(row["地区"].ToString());
                soshiki.SelectedIndex = soshiki.FindString(row["所属組織"].ToString());
                genba.SelectedIndex = genba.FindString(row["現場名"].ToString());

                yuubin.Text = row["現郵便番号"].ToString();
                zyuusyo.Text = row["現住所"].ToString();
                zyuusyocheck.Checked = row["住所フラグ"].ToString() == "1" ? true : false;
                keitai.Text = row["携帯番号"].ToString();

                zeikubun.SelectedIndex = zeikubun.FindString(row["税表区分"].ToString());
                syougai.SelectedIndex = syougai.FindString(row["本人障害"].ToString());
                kahu.SelectedIndex = kahu.FindString(row["寡フ"].ToString());
                gakusei.SelectedIndex = gakusei.FindString(row["勤労学生"].ToString());
                saigai.SelectedIndex = saigai.FindString(row["災害"].ToString());
                gaikoku.SelectedIndex = gaikoku.FindString(row["外国人"].ToString());

                kyuuka.SelectedIndex = kyuuka.FindString(row["休暇付与区分"].ToString());
                kinmu.SelectedIndex = kinmu.FindString(row["基本勤務時間"].ToString());
                zikyuu.Text = row["時給"].ToString();
                nikkyuu.Text = row["日給"].ToString();
                kaisuu1.Text = row["回数1"].ToString();
                //kaisuu2.Text = row["回数2"].ToString();

                honkyuu.Text = row["本給"].ToString();
                syokumu.Text = row["職務技能給"].ToString();
                tokuteate.Text = row["特別手当"].ToString();
                yakuteate.Text = row["役職手当"].ToString();
                menkyoteate.Text = row["免許手当"].ToString();
                huyou.Text = row["扶養手当"].ToString();
                ritou.Text = row["離島手当"].ToString();

                tenkin.Text = row["転勤手当"].ToString();
                tuukinhi2.Text = row["通勤手当(非)"].ToString();
                tuukinka2.Text = row["通勤手当(課)"].ToString();
                tourokuteate.Text = row["登録手当"].ToString();
                tuushinteate.Text = row["通信手当"].ToString();

                tomonokai.Text = row["友の会"].ToString();

                tokureason.Text = row["特別理由"].ToString();

                comboBoxTenkin.SelectedIndex = comboBoxTenkin.FindString(row["赴任元先"].ToString());
                
                comboBoxSyokusyu.SelectedIndex = comboBoxSyokusyu.FindString(row["職種"].ToString());
                comboBoxGaku.SelectedIndex = comboBoxGaku.FindString(row["最終学歴"].ToString());
                comboBoxKeiken.SelectedIndex = comboBoxKeiken.FindString(row["社外経験"].ToString());

                c01.Checked = row["01"].ToString() == "1" ? true : false;
                c02.Checked = row["02"].ToString() == "1" ? true : false;
                c03.Checked = row["03"].ToString() == "1" ? true : false;
                c04.Checked = row["04"].ToString() == "1" ? true : false;
                c05.Checked = row["05"].ToString() == "1" ? true : false;
                c06.Checked = row["06"].ToString() == "1" ? true : false;
                c07.Checked = row["07"].ToString() == "1" ? true : false;
                c08.Checked = row["08"].ToString() == "1" ? true : false;
                c09.Checked = row["09"].ToString() == "1" ? true : false;
                c10.Checked = row["10"].ToString() == "1" ? true : false;
                c11.Checked = row["11"].ToString() == "1" ? true : false;
                c12.Checked = row["12"].ToString() == "1" ? true : false;
                c13.Checked = row["13"].ToString() == "1" ? true : false;
                c14.Checked = row["14"].ToString() == "1" ? true : false;
                c15.Checked = row["15"].ToString() == "1" ? true : false;
                c16.Checked = row["16"].ToString() == "1" ? true : false;
                //c17.Checked = row["17"].ToString() == "1" ? true : false;
                c18.Checked = row["18"].ToString() == "1" ? true : false;
                c19.Checked = row["19"].ToString() == "1" ? true : false;
                c20.Checked = row["20"].ToString() == "1" ? true : false;

                hoyuurbmu.Checked = row["保有資格有無"].ToString() == "0" ? true : false;
                huyourbmu.Checked = row["扶養家族有無"].ToString() == "0" ? true : false;

                hoyuurbyuu.Checked = row["保有資格有無"].ToString() == "1" ? true : false;
                huyourbyuu.Checked = row["扶養家族有無"].ToString() == "1" ? true : false;

                kikkake.SelectedIndex = kikkake.FindString(row["入社きっかけ"].ToString());
                enko_no.Text = row["縁故社員番号"].ToString();

                if (enko_no.Text != "        ")
                {
                    //縁故紹介社員番号から情報取得
                    DataTable enkodt = new DataTable();
                    enkodt = Com.GetDB("select 氏名, 組織名, 現場名 from dbo.社員基本情報 where 社員番号 = '" + enko_no.Text + "'");

                    enko_name.Text = enkodt.Rows[0][0].ToString();
                    enko_soshiki.Text = enkodt.Rows[0][1].ToString();
                    enko_genba.Text = enkodt.Rows[0][2].ToString();
                }

                error.Text = row["警告とエラー"].ToString();
                status.SelectedIndex = status.FindString(row["削除理由"].ToString());
                bikou1.Text = row["備考1"].ToString();
                bikou2.Text = row["備考2"].ToString();
                bikou3.Text = row["備考3"].ToString();
            }

            //自家用車 通勤管理情報取得
            GetMyCar(str);


            //本給値取得
            GetHonkyuu();

            //資格情報取得
            DataDispShikaku();
            //家族情報取得
            DataDispKazoku();
            //労働条件情報取得
            DataDispRoudou();

            //緊急連絡先
            DataDispKinkyuu(str);

            Calc();
        }

        private void DataDispKinkyuu(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.k緊急連絡先  where 社員番号 = '" + str + "'");

            if (dt.Rows.Count > 0)
            {
                honkeitai.Text = dt.Rows[0]["携帯番号"].ToString();
                honkotei.Text = dt.Rows[0]["固定電話"].ToString();

                kaz1name.Text = dt.Rows[0]["氏名1"].ToString();
                kaz1kana.Text = dt.Rows[0]["カナ名1"].ToString();
                kaz1gara.Text = dt.Rows[0]["続柄1"].ToString();
                kaz1no.Text = dt.Rows[0]["電話番号1"].ToString();

                kaz2name.Text = dt.Rows[0]["氏名2"].ToString();
                kaz2kana.Text = dt.Rows[0]["カナ名2"].ToString();
                kaz2gara.Text = dt.Rows[0]["続柄2"].ToString();
                kaz2no.Text = dt.Rows[0]["電話番号2"].ToString();
            }
        }


        private void GetMyCar(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.t通勤管理テーブル a left join dbo.t通勤手当元データ b on a.社員番号 = b.社員番号 where a.社員番号 = '" + str + "' and a.管理No = '1'");

            if (dt.Rows.Count == 0) return;

            syainno.Text = dt.Rows[0]["社員番号"].ToString();

            tuukinteatekubun.SelectedIndex = tuukinteatekubun.FindString(dt.Rows[0]["通勤手当区分"].ToString());
            tuukinkubun.SelectedIndex = tuukinkubun.FindString(dt.Rows[0]["通勤方法"].ToString());

            if (dt.Rows[0]["免許証"].ToString() == "")
            {
                menkyonew.Value = null;
            }
            else
            {
                menkyonew.Value = Convert.ToDateTime(dt.Rows[0]["免許証"].ToString());
            }

            if (dt.Rows[0]["車検証"].ToString() == "")
            {
                syakennew.Value = null;
            }
            else
            {
                syakennew.Value = Convert.ToDateTime(dt.Rows[0]["車検証"].ToString());
            }

            if (dt.Rows[0]["自賠責"].ToString() == "")
            {
                zibainew.Value = null;
            }
            else
            {
                zibainew.Value = Convert.ToDateTime(dt.Rows[0]["自賠責"].ToString());
            }

            if (dt.Rows[0]["任意保険"].ToString() == "")
            {
                ninninew.Value = null;
            }
            else
            {
                ninninew.Value = Convert.ToDateTime(dt.Rows[0]["任意保険"].ToString());
            }

            //if (dt.Rows[0][4].ToString() != "") syaken.Value = Convert.ToDateTime(dt.Rows[0][4].ToString());
            //if (dt.Rows[0][5].ToString() != "") zibai.Value = Convert.ToDateTime(dt.Rows[0][5].ToString());
            //if (dt.Rows[0][6].ToString() != "") ninni.Value = Convert.ToDateTime(dt.Rows[0][6].ToString());
            tuukinteatekubun.SelectedIndex = tuukinteatekubun.FindString(dt.Rows[0]["通勤手当区分"].ToString());


            katakyori.Text = dt.Rows[0]["片道距離"].ToString();
            tuutanka.Text = dt.Rows[0]["通勤1日単価"].ToString();
            //kataryoukin.Text = dt.Rows[0]["片道料金"].ToString();

            //11メーカー
            //12車名
            //13色
            //14車両番号
            //15備考

            //TODO 備考のコントロールを追加
            //bikou.Text = dt.Rows[0][15].ToString();

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

                zikangairoudou.SelectedIndex = koushinkubun.FindString(roudoudt.Rows[0]["時間外労働区分"].ToString());
                yakankinmu.SelectedIndex = yakankinmu.FindString(roudoudt.Rows[0]["夜間勤務区分"].ToString());

                kinmuH.Text = "【" + kinmu.SelectedItem.ToString() + "時間】";

                //TODO 週労働数との連動処理
                syuuroucopy.Text = kyuuka.SelectedItem.ToString();

                switch (kyuuka.SelectedItem.ToString())
                {
                    case "0　５日以上":

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
                    case "1　４日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("3");
                        break;
                    case "2　３日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("4");
                        break;
                    case "3　２日":
                        kyuujitsukaisuu.SelectedIndex = kyuujitsukaisuu.FindString("5");
                        break;
                    case "4　１日":
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

                //TODO 途中で追加。　厚生年金・健康保険・雇用保険
                kouseicb.SelectedIndex = kouseicb.FindString(roudoudt.Rows[0]["入社時_厚生年金"].ToString());
                kenkoucb.SelectedIndex = kenkoucb.FindString(roudoudt.Rows[0]["入社時_健康保険"].ToString());
                koyoucb.SelectedIndex = koyoucb.FindString(roudoudt.Rows[0]["入社時_雇用保険"].ToString());
            }


            //選択無だった場合に初期データとして

            //期末日設定
            //DateTime s49ki = new System.DateTime(2020, 04, 01, 0, 0, 0, 0);
            DateTime s50ki = new System.DateTime(2021, 04, 01, 0, 0, 0, 0);
            DateTime s51ki = new System.DateTime(2022, 04, 01, 0, 0, 0, 0);
            DateTime s52ki = new System.DateTime(2023, 04, 01, 0, 0, 0, 0);
            DateTime s53ki = new System.DateTime(2024, 04, 01, 0, 0, 0, 0);
            DateTime s54ki = new System.DateTime(2025, 04, 01, 0, 0, 0, 0);
            DateTime s55ki = new System.DateTime(2026, 04, 01, 0, 0, 0, 0);

            DateTime kimatsuday = new System.DateTime(2021, 03, 31, 0, 0, 0, 0);
            if (saiyoudate.Value < s50ki)
            {
                //DateTime kimatsuday = new System.DateTime(2020, 03, 31, 0, 0, 0, 0);
            }
            else if (saiyoudate.Value < s51ki)
            {
                kimatsuday = new System.DateTime(2022, 03, 31, 0, 0, 0, 0);
            }
            else if (saiyoudate.Value < s52ki)
            {
                 kimatsuday = new System.DateTime(2023, 03, 31, 0, 0, 0, 0);
            }
            else if (saiyoudate.Value < s53ki)
            {
                kimatsuday = new System.DateTime(2024, 03, 31, 0, 0, 0, 0);
            }
            else if (saiyoudate.Value < s54ki)
            {
                kimatsuday = new System.DateTime(2025, 03, 31, 0, 0, 0, 0);
            }
            else if (saiyoudate.Value < s55ki)
            {
                kimatsuday = new System.DateTime(2026, 03, 31, 0, 0, 0, 0);
            }

            //共通設定
            if (keiyakunengetsu.Value.Equals(DBNull.Value)) keiyakunengetsu.Value = saiyoudate.Value; //作成日　なし

            if (kyuuyo.Text.Substring(0, 2) == "F1")
            {
                //アルバイトの場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = saiyoudate.Value.AddMonths(6); //雇用区分 入社日から6ヶ月後
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 3; //更新区分　契約の更新はしない
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 2; //休日勤務　なし
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり

                if (kouseicb.SelectedIndex == 0) kouseicb.SelectedIndex = 2;
                if (kenkoucb.SelectedIndex == 0) kenkoucb.SelectedIndex = 2;
                if (koyoucb.SelectedIndex == 0) koyoucb.SelectedIndex = 2;
            }
            else if (kyuuyo.Text.Substring(0, 2) == "E1")
            {
                //パートの場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = kimatsuday; //雇用区分 期末日
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 2; //更新区分　更新する場合がありえる
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 2; //休日勤務　なし
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり

                if (kouseicb.SelectedIndex == 0) kouseicb.SelectedIndex = 2;
                if (kenkoucb.SelectedIndex == 0) kenkoucb.SelectedIndex = 2;
                if (koyoucb.SelectedIndex == 0) koyoucb.SelectedIndex = 2;

            }
            else if (kyuuyo.Text.Substring(0, 2) == "D1")
            {
                //日給者の場合
                if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 2; //雇用区分 定めなし
                if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 1; //更新区分　自動更新
                if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 1; //賞与　あり
                if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり

                if (kouseicb.SelectedIndex == 0) kouseicb.SelectedIndex = 1;
                if (kenkoucb.SelectedIndex == 0) kenkoucb.SelectedIndex = 1;
                if (koyoucb.SelectedIndex == 0) koyoucb.SelectedIndex = 1;
            }
            else if (kyuuyo.Text.Substring(0, 2) == "C1")
            {

                //月給者　共通設定
                if (kouseicb.SelectedIndex == 0) kouseicb.SelectedIndex = 1;
                if (kenkoucb.SelectedIndex == 0) kenkoucb.SelectedIndex = 1;
                if (koyoucb.SelectedIndex == 0) koyoucb.SelectedIndex = 1;

                if (shiyoukikan.Text == "11 有")
                {
                    //試用期間の場合

                    if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                    if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                    if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = kimatsuday; //雇用区分 期末日
                    if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 2; //更新区分　更新する場合がありえる
                    if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                    if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                    if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 2; //賞与　なし
                    if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                    if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり

                }
                else if (keiyaku.Text != "")
                {
                    //TODO 入社で契約社員は現在の規程ではありえない。

                    //契約社員の場合
                    //keiyaku.Items.Add("10" + "　" + "一般契約社員");
                    //keiyaku.Items.Add("20" + "　" + "単年契約社員");

                    if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 1; //雇用区分 定めあり
                    if (koyoukaishibi.Value.Equals(DBNull.Value)) koyoukaishibi.Value = saiyoudate.Value; //雇用区分 入社年月日
                    if (koyousyuuryoubi.Value.Equals(DBNull.Value)) koyousyuuryoubi.Value = kimatsuday; //雇用区分 期末日
                    if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 2; //更新区分　更新する場合がありえる
                    if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                    if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 1; //定年　なし
                    if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 1; //賞与　あり
                    if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 2; //退職金　なし
                    if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり
                }
                else
                {
                    //管理職の場合
                    //TODO とりあえず設定無


                    //通常
                    if (koyoukubun.SelectedIndex == 0) koyoukubun.SelectedIndex = 2; //雇用区分 定めなし
                    if (koushinkubun.SelectedIndex == 0) koushinkubun.SelectedIndex = 1; //更新区分　自動更新
                    if (kyuusyustu.SelectedIndex == 0) kyuusyustu.SelectedIndex = 1; //休日勤務　あり
                    if (teinen.SelectedIndex == 0) teinen.SelectedIndex = 2; //定年　あり
                    if (syouyo.SelectedIndex == 0) syouyo.SelectedIndex = 1; //賞与　あり
                    if (taisyokukin.SelectedIndex == 0) taisyokukin.SelectedIndex = 1; //退職金　あり
                    if (zikangairoudou.SelectedIndex == 0) zikangairoudou.SelectedIndex = 1; //時間外労働　あり

 
                }


            }
            else
            {
                //兼務役員の場合
                //役員の場合

            }





        }



        private void DataDispShikaku()
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "'");
            shikakudgv.DataSource = dt;

            if (dt.Rows.Count == 0) return;
            if (shikakudgv.RowCount == 0) return;

            //TODO カラム幅
            shikakudgv.Columns["資格コード"].Width = 80;
            shikakudgv.Columns["資格名"].Width = 200;
            shikakudgv.Columns["資格取得日"].Width = 100;
            shikakudgv.Columns["資格取得番号"].Width = 150;
            shikakudgv.Columns["資格有効期限"].Width = 110;
            shikakudgv.Columns["規程額"].Width = 60;
            shikakudgv.Columns["適用終了日"].Width = 100;
            shikakudgv.Columns["期限"].Width = 60;

            shikakudgv.Columns["社員番号"].Visible = false;
            shikakudgv.Columns["個人識別ＩＤ"].Visible = false;
            shikakudgv.Columns["規程額"].DefaultCellStyle.Format = "#,0";
            shikakudgv.Columns["規程額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //免許手当総額を取得
            DataTable dt2 = new DataTable();
            dt2 = Com.GetDB("select * from dbo.m免許手当合計規程額取得_入社 where 社員番号 = '" + syainno.Text + "'");

            if (dt2.Rows.Count > 0)
            {
                if (dt2.Rows[0][0].Equals(DBNull.Value))
                {
                    
                }
                else
                {
                    menkyogaku.Text = Convert.ToDecimal(dt2.Rows[0][0]).ToString("#,0");
                }
            }


            //登録手当総額を取得
            DataTable dt3 = new DataTable();
            dt3 = Com.GetDB("select * from dbo.t登録手当規程額_入社 where 社員番号 = '" + syainno.Text + "'");

            if (dt3.Rows.Count > 0)
            {
                if (dt3.Rows[0][1].Equals(DBNull.Value))
                {

                }
                else
                {
                    tourokugaku.Text = Convert.ToDecimal(dt3.Rows[0][1]).ToString("#,0");
                }
            }

        }

        private void GetShikakukoumokuData(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "' and 資格コード = '" + str + "'");

            //shikakucombo.SelectedIndex = shikakucombo.FindString(dt.Rows[0][0].ToString());
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
            }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                syainno.Text = "";
            }
            else
            {
                AllClear();

                dataGridView1.CurrentCell = null;
            }
        }


        private void huri_Validating(object sender, CancelEventArgs e)
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


        private void name_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //スペースは全角に
            string result = ((TextBox)sender).Text.Replace(" ", "　").Trim();

            ((TextBox)sender).Text = result;
        }

        private void yuubin_Validating(object sender, CancelEventArgs e)
        {
            //初回エラークリア
            this.errorProvider1.SetError((TextBox)sender, "");

            if (((TextBox)sender).Text == "") return;

            //全角文字を半角文字に変換
            string result = Microsoft.VisualBasic.Strings.StrConv(((TextBox)sender).Text, VbStrConv.Narrow, 0).Trim();


            //郵便番号以外除去 
            Regex regex = new Regex(@"^(\d{7})|(\d{3}-\d{4})$");
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError((TextBox)sender, "郵便番号が正しくありません。");
                e.Cancel = true;
                return;
            }

            //((TextBox)sender).Text = GetAdress(result);
        }


        private void keitai_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "電話番号", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void ValidCom(Control cont, CancelEventArgs e, string str, string reg)
        {
            //初回エラークリア
            this.errorProvider1.SetError(cont, "");

            if (cont.Text == "") return;

            //全角文字を半角文字に変換
            string result = Microsoft.VisualBasic.Strings.StrConv(cont.Text, VbStrConv.Narrow, 0).Trim();

            cont.Text = result;

            //電話番号以外除去 
            Regex regex = new Regex(reg);
            if (!regex.IsMatch(result))
            {
                this.errorProvider1.SetError(cont, str + "が正しくありません。２つの-(ハイフン)は必須です。");
                e.Cancel = true;
                return;
            }

            //ハイフン除去
            //cont.Text = result.Replace("-", "");
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
        private void zikyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void nikkyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void kaisuu1_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void kaisuu2_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void honkyuu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void tyousei_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void tokubetsu_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void genteate_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void syukugenba_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void syukkou_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void tuushin_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void hoiku_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void tomonokai_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void kotei1_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void kotei2_Validating(object sender, CancelEventArgs e)
        {
            validCommon(((TextBox)sender), e);
        }

        private void birth_ValueChanged(object sender, EventArgs e)
        {


        }

        private void GetDataKihonkyuu()
        {
            syokumukyuu.Text = "0";
            nennreikyuu.Text = "0";
            keikenkyuu.Text = "0";
            gakurekikyuu.Text = "0";

            Syoku = Com.GetDB("select * from dbo.K_職務給_職種");

            foreach (DataRow row in Syoku.Rows)
            {
                comboBoxSyokusyu.Items.Add(row["備考"]);
            }

            Nen = Com.GetDB("select * from dbo.K_技能給_A年齢");

            Keiken = Com.GetDB("select * from dbo.K_技能給_B社外経験");
            foreach (DataRow row in Keiken.Rows)
            {
                comboBoxKeiken.Items.Add(row["備考"]);
            }


            Gaku = Com.GetDB("select * from dbo.K_技能給_C最終学歴");

            foreach (DataRow row in Gaku.Rows)
            {
                comboBoxGaku.Items.Add(row["備考"]);
            }

            Calc();
        }

        private void Calc()
        {
            //職種
            if (comboBoxSyokusyu.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Syoku.Select("備考 = '" + comboBoxSyokusyu.SelectedItem.ToString() + "'");
                syokumukyuu.Text = dr[0][1].ToString();
            }

            if (Nen.Rows.Count == 0) return;

            //年齢
            //int old = CalcAge(hatsurei.Value, dateTimePicker2.Value);
            //nenrei.Text = old.ToString();


            if (birthnew.Text != "")
            {
                DataRow[] dro = Nen.Select("年齢 = '" + old.Text + "'");
                nennreikyuu.Text = dro[0][1].ToString();
            }

            //学歴
            if (comboBoxGaku.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Gaku.Select("備考 = '" + comboBoxGaku.SelectedItem.ToString() + "'");
                gakurekikyuu.Text = dr[0][1].ToString();
            }

            //社外経験
            if (comboBoxKeiken.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Keiken.Select("備考 = '" + comboBoxKeiken.SelectedItem.ToString() + "'");
                keikenkyuu.Text = dr[0][1].ToString();
            }


            //合計額
            if (hatsurei.SelectedItem?.ToString() == "0001　正社員採用")
            {

                syokumu.Text = (Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) + Convert.ToDecimal(gakurekikyuu.Text)).ToString("#,0");

                //if (keiyaku.SelectedItem?.ToString() == "30　技・特実習生")
                //{
                //    //技能実習生はスルー
                //}
                //else
                //{ 
                //    syokumu.Text = (Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) + Convert.ToDecimal(gakurekikyuu.Text)).ToString("#,0");
                //}
            }
            else
            {
                syokumu.Text = "0";
            }

            kizyunnnai.Text = (Convert.ToDecimal(honkyuu.Text) + Convert.ToDecimal(syokumu.Text) + Convert.ToDecimal(tokuteate.Text) + Convert.ToDecimal(yakuteate.Text) + Convert.ToDecimal(menkyoteate.Text) + Convert.ToDecimal(tourokuteate.Text) + Convert.ToDecimal(tuushinteate.Text) + Convert.ToDecimal(tenkin.Text) + Convert.ToDecimal(ritou.Text)).ToString("#,0");
            shikyuu.Text = (Convert.ToDecimal(kizyunnnai.Text) + Convert.ToDecimal(huyou.Text) + Convert.ToDecimal(tuukinhi.Text) + Convert.ToDecimal(tuukinka.Text)).ToString("#,0");

            //日給換算
            //gaisan.Text = (Math.Round((Convert.ToDecimal(syokumu.Text) + Convert.ToDecimal(honkyuu.Text)) / Convert.ToDecimal(21.5))).ToString("#,0");
            //gaisanall.Text = (Math.Round((Convert.ToDecimal(kizyunnnai.Text)) / Convert.ToDecimal(21.5))).ToString("#,0");

        }

        public int CalcAge(DateTime birthday, DateTime NyusyaDay)
        {
            int i = 0;
            i = (int.Parse(NyusyaDay.ToString("yyyyMMdd")) - int.Parse(birthday.ToString("yyyyMMdd"))) / 10000;
            if (i < 0) i = 0;
            return i;
        }

        private void comboBoxSyoku_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();

            if (tenkin.Text == "")
            {

            }
            else if (Convert.ToDecimal(tenkin.Text) > 0)
            {
                ritou.Text = "0";
                return;
            }
            if (soshiki.SelectedItem == null || genba.SelectedItem == null || comboBoxSyokusyu.SelectedItem == null) return;
            if (kyuuyo.Text.Substring(0, 2) != "C1") return;
            //ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), genba.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
            ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
        }

        private void comboBoxGaku_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void comboBoxKeiken_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void old_TextChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void label30_Click(object sender, EventArgs e)
        {

        }



        private void maintab_Click(object sender, EventArgs e)
        {

        }

        private void zikyuu_ValueChanged(object sender, EventArgs e)
        {
            GetHonkyuu();
        }

        private void nikkyuu_ValueChanged(object sender, EventArgs e)
        {
            GetHonkyuu();
        }

        private void GetHonkyuu()
        {
            if (hatsurei.Text == "")
            {

            }
            else if (hatsurei.Text == "0001　正社員採用")
            {
                //if (keiyaku.SelectedItem?.ToString() == "30　技・特実習生")
                //{
                //    //技能実習生は変更可能に。
                //    honkyuu.Enabled = true;
                //    syokumu.Enabled = true;

                //    honkyuu.ReadOnly = false;
                //    syokumu.ReadOnly = false;
                //}
                //else
                //{
                    //TODO 
                    DataTable hkdt = new DataTable();
                    hkdt = Com.GetDB("select 本給 from dbo.HK_本給 where '" + Convert.ToDateTime(saiyoudate.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

                    //honkyuu.Value = 142000;
                    honkyuu.Text = hkdt.Rows[0][0].ToString();

                    honkyuu.Enabled = false;
                    syokumu.Enabled = false;

                    honkyuu.ReadOnly = true;
                    syokumu.ReadOnly = true;
                //}
            }
            else 
            {
                if (kinmu.Text != "")
                {
                    if (hatsurei.Text == "1100　日給者契約")
                    {
                        //日給
                        honkyuu.Text = Convert.ToString(Convert.ToInt32(nikkyuu.Value) * Getrday(kyuuka.Text));
                    }
                    else
                    {
                        //時給
                        honkyuu.Text = Convert.ToString(Convert.ToInt32(zikyuu.Value) * Convert.ToInt32(kinmu.Text) * Getrday(kyuuka.Text));
                    }
                }
            }
        }

        private double Getrday(string syuur)
        {
            double rday = 0;
            if (syuur == "0　５日以上")
            {
                rday = 21.5;
            }
            else if (syuur == "1　４日")
            {
                rday = 17.5;
            }
            else if (syuur == "2　３日")
            {
                rday = 13.5;
            }
            else if (syuur == "3　２日")
            {
                rday = 9;
            }
            else if (syuur == "4　１日")
            {
                rday = 4.5;
            }
            else if (syuur == "9　付与なし")
            {
                rday = 21.5;
            }
            return rday;
        }

        private void tuukinkubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            //TODO:　全リセットはしなくていいパターンによって変化させるべき
            //kataryoukin.Value = 0;
            //katakyori.Value = 0;

            tuukinCalc();

            if (tuukinkubun.SelectedIndex == 0)
            {
                //車
                c13.Enabled = true;
                c14.Enabled = true;
                c15.Enabled = true;

                mycarkigenpanel.Visible = true;

                menkyonew.Visible = true;
                //menkyonew.Value = null;

                syakennew.Visible = true;
                //syakennew.Value = null;

                zibainew.Visible = false;
                zibainew.Value = null;

                ninninew.Visible = true;
                //ninninew.Value = null;
            }
            else if (tuukinkubun.SelectedIndex == 1)
            {
                //バイク
                c13.Enabled = true;
                c14.Enabled = true;
                c15.Enabled = true;

                mycarkigenpanel.Visible = true;

                menkyonew.Visible = true;
                //menkyonew.Value = null;

                syakennew.Visible = false;
                syakennew.Value = null;

                zibainew.Visible = true;
                //zibainew.Value = null;

                ninninew.Visible = true;
                //ninninew.Value = null;
            }
            else if (tuukinkubun.SelectedIndex == 2)
            {
                //バス・モノレール
                c13.Enabled = false;
                c14.Enabled = false;
                c15.Enabled = false;

                mycarkigenpanel.Visible = false;
            }
            else if (tuukinkubun.SelectedIndex == 3)
            {
                //送迎(会社)
                c13.Enabled = false;
                c14.Enabled = false;
                c15.Enabled = false;

                mycarkigenpanel.Visible = false;


            }
            else if (tuukinkubun.SelectedIndex == 4
                )
            {
                //送迎(知人・親族)
                c13.Enabled = false;
                c14.Enabled = false;
                c15.Enabled = false;

                mycarkigenpanel.Visible = false;

            }
            else if (tuukinkubun.SelectedIndex == 5
                )
            {
                //業務車両
                c13.Enabled = false;
                c14.Enabled = false;
                c15.Enabled = false;

                mycarkigenpanel.Visible = true;

                menkyonew.Visible = true;
                //menkyonew.Value = null;

                syakennew.Visible = false;
                syakennew.Value = null;

                zibainew.Visible = false;
                zibainew.Value = null;

                ninninew.Visible = false;
                ninninew.Value = null;
            }
            else if (tuukinkubun.SelectedIndex == 6)
            {
                //徒歩
                c13.Enabled = false;
                c14.Enabled = false;
                c15.Enabled = false;

                mycarkigenpanel.Visible = false;

            }
            else if (tuukinkubun.SelectedIndex == 7)
            {
                //自転車
                c13.Enabled = true;
                c14.Enabled = true;
                c15.Enabled = true;

                mycarkigenpanel.Visible = true;

                menkyonew.Visible = false;
                menkyonew.Value = null;

                syakennew.Visible = false;
                syakennew.Value = null;

                zibainew.Visible = false;
                zibainew.Value = null;

                ninninew.Visible = true;
                //ninninew.Value = null;
            }


        }



        private void katakyori_ValueChanged(object sender, EventArgs e)
        {
            tuukinCalc();
         }

        private void tuukinteatekubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            tuukinCalc();
        }

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
                //case "8 車(実費精算)": flg = 0; kotei = 0; break;
                default: break;
            }

            //通勤1日単価
            decimal tanka = katakyori.Value * 15 * 2 * flg + kotei;

            //40overの場合の単価
            if (katakyori.Value > 40) tanka = 40 * 15 * 2 * flg + kotei;

            //概算通勤手当総額 ※
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

 
        private void syainno_TextChanged(object sender, EventArgs e)
        {
            if (syainno.Text == "")
            {
                button2.Text = "登録";
                tabControl1.Visible = false;

                button2.BackColor = Color.Transparent;
                button2.ForeColor = Color.Black;

                //出力ボタン
                syuturyoku.Visible = false;

                //エラー情報
                //label95.Visible = false;
                error.Visible = false;

                //label96.Visible = false;
                status.Visible = false;
            }
            else
            {
                button2.Text = "更新";
                tabControl1.Visible = true;
                button2.BackColor = Color.Blue;
                button2.ForeColor = Color.White;

                //出力ボタン
                syuturyoku.Visible = true;

                //エラー情報
                //label95.Visible = true;
                error.Visible = true;

                //label96.Visible = true;
                status.Visible = true;
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //Form2に送るテキスト
            string sendText = saiyoudate.Value.ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectEmpNyu.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。

            enko_no.Text = receiveText[0];
            enko_name.Text = receiveText[1];
            enko_soshiki.Text = receiveText[2] + ' ' + receiveText[3];
            enko_genba.Text = receiveText[4] + ' ' + receiveText[5];
        }

        //Excel出力ボタン
        private void button5_Click(object sender, EventArgs e)
        {
            //必須項目未入力有無チェック
            string msg = "";
            //if (saiyoudate.Text == "") msg += "入社年月日が選択されていません" + nl;
            if (genba.Text == "") msg += "地区名、組織名、現場名が選択されていません" + nl;
            if (sei.Text == "" || mei.Text == "" || seihuri.Text == "" || meihuri.Text == "") msg += "氏名が選択されていません" + nl;
            if (genba.Text == "") msg += "地区名、組織名、現場名が選択されていません" + nl;
            if (birthnew.Text == "") msg += "生年月日が入力されていません" + nl;
            if (seibetsu.Text == "") msg += "性別が選択されていません" + nl;
            //if (keitai.Text == "") msg += "電話番号が入力されていません" + nl;


            if (yuubin.Text == "") msg += "郵便番号が入力されていません" + nl;
            if (zyuusyo.Text == "") msg += "住所が入力されていません" + nl;
            if (katakyori.Value == 0) msg += "通勤距離が入力されていません。" + nl + "　2020年12月より通勤手段や契約区分に限らず、全従業員必須入力となります。" + nl + "　お手数ですがGoogleMapにて距離測定し、ご入力頂くようお願い致します。" + nl + "　※100m以内でしたら0.1を入力してください。" + nl;
            if (tuukinkubun.Text == "") msg += "通勤手段区分が入力されていません" + nl;

            if (tuukinkubun.Text == "1 車" || tuukinkubun.Text == "2 バイク" || tuukinkubun.Text == "6 送迎(知人・親族)" | tuukinkubun.Text == "4 バス・モノレール" | tuukinkubun.Text == "8 徒歩" | tuukinkubun.Text == "9 自転車")
            {
                if (katakyori.Text == "") msg += "片道通勤距離が入力されていません" + nl;

            }

            if (hatsurei.Text == "0001　正社員採用")
            {
                if (comboBoxSyokusyu.Text == "") msg += "職種が入力されていません" + nl; //*
                if (comboBoxGaku.Text == "") msg += "最終学歴が入力されていません" + nl; //*
                if (comboBoxKeiken.Text == "") msg += "社外経験が入力されていません" + nl; //*
            }
            else if (hatsurei.Text == "1100　日給者契約")
            {
                if (nikkyuu.Text == "") msg += "日給が入力されていません" + nl;

                if (comboBoxSyokusyu.Text == "") msg += "職種が入力されていません" + nl; //*
                if (comboBoxGaku.Text == "") msg += "最終学歴が入力されていません" + nl; //*
                if (comboBoxKeiken.Text == "") msg += "社外経験が入力されていません" + nl; //*
            }
            else
            {
                if (zikyuu.Text == "") msg += "時給が入力されていません" + nl;
            }

            if (msg != "")
            {
                MessageBox.Show(msg);
                return;
            }

            //一旦更新
            DataInsertUpdate();

            //テーブルより再取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.n入社データ出力 where 社員番号 = '" + syainno.Text + "'");
            
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("出力対象を選んでください。");
                return;
            }

            //資格
            DataTable sdt = new DataTable();
            sdt = Com.GetDB("select * from s資格データ取得 where 社員番号 = '" + syainno.Text + "'");

            //家族
            DataTable kdt = new DataTable();
            kdt = Com.GetDB("select * from 入社時家族出力 where 社員番号 = '" + syainno.Text + "'");

            //緊急連絡先
            DataTable edt = new DataTable();
            edt = Com.GetDB("select * from k緊急連絡先 where 社員番号 = '" + syainno.Text + "'");


            //ボタン無効化・カーソル変更
            syuturyoku.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;


            string fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\10_入社登録.xlsx";

            //手順1：新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            c1XLBook1.Load(fileName);

            // 手順2：セルに値を挿入します。
            XLSheet sheet = c1XLBook1.Sheets[4];

            //メインデータ
            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    sheet[j, i].Value = dt.Rows[i][j];
                }
            }

            //資格データ
            int r2 = sdt.Rows.Count;
            int c2 = sdt.Columns.Count;

            for (int i = 0; i < r2; i++)
            {
                for (int j = 0; j < c2; j++)
                {
                    sheet[j, i + 14].Value = sdt.Rows[i][j];
                }
            }

            //家族データ
            int r3 = kdt.Rows.Count;
            int c3 = kdt.Columns.Count;

            for (int i = 0; i < r3; i++)
            {
                for (int j = 0; j < c3; j++)
                {
                    sheet[j, i + 4].Value = kdt.Rows[i][j];
                }
            }

            //緊急連絡先データ
            int r4 = edt.Rows.Count;
            int c4 = edt.Columns.Count;

            for (int i = 0; i < r4; i++)
            {
                for (int j = 0; j < c4; j++)
                {
                    sheet[j + 12, i + 14].Value = edt.Rows[i][j];
                }
            }

            //その他情報
            sheet[49, 4].Value = DateTime.Now.ToString("yyyy年MM月dd日"); //作成日
            sheet[50, 4].Value = Program.loginname; //作成者
            sheet[51, 4].Value = saiyoudate.Value.AddMonths(1).ToString("yyyy年MM月支給 ") + saiyoudate.Value.ToString("(yyyy年MM月分)"); //給与計算月度　YYYY年MM月支給 (XXXX分)

            string localPass = @"C:\ODIS\Zinzi\";
            string exlName = localPass + DateTime.Now.ToString("_yyyy年MM月dd日_HH時mm分ss秒_") + syainno.Text + "_" + sei.Text + mei.Text + "_" + hatsurei.SelectedItem;

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            // 手順3：ファイルを保存します。
            c1XLBook1.Save(exlName + ".xlsx");

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            syuturyoku.Enabled = true;

            //excel出力
            System.Diagnostics.Process.Start(exlName + ".xlsx");

        }

        private void kikkake_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (kikkake.SelectedIndex == 1)
            {
                enkopanel.Visible = true;
            }
            else
            {
                enkopanel.Visible = false;

                enko_no.Text = "";
                enko_name.Text = "";
                enko_soshiki.Text = "";
                enko_genba.Text = "";
            }
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
                    Cmd.CommandText = "[dbo].[入社家族データ登録更新]";

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
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

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainno.Text;

                    if (this.kazokuid.Text == "")
                    {
                        DataTable kazokuid = new DataTable();
                        kazokuid = Com.GetDB("select max(家族識別ＩＤ) from dbo.入社家族データ where 社員番号 = '" + syainno.Text + "'");
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

                    //Cmd.Parameters["生年月日"].Value = kazoseinengappi.Value.ToString("yyyy/MM/dd");
                    if (kazoseinengappinew.Text == "")
                    {
                        Cmd.Parameters["生年月日"].Value = "";
                    }
                    else
                    {
                        Cmd.Parameters["生年月日"].Value = Convert.ToDateTime(kazoseinengappinew.Text).ToString("yyyy/MM/dd");
                    }

                    Cmd.Parameters["続柄区分"].Value = zokugara.Text;
                    Cmd.Parameters["同居区分"].Value = doukyokubun.Text.Substring(0, 1);

                    if (zokugara.Text == "00　夫" || zokugara.Text == "01　妻")
                    {
                        Cmd.Parameters["配偶者"].Value = 1; //妻or夫の場合
                    }
                    else
                    { 
                        Cmd.Parameters["配偶者"].Value = 0; //妻or夫の場合
                    }

                    Cmd.Parameters["税扶養区分"].Value = huyoukubun.Text.Substring(0,1);
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

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void DataDispKazoku()
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from 入社家族データ取得 where 社員番号 = '" + syainno.Text + "'");
            kazokudgv.DataSource = dt;

            if (dt.Rows.Count == 0) return;

            Int64 gaku = 0;
            foreach (DataRow row in dt.Rows)
            {
                gaku += Convert.ToInt64(row["手当額"]);
            }

            huyougaku.Text = gaku.ToString("#,0");


        }


        private void GetKazokukoumokuData(string str)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from 入社家族データ where 社員番号 = '" + syainno.Text + "' and 家族識別ＩＤ = '" + str + "'");

            kazosei.Text = dt.Rows[0][2].ToString();
            kazomei.Text = dt.Rows[0][3].ToString();
            kazokanasei.Text = dt.Rows[0][4].ToString();
            kazokanamei.Text = dt.Rows[0][5].ToString();
            if (dt.Rows[0][6].ToString() == "")
            {
                kazoseinengappinew.Value = null;
            }
            else
            {
                kazoseinengappinew.Value = Convert.ToDateTime(dt.Rows[0][6].ToString());
            }
            zokugara.SelectedIndex = zokugara.FindString(dt.Rows[0][7].ToString());
            doukyokubun.SelectedIndex = doukyokubun.FindString(dt.Rows[0][8].ToString());
            huyoukubun.SelectedIndex = huyoukubun.FindString(dt.Rows[0][10].ToString());
            gensenkubun.SelectedIndex = gensenkubun.FindString(dt.Rows[0][11].ToString());
            kenpokanyuu.SelectedIndex = kenpokanyuu.FindString(dt.Rows[0][12].ToString());
            setainushi.SelectedIndex = setainushi.FindString(dt.Rows[0][13].ToString());

            if (dt.Rows[0][15].ToString() == "1")
            {
                //特別障碍者の場合
                syougaikubun.SelectedIndex = syougaikubun.FindString("2");

            }
            else
            { 
                syougaikubun.SelectedIndex = syougaikubun.FindString(dt.Rows[0][14].ToString());
            }

            kazokuid.Text = str;
        }



        //家族登録更新ボタン
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
                GetKazokukoumokuData(drv[1].ToString());

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


            DataInsertUpdate();
        }


        private void shikakubtn_Click(object sender, EventArgs e)
        {
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

            if (shikakubtn.Text == "資格更新")
            {
                //更新時の処理
                dgvRow = shikakudgv.CurrentCell.RowIndex;

                //一覧データ取得
                DataDispShikaku();

                shikakudgv.CurrentCell = shikakudgv[1, dgvRow];

                DataGridViewRow dgr = shikakudgv.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;

                //TODO:0でOK？　下にも同じ処理ある
                GetShikakukoumokuData(drv[0].ToString());

                MessageBox.Show("資格情報を更新しました。");

            }
            else
            {
                //一覧データ取得
                DataDispShikaku();

                //選択無
                shikakudgv.CurrentCell = null;

                //項目クリア
                ShikakuClear();

                MessageBox.Show("資格情報を登録しました。");
            }

            //本データも更新
            DataInsertUpdate();
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
                    Cmd.Parameters["適用開始日"].Value = saiyoudate.Value.ToString("yyyy/MM/dd");

                    //資格コードのみにする
                    string[] del = { "　" };
                    string[] shikakucd = shikakutextb.Text.Split(del, StringSplitOptions.None);
                    Cmd.Parameters["資格コード"].Value = shikakucd[0];
                    Cmd.Parameters["個人識別ＩＤ"].Value = "";
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
                    Cmd.Parameters["更新日時"].Value = DateTime.Now;
                    Cmd.Parameters["更新ユーザＩＤ"].Value = "24";
                    Cmd.Parameters["更新者"].Value = Program.loginname;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
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

            shikakubtn.Text = "資格更新";
            shikakubtn.BackColor = Color.Blue;
            shikakubtn.ForeColor = Color.White;

            //更新の時は免許の変更はできない
            //shikakuselect.Enabled = false;

            delshikaku.Visible = true;

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
                GetKazokukoumokuData(drv[1].ToString());
            }

            
            //TODO 資格登録更新ボタン表示の変更
            kazokubtn.Text = "家族更新";
            kazokubtn.BackColor = Color.Blue;
            kazokubtn.ForeColor = Color.White;

            delkazoku.Visible = true;
        }

        private void kazokunew_Click(object sender, EventArgs e)
        {
            KazokuClear();
        }

        private void shikakunew_Click(object sender, EventArgs e)
        {
            ShikakuClear();
        }

        private void huyourbmu_CheckedChanged(object sender, EventArgs e)
        {
            if (huyourbmu.Checked == true)
            {
                c20.Enabled = false;
                tabControl1.TabPages.Remove(this.kazokutab);
            }
            else
            {
                c20.Enabled = true;
                tabControl1.TabPages.Insert(tabControl1.TabCount, this.kazokutab);
            }

            ErrorCheck();
        }

        private void hoyuurbmu_CheckedChanged(object sender, EventArgs e)
        {
            if (hoyuurbmu.Checked == true)
            {
                c19.Enabled = false;
                tabControl1.TabPages.Remove(this.shikakutab);
            }
            else
            {
                c19.Enabled = true;
                tabControl1.TabPages.Insert(tabControl1.TabCount, this.shikakutab);
            }

            ErrorCheck();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            sei.Text = "姓漢字";
            seihuri.Text = "ｾｲｶﾀｶﾅ";
            mei.Text = "名漢字";
            meihuri.Text = "ﾒｲｶﾀｶﾅ";
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

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void menkyogaku_TextChanged(object sender, EventArgs e)
        {
            if (hatsurei.Text == "0001　正社員採用" || hatsurei.Text == "1100　日給者契約")
            {
                menkyoteate.Text = menkyogaku.Text;
            }
            else
            {
                menkyoteate.Text = "0";
            }
        }

        private void huyougaku_TextChanged(object sender, EventArgs e)
        {
            if (hatsurei.Text == "0001　正社員採用")
            {
                huyou.Text = huyougaku.Text;
            }
            else
            {
                huyou.Text = "0";
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            shikakudgv.CurrentCell = null;
            kazokudgv.CurrentCell = null;

            ShikakuClear();
            KazokuClear();
        }

        private void tomokubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (hatsurei.Text == "0900　パート契約" || hatsurei.Text == "1100　日給者契約" || hatsurei.Text == "0001　正社員採用")
            {
                if (tomokubun.Text == "1　非加入")
                {
                    tomonokai.Text = "0";
                }
                else
                {
                    tomonokai.Text = "300";
                }
            }
            else if (hatsurei.Text == "1000　アルバイト契約")
            {
                if (tomokubun.Text == "2　アルバイト加入")
                {
                    tomonokai.Text = "300";
                }
                else
                {
                    tomonokai.Text = "0";
                }
            }

        }

        //入社資料
        private void cXX_CheckedChanged(object sender, EventArgs e)
        {
            ErrorCheck();
        }

        //エラーチェック
        private void ErrorCheck()
        {
            error.Text = "";

            //発令区分
            if (hatsurei.Text == "") error.Text += "発令区分　未入力" + nl;

            //氏名チェック
            if (sei.Text == "" || mei.Text == "" || seihuri.Text == "" | meihuri.Text == "") error.Text += "氏名　未入力" + nl;

            //地区・組織・現場チェック
            if (tiku.Text == "" || soshiki.Text == "" || genba.Text == "") error.Text += "地区or組織or現場　未入力" + nl;

            //電話番号・郵便・住所
            if (yuubin.Text == "" || zyuusyo.Text == "") error.Text += "郵便or住所　未入力" + nl;

            //TODO 緊急連絡先の警告
            if (honkeitai.Text == "") error.Text += "本人携帯電話未入力" + nl;

            if (kaz1name.Text == "") error.Text += "ご家族優先1_名前　未入力" + nl;
            if (kaz1kana.Text == "") error.Text += "ご家族優先1_カナ名　未入力" + nl;
            if (kaz1gara.Text == "") error.Text += "ご家族優先1_続柄　未入力" + nl;
            if (kaz1no.Text == "") error.Text += "ご家族優先1_電話番号　未入力" + nl;

            if (kaz2name.Text == "") error.Text += "ご家族優先2_名前　未入力" + nl;
            if (kaz2kana.Text == "") error.Text += "ご家族優先2_カナ名　未入力" + nl;
            if (kaz2gara.Text == "") error.Text += "ご家族優先2_続柄　未入力" + nl;
            if (kaz2no.Text == "") error.Text += "ご家族優先2_電話番号　未入力" + nl;

            //距離
            if (katakyori.Value == 0) error.Text += "距離未入力" + nl;
 
            //入社資料カウント
            int ct = 0;
            int ctall = 19;

            if (c01.Checked == false) ct++;
            if (c02.Checked == false) ct++;
            if (c03.Checked == false) ct++;
            if (c04.Checked == false) ct++;
            if (c05.Checked == false) ct++;
            if (c06.Checked == false) ct++;
            if (c07.Checked == false) ct++;
            if (c08.Checked == false) ct++;
            if (c09.Checked == false) ct++;
            if (c10.Checked == false) ct++;
            if (c11.Checked == false) ct++;
            if (c12.Checked == false) ct++;

            if (c13.Enabled == false)
            {
                ctall--;
            }
            else
            {
                if (c13.Checked == false) ct++;
            }

            if (c14.Enabled == false)
            {
                ctall--;
            }
            else
            {
                if (c14.Checked == false) ct++;
            }

            if (c15.Enabled == false)
            {
                ctall--;
            }
            else
            {
                if (c15.Checked == false) ct++;
            }

            if (c16.Checked == false) ct++;
            //if (c17.Checked == false) ct++;
            if (c18.Checked == false) ct++;

            if (c19.Enabled == false)
            {
                ctall--;
            }
            else
            {
                if (c19.Checked == false) ct++;
            }

            if (c20.Enabled == false)
            {
                ctall--;
            }
            else
            {
                if (c20.Checked == false) ct++;
            }

            if (ct > 0) error.Text += "入社資料" + ctall.ToString() + "点中" + ct.ToString() + "点未チェック" + nl;

            //TODO ついでに年齢入力
            if (birthnew.Text != "")
            {
                int age = saiyoudate.Value.Year - Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy"));

                //誕生日がまだ来ていなければ、1引く
                if (saiyoudate.Value.Month < Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("MM")) ||
                    (saiyoudate.Value.Month == Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("MM")) &&
                    saiyoudate.Value.Day < Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("dd"))))
                {
                    age--;
                }

                old.Text = Convert.ToString(age);
            }
        }

        private void SimeiTextChanged(object sender, EventArgs e)
        {
            ErrorCheck();
        }

        private void genba_SelectedValueChanged(object sender, EventArgs e)
        {
            ErrorCheck();
        }

        private void keitai_TextChanged(object sender, EventArgs e)
        {
            ErrorCheck();

        }

        private void yuubin_TextChanged(object sender, EventArgs e)
        {
            ErrorCheck();
        }

        private void zyuusyo_TextChanged(object sender, EventArgs e)
        {
            ErrorCheck();
        }

        /// <summary>
        /// 指定された文字列が電話番号かどうかを返します
        /// </summary>
        public static bool IsPhoneNumber(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return false;
            }
            return Regex.IsMatch(
                input,
                @"^0\d{1,4}-\d{1,4}-\d{4}$"
            );
        }

        private void shikakusyutokubi_ValueChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\入社入力及び異動入力操作方法.xlsx"); return;
        }



        private void delshikaku_Click(object sender, EventArgs e)
        {
            //TODO 更新モードの場合のみ表示
            DataGridViewRow dgr = shikakudgv.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            DialogResult result = MessageBox.Show("ホントに削除していいっすか？" + nl + drv[0].ToString() + "　" + drv[1].ToString(),
                            "警告",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation,
                            MessageBoxDefaultButton.Button2);

            //何が選択されたか調べる
            if (result == DialogResult.Yes)
            {
                //「はい」が選択された時
                DataTable dtdel = new DataTable();
                dtdel = Com.GetDB("delete from QUATRO.dbo.SJMTSHIKAK where 会社コード = 'E0' and 個人識別ＩＤ = '' and 社員番号 = '" + syainno.Text + "' and 資格コード = '" + drv[0].ToString() + "' ");

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

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\10_標準書式\雛形\00_従業員の採用異動手順と表紙.xlsx"); return;
        }

        private void kinmu_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetHonkyuu(); 
        }

        private void kyuuka_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetHonkyuu();
        }

        private void delkazoku_Click(object sender, EventArgs e)
        {
            //TODO 更新モードの場合のみ表示
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
                dtdel = Com.GetDB("delete from 入社家族データ where 社員番号 = '" + syainno.Text + "' and 家族識別ＩＤ = '" + drv[1].ToString() + "'");

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

        private void zokugara_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (zokugara.Text == "00　夫" || zokugara.Text == "01　妻")
            {
                label84.Visible = true;
                gensenkubun.Visible = true;
            }
            else
            {
                label84.Visible = false;
                gensenkubun.Visible = false;
                gensenkubun.SelectedIndex = 0;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (warekibirth.SelectedItem == null) return;

            int mm = 1;
            int dd = 1;

            if (birthnew.Text != "")
            {
                mm = Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("MM"));
                dd = Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("dd"));
            }





            if (warekibirth.SelectedItem.ToString().Substring(0, 2) == "昭和")
            {
                int i = Convert.ToInt16(warekibirth.SelectedItem.ToString().Substring(2, 2)) + 1925;
                birthnew.Value = new DateTime(i, mm, dd);
                return;
            }

            if (warekibirth.SelectedItem.ToString() == "平成元年")
            {
                birthnew.Value = new DateTime(1989, mm, dd);
                return;
            }


            if (warekibirth.SelectedItem.ToString().Substring(0, 2) == "平成")
            {
                //平成1桁対応
                if (warekibirth.SelectedItem.ToString().Substring(3, 1) == "年")
                {
                    int i = Convert.ToInt16(warekibirth.SelectedItem.ToString().Substring(2, 1)) + 1988;
                    birthnew.Value = new DateTime(i, mm, dd);
                }
                else
                {
                    int i = Convert.ToInt16(warekibirth.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    birthnew.Value = new DateTime(i, mm, dd);
                }
            }

        }

        private void yu_Click(object sender, EventArgs e)
        {
            //Form2に送るテキスト
            string sendText = saiyoudate.Value.ToString("yyyy/MM/dd");

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = SelectAdress.ShowMiniForm(sendText);　//Form2を開く

            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            yuubin.Text = receiveText[0];
            zyuusyo.Text = receiveText[1];
        }



        private void tuukinhi2_ValueChanged(object sender, EventArgs e)
        {
            tuukinhi.Text = tuukinhi2.Value.ToString();
        }

        private void tuukinka2_ValueChanged(object sender, EventArgs e)
        {
            tuukinka.Text = tuukinka2.Value.ToString();
        }

        private void filelink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string localPass = @"C:\ODIS\Zinzi\";

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            System.Diagnostics.Process.Start(localPass); return;
        }

        #region 和暦対応
        private void wareki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki.SelectedItem == null) return;

            //if (menkyonew.Text == "") menkyonew.Value = DateTime.Today;
            if (menkyonew.Text == "")
            {
                if (wareki.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    menkyonew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                }

                if (wareki.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    menkyonew.Value = new DateTime(2019, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    return;
                }

                if (wareki.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        menkyonew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        menkyonew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                }
            }
            else
            {
                if (wareki.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    menkyonew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("dd")));
                }

                if (wareki.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    menkyonew.Value = new DateTime(2019, Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("dd")));
                    return;
                }

                if (wareki.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        menkyonew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        menkyonew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("dd")));
                    }
                }
            }

                
        }

        private void wareki2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki2.SelectedItem == null) return;

            //if (syakennew.Text == "") syakennew.Value = DateTime.Today;

            if (syakennew.Text == "")
            {
                if (wareki2.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    syakennew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                }

                if (wareki2.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    syakennew.Value = new DateTime(2019, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    return;
                }

                if (wareki2.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki2.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        syakennew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        syakennew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                }
            }
            else
            {
                if (wareki2.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    syakennew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("dd")));
                }

                if (wareki2.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    syakennew.Value = new DateTime(2019, Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("dd")));
                    return;
                }

                if (wareki2.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki2.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        syakennew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki2.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        syakennew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("dd")));
                    }
                }
            }

                
        }

        private void wareki3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki3.SelectedItem == null) return;

            if (zibainew.Text == "") 
            {
                if (wareki3.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    zibainew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                }

                if (wareki3.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    zibainew.Value = new DateTime(2019, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    return;
                }

                if (wareki3.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki3.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        zibainew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        zibainew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                }
            }
            else
            {
                if (wareki3.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    zibainew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("dd")));
                }

                if (wareki3.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    zibainew.Value = new DateTime(2019, Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("dd")));
                    return;
                }

                if (wareki3.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki3.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        zibainew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki3.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        zibainew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("dd")));
                    }
                }
            }
        }

        private void wareki4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki4.SelectedItem == null) return;

            //if (ninninew.Text == "") ninninew.Value = DateTime.Today;

            if (ninninew.Text == "")
            {
                if (wareki4.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    ninninew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                }

                if (wareki4.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    ninninew.Value = new DateTime(2019, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    return;
                }

                if (wareki4.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki4.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        ninninew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        ninninew.Value = new DateTime(i, Convert.ToInt16(DateTime.Today.ToString("MM")), Convert.ToInt16(DateTime.Today.ToString("dd")));
                    }
                }
            }
            else
            {
                if (wareki4.SelectedItem.ToString().Substring(0, 2) == "平成")
                {
                    int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 2)) + 1988;
                    ninninew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("dd")));
                }

                if (wareki4.SelectedItem.ToString().Substring(0, 4) == "令和元年")
                {
                    ninninew.Value = new DateTime(2019, Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("dd")));
                    return;
                }

                if (wareki4.SelectedItem.ToString().Substring(0, 2) == "令和")
                {
                    //令和1桁対応
                    if (wareki4.SelectedItem.ToString().Substring(3, 1) == "年")
                    {
                        int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 1)) + 2018;
                        ninninew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("dd")));
                    }
                    else
                    {
                        int i = Convert.ToInt16(wareki4.SelectedItem.ToString().Substring(2, 2)) + 2018;
                        ninninew.Value = new DateTime(i, Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("MM")), Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("dd")));
                    }
                }
            }

            
        }



        private void menkyonew_ValueChanged(object sender, EventArgs e)
        {
            if (menkyonew.Text == "")
            {
                wareki.SelectedIndex = -1;
                return;
            }

            if (Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("yyyy")) > 1989)
            {
                wareki.SelectedIndex = wareki.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("yyyy")) == 2019)
            {
                wareki.SelectedIndex = wareki.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("yyyy")) > 2019)
            {
                wareki.SelectedIndex = wareki.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(menkyonew.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        private void syakennew_ValueChanged(object sender, EventArgs e)
        {
            if (syakennew.Text == "")
            {
                wareki2.SelectedIndex = -1;
                return;
            }

            if (Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("yyyy")) > 1989)
            {
                wareki2.SelectedIndex = wareki2.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("yyyy")) == 2019)
            {
                wareki2.SelectedIndex = wareki2.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("yyyy")) > 2019)
            {
                wareki2.SelectedIndex = wareki2.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(syakennew.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        private void zibainew_ValueChanged(object sender, EventArgs e)
        {
            if (zibainew.Text == "")
            {
                wareki3.SelectedIndex = -1;
                return;
            }

            if (Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("yyyy")) > 1989)
            {
                wareki3.SelectedIndex = wareki3.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("yyyy")) == 2019)
            {
                wareki3.SelectedIndex = wareki3.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("yyyy")) > 2019)
            {
                wareki3.SelectedIndex = wareki3.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(zibainew.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        private void ninninew_ValueChanged(object sender, EventArgs e)
        {
            if (ninninew.Text == "")
            {
                wareki4.SelectedIndex = -1;
                return;
            }

            if (Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("yyyy")) > 1989)
            {
                wareki4.SelectedIndex = wareki4.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("yyyy")) == 2019)
            {
                wareki4.SelectedIndex = wareki4.FindString("令和元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("yyyy")) > 2019)
            {
                wareki4.SelectedIndex = wareki4.FindString("令和" + (Convert.ToInt16(Convert.ToDateTime(ninninew.Value).ToString("yyyy")) - 2018).ToString() + "年");
            }
        }

        #endregion

        private void birthnew_ValueChanged(object sender, EventArgs e)
        {
            //TODO 200131追加
            if (birthnew.Text == "") return;

            //wareki add
            if (Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy")) < 1989)
            {
                warekibirth.SelectedIndex = warekibirth.FindString("昭和" + (Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy")) - 1925).ToString() + "年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy")) == 1989)
            {
                warekibirth.SelectedIndex = warekibirth.FindString("平成元年");
            }

            if (Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy")) > 1989)
            {
                warekibirth.SelectedIndex = warekibirth.FindString("平成" + (Convert.ToInt16(Convert.ToDateTime(birthnew.Value).ToString("yyyy")) - 1988).ToString() + "年");
            }

            ErrorCheck();
        }

        private void kazoseinengappinew_ValueChanged(object sender, EventArgs e)
        {
            //TODO 200131追加
            if (kazoseinengappinew.Text == "") return;


            //wareki add
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
                //int i = Convert.ToInt16(warekicb.SelectedItem.ToString().Substring(2, 2)) + 1925;
                //kazoseinengappinew.Value = new DateTime(i, mm, dd);
                //return;

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

        private void ymlist_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "入社月" && Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01") < Convert.ToDateTime("2020/04/01"))
            {
                MessageBox.Show("2020年4月出勤簿から対象です。");
                return;
            }

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button4.Enabled = false;

            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from s出勤簿データ取得_入社('" + ymlist.SelectedItem.ToString().Replace("/", "") + "') where 社員番号 = '" + syainno.Text + "'order by 組織名, 現場名, カナ名");

            if (dt.Rows.Count == 0)
            {

            }
            else
            {
                if (Program.loginbusyo == "施設")
                {
                    if (((Button)sender).Text == "入社月")
                    { 
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(0),false, false);
                    }
                    else
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(1), false, false);
                    }
                }
                else
                {
                    if (((Button)sender).Text == "入社月")
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(0), true, false);
                    }
                    else
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(1), true, false);
                    }
                }
            }


            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button4.Enabled = true;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "入社月" && Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01") < Convert.ToDateTime("2020/04/01"))
            {
                MessageBox.Show("2020年4月出勤簿から対象です。");
                return;
            }

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button5.Enabled = false;

            //対象データ取得
            DataTable dt = new DataTable();
            string sql = "";
            sql = "select * from s出勤簿データ取得_入社('" + ymlist.SelectedItem.ToString().Replace("/", "") + "') where 入社年月日 like '" + ymlist.SelectedItem.ToString() + "%' and 担当区分 ";

            if (Program.loginname == "親泊　美和子" || Program.loginname == "石井　優子" || Program.loginname == "下地　明香里" || Program.loginname == "小園　玲奈")
            {
                sql = sql + " in ('03_施設', '04_エンジ') ";
            }
            //else if (Program.loginname == "大浜　綾希子")
            //{
            //    sql = sql + " in ('01_現業', '02_客室') ";
            //}
            else
            {
                sql = sql + " like '%" + Program.loginbusyo + "%'";
            }

                sql = sql + " order by 組織名, 現場名, カナ名";

            dt = Com.GetDB(sql);

            if (dt.Rows.Count == 0)
            {

            }
            else
            {
                if (Program.loginbusyo == "03_施設")
                { 
                    if (((Button)sender).Text == "入社月")
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(0), false,false);
                    }
                    else
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(1), false, false);
                    }
                }
                else
                {
                    if (((Button)sender).Text == "入社月")
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(0), true, false);
                    }
                    else
                    {
                        Com.GetSyukkinbo(dt, Convert.ToDateTime(ymlist.SelectedItem.ToString() + "/01").AddMonths(1), true, false);
                    }
                }

            }



            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button5.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (syainno.Text == "") return;

            string msg = "入社年月日：" + saiyoudate.Value.ToString("yyyy/MM/dd") + nl;
            msg += "氏名：" + sei.Text + "　" + mei.Text + nl;
            msg += "組織名：" + soshiki.Text + nl;
            msg += "現場名：" + genba.Text + nl;

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
                dtdel = Com.GetDB("delete from dbo.n入社データ where 社員番号 = '" + syainno.Text + "' and 採用年月日 = '" + saiyoudate.Value.ToString("yyyy/MM/dd") + "' ");

                //入社一覧取得
                GetData();

                //入力フォームクリア
                AllClear();

                MessageBox.Show("消し去りましたー");
            }
            else if (result == DialogResult.No)
            {
                //「いいえ」が選択された時
            }
        }

        private void genba_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (tenkin.Text == "")
            {

            }
            else if (Convert.ToDecimal(tenkin.Text) > 0)
            {
                ritou.Text = "0";
                return;
            }
            if (soshiki.SelectedItem == null || genba.SelectedItem == null || comboBoxSyokusyu.SelectedItem == null) return;
            if (kyuuyo.Text.Substring(0, 2) != "C1") return;
            ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
        }

        private void syukkou_TextChanged(object sender, EventArgs e)
        {
            //転勤手当

            if (tenkin.Text == "")
            {

            }
            else if (Convert.ToDecimal(tenkin.Text) > 0)
            {
                ritou.Text = "0";
                return;
            }
            if (soshiki.SelectedItem == null || genba.SelectedItem == null || comboBoxSyokusyu.SelectedItem == null) return;
            if (kyuuyo.Text.Substring(0, 2) != "C1") return;
            ritou.Text = Com.RitouCalc(soshiki.SelectedItem.ToString().Substring(0, 5), comboBoxSyokusyu.SelectedItem.ToString());
        }

        private void button9_Click(object sender, EventArgs e)
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

        private void shikakutextb_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxTenkin_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTenkin.Text == "")
            {
                tenkin.Text = "0";
            }
            else
            {
                tenkin.Text = "50,000";
            }
        }

        private void ritou_TextChanged(object sender, EventArgs e)
        {
            if (hatsurei.Text == "0001　正社員採用")
            {
                ritou.Text = ritou.Text;
            }
            else
            {
                ritou.Text = "0";
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            if (syainno.Text == "") return;

            //TODO 従業員検索に同じ処理があります。。
            //TODO 異動にも同じ処理があります。。

            //if (BeforeCheck()) return;

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
                    Cmd.CommandText = "[dbo].[r労働条件更新_入社時]";
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

                    Cmd.Parameters.Add(new SqlParameter("入社時_厚生年金", SqlDbType.Char)); Cmd.Parameters["入社時_厚生年金"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("入社時_健康保険", SqlDbType.Char)); Cmd.Parameters["入社時_健康保険"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("入社時_雇用保険", SqlDbType.Char)); Cmd.Parameters["入社時_雇用保険"].Direction = ParameterDirection.Input;

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

                    Cmd.Parameters["入社時_厚生年金"].Value = kouseicb.Text;
                    Cmd.Parameters["入社時_健康保険"].Value = kenkoucb.Text;
                    Cmd.Parameters["入社時_雇用保険"].Value = koyoucb.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        //MessageBox.Show("更新しました。");
                    }
                }
            }
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

            if (!keiyakunengetsu.Value.Equals(DBNull.Value))
            { 
                if (Convert.ToDateTime(keiyakunengetsu?.Value) > saiyoudate?.Value) msg += "作成日が入社日後になってる。。" + Com.nl;
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
            dt = Com.GetDB("select * from r労働条件取得_入社入力 where 社員番号 = '" + syainno.Text + "'");

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
            if (kyuuyo.Text.Substring(0,2) == "E1" || kyuuyo.Text.Substring(0, 2) == "F1")
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

        private void koyoukubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (koyoukubun.SelectedItem?.ToString() == "1 期間の定めあり")
            {
                //koyoukaishibi.Visible = true;
                koyousyuuryoubi.Visible = true;
            }
            else
            {
                //koyoukaishibi.Value = null;
                //koyoukaishibi.Visible = false;

                koyousyuuryoubi.Value = null;
                koyousyuuryoubi.Visible = false;
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

        private void label144_Click(object sender, EventArgs e)
        {

        }

        private void roudoutab_Click(object sender, EventArgs e)
        {

        }

        private void saiyoudate_ValueChanged(object sender, EventArgs e)
        {
            //時給と日給の最賃設定
            DataTable tbsai = Com.GetDB("select * from dbo.HK_本給 where '" + saiyoudate.Value.ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

            tbsai.Rows[0]["最賃"].ToString();
            zikyuu.Minimum = Convert.ToInt32(tbsai.Rows[0]["最賃"].ToString());
            nikkyuu.Minimum = Convert.ToInt32(tbsai.Rows[0]["最賃"].ToString()) * 7;
        }

        private void keiyaku_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (keiyaku.SelectedItem?.ToString() == "30　技・特実習生")
            //{
            //    //技能実習生は変更可能に。
            //    honkyuu.Enabled = true;
            //    syokumu.Enabled = true;

            //    honkyuu.ReadOnly = false;
            //    syokumu.ReadOnly = false;
            //}
            //else
            //{
                honkyuu.Enabled = false;
                syokumu.Enabled = false;

                honkyuu.ReadOnly = true;
                syokumu.ReadOnly = true;
            //}
        }

        private void tourokugaku_TextChanged(object sender, EventArgs e)
        {
            if (hatsurei.Text == "0001　正社員採用")
            {
                tourokuteate.Text = tourokugaku.Text;
            }
            else
            {
                tourokuteate.Text = "0";
            }
        }

        private void honkeitai_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "本人携帯電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void honkotei_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "本人固定電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kaz1no_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "ご家族優先1電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kaz2no_Validating(object sender, CancelEventArgs e)
        {
            ValidCom((TextBox)sender, e, "ご家族優先2電話", @"^0\d{1,4}-\d{1,4}-\d{4}$");
        }

        private void kazkana_Validating(object sender, CancelEventArgs e)
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

        private void kyuuzitsukubun_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
