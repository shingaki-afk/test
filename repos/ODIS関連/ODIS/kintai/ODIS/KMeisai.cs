using C1.C1Excel;
using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KMeisai : Form
    {
        private DataTable dt = new DataTable();
        private DateTime ym;
        private DateTime ymex;

        private TargetDays td = new TargetDays();
        private Int32 maxymd;

        private DataTable dtkouzyo = new DataTable();

        private string selectnum;

        public KMeisai()
        {
            // ここを書き換え
            if (!int.TryParse(Program.access, out var access) || access == 1)
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("39_給与明細表示権限無", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //列ヘッダーを非表示にする
            dataGridView2.ColumnHeadersVisible = false;
            dataGridView3.ColumnHeadersVisible = false;

            comboBox1.Items.Add("2013");
            comboBox1.Items.Add("2014");
            comboBox1.Items.Add("2015");
            comboBox1.Items.Add("2016");
            comboBox1.Items.Add("2017");
            comboBox1.Items.Add("2018");
            comboBox1.Items.Add("2019");
            comboBox1.Items.Add("2020");
            comboBox1.Items.Add("2021");
            comboBox1.Items.Add("2022");
            comboBox1.Items.Add("2023");
            comboBox1.Items.Add("2024");
            comboBox1.Items.Add("2025");
            comboBox1.Items.Add("2026");
            comboBox1.Items.Add("2027");
            comboBox1.Items.Add("2028");
            comboBox1.Items.Add("2029");
            comboBox1.Items.Add("2030");
            //TODO 毎年追加しなければならない

            comboBox2.Items.Add("01");
            comboBox2.Items.Add("02");
            comboBox2.Items.Add("03");
            comboBox2.Items.Add("04");
            comboBox2.Items.Add("05");
            comboBox2.Items.Add("06");
            comboBox2.Items.Add("07");
            comboBox2.Items.Add("08");
            comboBox2.Items.Add("09");
            comboBox2.Items.Add("10");
            comboBox2.Items.Add("11");
            comboBox2.Items.Add("12");

            comboBox1.SelectedItem = td.StartYMD.AddMonths(1).ToString("yyyy");
            comboBox2.SelectedItem = td.StartYMD.AddMonths(1).ToString("MM");

            numericUpDown1.Value = 50000;

            //先月比は非表示
            checkBox1.Checked = true;

            GetCount();

            Com.InHistory("39_給与明細表示", "", "");

        }

        private void GetCount()
        {
            maxymd = Convert.ToInt32(td.StartYMD.AddMonths(1).ToString("yyyyMM"));
            GetMeisai();
        }

        private void GetMeisai()
        {
            dt.Clear();
            dtkouzyo.Clear();

            string y = ym.Year.ToString();
            string m = ym.ToString("MM");
            string ymd = ym.AddDays(-1).ToString("yyyy/MM/dd");
            string exy = ymex.Year.ToString();
            string exm = ymex.ToString("MM");
            string exymd = ymex.AddDays(-1).ToString("yyyy/MM/dd");

            dtkouzyo = Com.GetDB("select 社員番号, 項目, 内容, 金額 from dbo.k固定変動控除と変動手当 where 処理年 = '" + y + "' and 処理月 = '" + m + "' and 金額 <> 0");

            string strCon = Common.constr;
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter Adapter;

            using (Cn = new SqlConnection(strCon))
            {
                Cmd = Cn.CreateCommand();

                string sql = "select * from dbo.KM_給与明細 a ";
                sql += "left join dbo.KM_給与明細 b on b.処理年 = '" + exy + "' and b.処理月 = '" + exm + "' and a.社員番号 = b.社員番号 ";
                sql += "left join dbo.担当テーブル c on a.組織CD = c.組織CD and a.現場CD = c.現場CD ";
                sql += "left join dbo.発令一覧('" + ym.AddMonths(-1).ToString("yyyy/MM") + "') d on a.社員番号 = d.社員番号 ";
                //sql += "left join QUATRO.dbo.QCTTMSG e on 会社コード = 'E0' and 年 = '" + y + "' and 月 = '" + m + "' and a.社員番号 = e.社員番号 ";
                sql += "left join dbo.web明細お知らせ e on e.処理年 = '" + y + "' and e.処理月 = '" + m + "' and a.社員番号 = e.社員番号 ";


                sql += "where a.処理年 = '" + y + "' and a.処理月 = '" + m + "' ";
                

                if (checkBox2.Checked)
                {
                    sql += "and 担当管理 like '%" + Program.loginname + "%' ";
                }


                //役職別で参照設定
                //役員
                //取締役部長(0060)・相談役(0066)・顧問(0070)・
                //部長(0110 / 0102)
                //副部長(0120 / 0112)
                //課長(0130 / 0122)


                if (Convert.ToInt16(Program.yakusyokucd) < 60 || Convert.ToInt16(Program.dispZinzi) == 99)
                {
                    //対象範囲：役員/シス管/人給管が参照可能
                    //参照範囲：全て
                }
                else if (Convert.ToInt16(Program.yakusyokucd) <= 110)
                {
                    //対象範囲：部門長以上
                    //参照範囲：役員を除く
                    sql += "and a.所定 + a.公休 > 0 and a.勤務時間 is not null and a.組織名 <> '役員室'";
                }
                else if (Convert.ToInt16(Program.yakusyokucd) <= 130)
                {
                    //対象範囲：課長以上
                    //参照範囲：組織名(役員室)と副部長以上を除く
                    sql += "and a.所定 + a.公休 > 0 and a.勤務時間 is not null and a.組織名 <> '役員室' and a.役職名 not like '%部長'";
                }
                else //if (Convert.ToInt16(Program.yakusyokucd) < 130)
                {
                    //対象範囲：その他
                    //参照範囲：現場名が事務所以外の所属部署
                    sql += "and 担当区分 like '%" + Program.loginbusyo + "%' and a.所定 + a.公休 > 0 and a.勤務時間 is not null and a.現場名 not like '事務所%' ";
                }


                string res = textBox1.Text.Trim().Replace("　", " ");
                string[] ar = res.Split(' ');
                if (ar[0] != "")
                {
                    foreach (string s in ar)
                    {
                        sql += " and (a.reskey like '%" + s + "%' or a.reskey like '%" + Com.isOneByteChar(s) + "%' or a.reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or a.reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or a.reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or a.reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }

                //order by 対応
                sql += " order by a.現場CD, a.カナ氏名";

                Cmd.CommandText = sql;
                Adapter = new SqlDataAdapter(Cmd);
                Adapter.Fill(dt);
            }


            DataTable disp = new DataTable();
            disp.Columns.Add("社員番号", typeof(string));
            disp.Columns.Add("氏名", typeof(string));
            disp.Columns.Add("組織名", typeof(string));
            disp.Columns.Add("現場名", typeof(string));
            disp.Columns.Add("発令", typeof(string));
            disp.Columns.Add("基本情報変更", typeof(string));
            disp.Columns.Add("固定給変更", typeof(string));
            //disp.Columns.Add("固定控除変更", typeof(string));
            disp.Columns.Add("支給差額", typeof(string));
            disp.Columns.Add("本給", typeof(int));
            disp.Columns.Add("職務技能給", typeof(int));
            disp.Columns.Add("調整手当", typeof(int));
            disp.Columns.Add("特別手当", typeof(int));
            disp.Columns.Add("皆勤手当", typeof(int));
            disp.Columns.Add("役職手当", typeof(int));
            disp.Columns.Add("現場手当", typeof(int));
            disp.Columns.Add("免許手当", typeof(int));
            disp.Columns.Add("離島手当", typeof(int)); //宿現場からリネーム！
            disp.Columns.Add("扶養手当", typeof(int));
            disp.Columns.Add("転勤手当", typeof(int)); 
            disp.Columns.Add("通勤非課税", typeof(int));
            disp.Columns.Add("通勤課税", typeof(int));
            disp.Columns.Add("退職積立金", typeof(int));
            disp.Columns.Add("持株奨励金", typeof(int));
            disp.Columns.Add("延長手当", typeof(int));
            disp.Columns.Add("法休出手当", typeof(int));
            disp.Columns.Add("所休出手当", typeof(int));
            disp.Columns.Add("残業手当", typeof(int));
            disp.Columns.Add("60超残手当", typeof(int));
            disp.Columns.Add("深夜手当", typeof(int));
            disp.Columns.Add("回数手当１", typeof(int));
            disp.Columns.Add("回数手当２", typeof(int));
            disp.Columns.Add("臨時手当", typeof(int));
            disp.Columns.Add("臨作業手当", typeof(int));
            disp.Columns.Add("正月期末", typeof(int));
            disp.Columns.Add("前払金(+)", typeof(int));
            disp.Columns.Add("臨休業手当", typeof(int));
            disp.Columns.Add("欠勤控除", typeof(int));
            disp.Columns.Add("支給合計額", typeof(int));
            disp.Columns.Add("健保", typeof(int));
            disp.Columns.Add("介保", typeof(int));
            disp.Columns.Add("厚年", typeof(int));
            disp.Columns.Add("雇保", typeof(int));
            disp.Columns.Add("所得税", typeof(int));
            disp.Columns.Add("住民税", typeof(int));
            disp.Columns.Add("財形積立", typeof(int));
            disp.Columns.Add("生命保険", typeof(int));
            disp.Columns.Add("友の会", typeof(int));
            disp.Columns.Add("固定他１", typeof(int));
            disp.Columns.Add("固定他２", typeof(int));
            disp.Columns.Add("積立金", typeof(int));
            disp.Columns.Add("前払金(-)", typeof(int));
            disp.Columns.Add("変動他１", typeof(int));
            disp.Columns.Add("変動他２", typeof(int));
            disp.Columns.Add("差押金", typeof(int));
            disp.Columns.Add("年調過不足額", typeof(int));
            disp.Columns.Add("控除合計額", typeof(int));
            disp.Columns.Add("延長時間", typeof(decimal));
            disp.Columns.Add("法休時間", typeof(decimal));
            disp.Columns.Add("所休時間", typeof(decimal));
            disp.Columns.Add("残業時間", typeof(decimal));
            disp.Columns.Add("60超残Ｈ", typeof(decimal));
            disp.Columns.Add("深夜時間", typeof(decimal));
            disp.Columns.Add("遅刻回数", typeof(int));
            disp.Columns.Add("遅刻時間", typeof(decimal));
            disp.Columns.Add("時給", typeof(decimal));
            disp.Columns.Add("所定", typeof(decimal));
            disp.Columns.Add("法休", typeof(decimal));
            disp.Columns.Add("所休", typeof(decimal));
            disp.Columns.Add("有給", typeof(decimal));
            disp.Columns.Add("特休", typeof(decimal));
            disp.Columns.Add("無特", typeof(decimal));
            disp.Columns.Add("振休", typeof(decimal));
            if (Convert.ToInt32(comboBox1.SelectedItem.ToString() + comboBox2.SelectedItem.ToString()) >= 201902)
            {
                disp.Columns.Add("休日", typeof(decimal));
            }
            else
            {
                disp.Columns.Add("公休", typeof(decimal));
                disp.Columns.Add("調休", typeof(decimal));
            }
            disp.Columns.Add("届欠", typeof(decimal));
            disp.Columns.Add("無届", typeof(decimal));
            disp.Columns.Add("回数１", typeof(int));
            disp.Columns.Add("回数２", typeof(int));
            disp.Columns.Add("通勤1日単価", typeof(int));
            disp.Columns.Add("標準報酬月額", typeof(int));
            disp.Columns.Add("有給残日数", typeof(decimal));
            disp.Columns.Add("振込口座額", typeof(int));
            disp.Columns.Add("現金支給額", typeof(int));
            disp.Columns.Add("差引支給額", typeof(int));
            disp.Columns.Add("備考", typeof(string));

            //TODO 最後に追加してみる
            disp.Columns.Add("登録手当", typeof(int));
            disp.Columns.Add("通信手当", typeof(int));
            disp.Columns.Add("車両手当", typeof(int));

            disp.Columns.Add("kintone情報", typeof(string));

            foreach (DataRow row in dt.Rows)
            {
                DataRow dr = disp.NewRow();
                dr["社員番号"] = row["社員番号"];
                dr["氏名"] = row["氏名"];
                dr["組織名"] = row["組織名"];
                dr["現場名"] = row["現場名"];
                dr["発令"] = row["発令名称"];

                    if (Convert.ToDateTime(row["入社年月日"]).ToString("yyyyMM") == ymex.ToString("yyyyMM"))
                    {
                        dr["基本情報変更"] = "※初給";
                    }
                    else if (row["本給1"].Equals(DBNull.Value))
                    {
                        dr["基本情報変更"] = "先月給与無";
                    }
                    else
                    {

                        if (Convert.ToInt32(row["職務技能給"]) - Convert.ToInt32(row["職務技能給1"]) != 0 |
                            Convert.ToInt32(row["調整手当"]) - Convert.ToInt32(row["調整手当1"]) != 0 |
                            Convert.ToInt32(row["特別手当"]) - Convert.ToInt32(row["特別手当1"]) != 0 |
                            Convert.ToInt32(row["皆勤手当"]) - Convert.ToInt32(row["皆勤手当1"]) != 0 |
                            Convert.ToInt32(row["役職手当"]) - Convert.ToInt32(row["役職手当1"]) != 0 |
                            Convert.ToInt32(row["現場手当"]) - Convert.ToInt32(row["現場手当1"]) != 0 |
                            Convert.ToInt32(row["免許手当"]) - Convert.ToInt32(row["免許手当1"]) != 0 |
                            Convert.ToInt32(row["離島手当"]) - Convert.ToInt32(row["離島手当1"]) != 0 |
                            Convert.ToInt32(row["扶養手当"]) - Convert.ToInt32(row["扶養手当1"]) != 0 |
                            Convert.ToInt32(row["転勤手当"]) - Convert.ToInt32(row["転勤手当1"]) != 0 |
                            Convert.ToInt32(row["通勤非課税"]) - Convert.ToInt32(row["通勤非課税1"]) != 0 |
                            Convert.ToInt32(row["通勤課税"]) - Convert.ToInt32(row["通勤課税1"]) != 0 |
                            Convert.ToInt32(row["登録手当"]) - Convert.ToInt32(row["登録手当1"]) != 0 |
                            Convert.ToInt32(row["通信手当"]) - Convert.ToInt32(row["通信手当1"]) != 0 |
                            Convert.ToInt32(row["車両手当"]) - Convert.ToInt32(row["車両手当1"]) != 0 |
                            Convert.ToInt32(row["退職積立金"]) - Convert.ToInt32(row["退職積立金1"]) != 0 |
                            Convert.ToInt32(row["持株奨励金"]) - Convert.ToInt32(row["持株奨励金1"]) != 0
                            )
                        {
                            dr["固定給変更"] = "有";

                        }
                        else
                        {
                            dr["固定給変更"] = "";
                        }

                        if (System.Math.Abs(Convert.ToInt32(row["差引支給額"]) - Convert.ToInt32(row["差引支給額1"])) > Convert.ToInt32(numericUpDown1.Value))
                        {
                            dr["支給差額"] += "差引支給額" + Convert.ToInt32(numericUpDown1.Value).ToString() + "円以上変化";
                        }

                        if (row["氏名"].ToString() != row["氏名1"].ToString() |
                            row["地区名"].ToString() != row["地区名1"].ToString() |
                            row["組織名"].ToString() != row["組織名1"].ToString() |
                            row["現場名"].ToString() != row["現場名1"].ToString() |
                            row["役職名"].ToString() != row["役職名1"].ToString() |
                            row["支給区分"].ToString() != row["支給区分1"].ToString() |
                            row["週労働数"].ToString() != row["週労働数1"].ToString() |
                            row["勤務時間"].ToString() != row["勤務時間1"].ToString()
                            )
                        {
                            dr["基本情報変更"] += "有";
                        }
                    }

                    dr["本給"] = Convert.ToInt32(row["本給"]);
                    dr["職務技能給"] = Convert.ToInt32(row["職務技能給"]);
                    dr["調整手当"] = Convert.ToInt32(row["調整手当"]);
                    dr["特別手当"] = Convert.ToInt32(row["特別手当"]);
                    dr["皆勤手当"] = Convert.ToInt32(row["皆勤手当"]);
                    dr["役職手当"] = Convert.ToInt32(row["役職手当"]);
                    dr["現場手当"] = Convert.ToInt32(row["現場手当"]);
                    dr["免許手当"] = Convert.ToInt32(row["免許手当"]);
                    dr["離島手当"] = Convert.ToInt32(row["離島手当"]);
                    dr["扶養手当"] = Convert.ToInt32(row["扶養手当"]);
                    dr["転勤手当"] = Convert.ToInt32(row["転勤手当"]);
                    dr["通勤非課税"] = Convert.ToInt32(row["通勤非課税"]);
                    dr["通勤課税"] = Convert.ToInt32(row["通勤課税"]);
                    dr["登録手当"] = Convert.ToInt32(row["登録手当"]);
                    dr["通信手当"] = Convert.ToInt32(row["通信手当"]);
                    dr["車両手当"] = Convert.ToInt32(row["車両手当"]);
                    dr["退職積立金"] = Convert.ToInt32(row["退職積立金"]);
                    dr["延長手当"] = Convert.ToInt32(row["延長手当"]);
                    dr["法休出手当"] = Convert.ToInt32(row["法休出手当"]);
                    dr["所休出手当"] = Convert.ToInt32(row["所休出手当"]);
                    dr["残業手当"] = Convert.ToInt32(row["残業手当"]);
                    dr["60超残手当"] = Convert.ToInt32(row["60超残手当"]);
                    dr["深夜手当"] = Convert.ToInt32(row["深夜手当"]);
                    dr["回数手当１"] = Convert.ToInt32(row["回数手当１"]);
                    dr["回数手当２"] = Convert.ToInt32(row["回数手当２"]);
                    dr["臨時手当"] = Convert.ToInt32(row["臨時手当"]);
                    dr["臨作業手当"] = Convert.ToInt32(row["臨作業手当"]);
                    dr["正月期末"] = Convert.ToInt32(row["正月期末"]);
                    dr["前払金(+)"] = Convert.ToInt32(row["前払金(+)"]);
                    dr["臨休業手当"] = Convert.ToInt32(row["臨休業手当"]);
                    dr["欠勤控除"] = Convert.ToInt32(row["欠勤控除"]);
                    dr["支給合計額"] = Convert.ToInt32(row["支給合計額"]);
                    dr["健保"] = Convert.ToInt32(row["健保"]);
                    dr["介保"] = Convert.ToInt32(row["介保"]);
                    dr["厚年"] = Convert.ToInt32(row["厚年"]);
                    dr["雇保"] = Convert.ToInt32(row["雇保"]);
                    dr["所得税"] = Convert.ToInt32(row["所得税"]);
                    dr["住民税"] = Convert.ToInt32(row["住民税"]);
                    dr["財形積立"] = Convert.ToInt32(row["財形積立"]);
                    dr["生命保険"] = Convert.ToInt32(row["生命保険"]);
                    dr["友の会"] = Convert.ToInt32(row["友の会"]);
                    dr["固定他１"] = Convert.ToInt32(row["固定他１"]);
                    dr["固定他２"] = Convert.ToInt32(row["固定他２"]);
                    dr["積立金"] = Convert.ToInt32(row["積立金"]);
                    dr["前払金(-)"] = Convert.ToInt32(row["前払金(-)"]);
                    dr["変動他１"] = Convert.ToInt32(row["変動他１"]);
                    dr["変動他２"] = Convert.ToInt32(row["変動他２"]);
                    dr["差押金"] = Convert.ToInt32(row["差押金"]);
                    dr["年調過不足額"] = Convert.ToInt32(row["年調過不足額"]);
                    dr["控除合計額"] = Convert.ToInt32(row["控除合計額"]);
                    dr["延長時間"] = Convert.ToDecimal(row["延長時間"]);
                    dr["法休時間"] = Convert.ToDecimal(row["法休時間"]);
                    dr["所休時間"] = Convert.ToDecimal(row["所休時間"]);
                    dr["残業時間"] = Convert.ToDecimal(row["残業時間"]);
                    dr["60超残Ｈ"] = Convert.ToDecimal(row["60超残Ｈ"]);
                    dr["深夜時間"] = Convert.ToDecimal(row["深夜時間"]);
                    dr["遅刻回数"] = Convert.ToInt32(row["遅刻回数"]);
                    dr["遅刻時間"] = Convert.ToDecimal(row["遅刻時間"]);
                    dr["時給"] = Convert.ToDecimal(row["時給"]);
                    dr["所定"] = Convert.ToDecimal(row["所定"]);
                    dr["法休"] = Convert.ToDecimal(row["法休"]);
                    dr["所休"] = Convert.ToDecimal(row["所休"]);
                    dr["有給"] = Convert.ToDecimal(row["有給"]);
                    dr["特休"] = Convert.ToDecimal(row["特休"]);
                    dr["無特"] = Convert.ToDecimal(row["無特"]);
                    dr["振休"] = Convert.ToDecimal(row["振休"]);
                    if (Convert.ToInt32(comboBox1.SelectedItem.ToString() + comboBox2.SelectedItem.ToString()) >= 201902)
                    {
                        dr["休日"] = Convert.ToDecimal(row["公休"]) + Convert.ToDecimal(row["調休"]);
                    }
                    else
                    {
                        dr["公休"] = Convert.ToDecimal(row["公休"]);
                        dr["調休"] = Convert.ToDecimal(row["調休"]);
                    }
                    dr["届欠"] = Convert.ToDecimal(row["届欠"]);
                    dr["無届"] = Convert.ToDecimal(row["無届"]);
                    dr["回数１"] = Convert.ToInt32(row["回数１"]);
                    dr["回数２"] = Convert.ToInt32(row["回数２"]);
                    dr["通勤1日単価"] = Convert.ToInt32(row["通勤1日単価"]);
                    dr["標準報酬月額"] = Convert.ToInt32(row["標準報酬月額"]);
                    dr["有給残日数"] = Convert.ToDecimal(row["有給残日数"]);
                    dr["振込口座額"] = Convert.ToInt32(row["振込口座額"]);
                    dr["現金支給額"] = Convert.ToInt32(row["現金支給額"]);
                    dr["差引支給額"] = Convert.ToInt32(row["差引支給額"]);
                    dr["備考"] = row["お知らせ"];
                　　dr["kintone情報"] = row["kintone情報"];

                disp.Rows.Add(dr);

            }

            //TODO
            dataGridView1.DataSource = disp;

            //	0	社員番号	string
            //	1	氏名	string
            //	2	組織名	string
            //	3	現場名	string
            //	4	発令	string
            //	5	基本情報変更	string
            //	6	固定給変更	string
            //	7	支給差額	string
            //	8	本給	int
            //	9	職務技能給	int
            //	10	調整手当	int
            //	11	特別手当	int
            //	12	皆勤手当	int
            //	13	役職手当	int
            //	14	現場手当	int
            //	15	免許手当	int
            //	16	離島手当	int RENAME
            //	17	扶養手当	int
            //	18	転勤手当	int
            //	19	通勤非課税	int
            //	20	通勤課税	int
            //	21	登録手当	int
            //	22	登録手当	int　NEW　
            //	23	退職積立金	int
            //　24　持株奨励金　int　NEW
            //	24	延長手当	int
            //	25	法休出手当	int
            //	26	所休出手当	int
            //	27	残業手当	int
            //	28	60超残手当	int
            //	29	深夜手当	int
            //	30	回数手当１	int
            //	31	回数手当２	int
            //	32	臨時手当	int
            //	33	臨作業手当	int
            //	34	正月期末	int
            //	35	前払金(+)	int
            //	36	欠勤控除	int
            //	37	支給合計額	int
            //	38	健保	int
            //	39	介保	int
            //	40	厚年	int
            //	41	雇保	int
            //	42	所得税	int
            //	43	住民税	int
            //	44	財形積立	int
            //	45	生命保険	int
            //	46	友の会	int
            //	47	固定他１	int
            //	48	固定他２	int
            //	49	積立金	int
            //	50	前払金(-)	int
            //	51	変動他１	int
            //	52	変動他２	int
            //	53	差押金	int
            //	54	年調過不足額	int
            //	55	控除合計額	int
            //	56	延長時間	decimal
            //	57	法休時間	decimal
            //	58	所休時間	decimal
            //	59	残業時間	decimal
            //	60	60超残Ｈ	decimal
            //	61	深夜時間	decimal
            //	62	遅刻回数	decimal
            //	63	遅刻時間	decimal
            //	64	時給	decimal
            //	65	所定	decimal
            //	66	法休	int
            //	67	所休	int
            //	68	有給	decimal
            //	69	特休	int
            //	70	無特	int
            //	71	振休	int
            //	72	休日	int
            //	73	届欠	int
            //	74	無届	int
            //	75	回数１	int
            //	76	回数２	int
            //
            //	77	標準報酬月額	int
            //	78	有給残日数	decimal
            //	79	振込口座額	int
            //	80	現金支給額	int
            //	81	差引支給額	int
            //	82	備考	string

            //右寄左寄
            for (int i = 8; i < 81; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            //3桁区切りと少数点表示

            for (int i = 0; i < disp.Columns.Count; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Format = "#,##0.#";
            }
            //for (int i = 8; i < 55; i++)
            //{
            //    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,##0";
            //}

            //for (int i = 55; i < 65; i++)
            //{
            //    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,##0.0";
            //}

            //dataGridView1.Columns[65].DefaultCellStyle.Format = "#,##0";
            //dataGridView1.Columns[66].DefaultCellStyle.Format = "#,##0.0";
            //dataGridView1.Columns[67].DefaultCellStyle.Format = "#,##0.0";

            //for (int i = 68; i < 77; i++)
            //{
            //    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,##0";
            //}

            //dataGridView1.Columns[77].DefaultCellStyle.Format = "#,##0.0";
            //dataGridView1.Columns[78].DefaultCellStyle.Format = "#,##0";
            //dataGridView1.Columns[79].DefaultCellStyle.Format = "#,##0";
            //dataGridView1.Columns[80].DefaultCellStyle.Format = "#,##0";

            //dataGridView1.DataSource = dt;

            if (disp.Rows.Count == 0)
            {
                label3.Text = "該当データ無";
            }
            else
            {
                label3.Text = dt.Rows.Count.ToString() + "人";
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e) 
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            if (drv == null)
            {
                return;
            }
            selectnum = drv[0].ToString();
            SyousaiData(selectnum);
        }

        private void SyousaiData(string str)
        {
            //対象者のみに絞込
            DataRow[] targetDr = dt.Select("社員番号 = " + str.Substring(0, 8), "");

            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;

            textBox2.Text = "";

            if (dt.Rows.Count == 0) return;

            //固定控除
            //対象者のみに絞込
            DataRow[] ttDr_kouzyo = dtkouzyo.Select("社員番号 = " + str.Substring(0, 8), "");

            DataTable dispkotei = new DataTable();
            dispkotei.Columns.Add("項目");
            dispkotei.Columns.Add("内容");
            dispkotei.Columns.Add("金額");

            DataRow wr2;
            foreach (DataRow row in ttDr_kouzyo)
            {
                wr2 = dispkotei.NewRow();
                wr2["項目"] = row[1].ToString();
                wr2["内容"] = row[2].ToString();
                wr2["金額"] = Convert.ToDecimal(row[3].ToString());

                dispkotei.Rows.Add(wr2);
            }

            dataGridView4.DataSource = dispkotei;





            DataTable kihondt = new DataTable();
            kihondt.Columns.Add("0列");
            kihondt.Columns.Add("1列");
            kihondt.Columns.Add("2列");
            kihondt.Columns.Add("3列");
            kihondt.Columns.Add("4列");
            kihondt.Columns.Add("5列");

            DataTable dispdt = new DataTable();

            dispdt.Columns.Add("0列");
            dispdt.Columns.Add("1列");
            dispdt.Columns.Add("2列");
            dispdt.Columns.Add("3列");
            dispdt.Columns.Add("4列");
            dispdt.Columns.Add("5列");
            dispdt.Columns.Add("6列");
            dispdt.Columns.Add("7列");
            dispdt.Columns.Add("8列");
            dispdt.Columns.Add("9列");

            DataRow wr;

            foreach (DataRow row in targetDr)
            {
                //年月
                textBox3.Text = row[213].ToString() + "年" + row[214].ToString() + "月給与";

                //基本情報1
                wr = kihondt.NewRow();
                wr["0列"] = row[0].Equals(DBNull.Value) ? "" : "社員番号"; //94
                wr["1列"] = row[1].Equals(DBNull.Value) ? "" : "氏名";
                wr["2列"] = row[2].Equals(DBNull.Value) ? "" : "カナ名";
                wr["3列"] = row[3].Equals(DBNull.Value) ? "" : "地区名";
                wr["4列"] = row[4].Equals(DBNull.Value) ? "" : "組織名";
                wr["5列"] = row[5].Equals(DBNull.Value) ? "" : "現場名";
                kihondt.Rows.Add(wr);

                wr = kihondt.NewRow();
                wr["0列"] = row[0].Equals(DBNull.Value) ? "" : row[0].ToString();
                wr["1列"] = row[1].Equals(DBNull.Value) ? "" : row[1].ToString();
                wr["2列"] = row[2].Equals(DBNull.Value) ? "" : row[2].ToString();
                wr["3列"] = row[3].Equals(DBNull.Value) ? "" : row[3].ToString();
                wr["4列"] = row[4].Equals(DBNull.Value) ? "" : row[4].ToString();
                wr["5列"] = row[5].Equals(DBNull.Value) ? "" : row[5].ToString();

                wr["0列"] += row[0].Equals(DBNull.Value) | row[97].Equals(DBNull.Value) | row[0].Equals(row[97]) ? "" : "⇐" + row[97].ToString();
                wr["1列"] += row[1].Equals(DBNull.Value) | row[98].Equals(DBNull.Value) | row[1].Equals(row[98]) ? "" : "⇐" + row[98].ToString();
                wr["2列"] += row[2].Equals(DBNull.Value) | row[99].Equals(DBNull.Value) | row[2].Equals(row[99]) ? "" : "⇐" + row[99].ToString();
                wr["3列"] += row[3].Equals(DBNull.Value) | row[100].Equals(DBNull.Value) | row[3].Equals(row[100]) ? "" : "⇐" + row[100].ToString();
                wr["4列"] += row[4].Equals(DBNull.Value) | row[101].Equals(DBNull.Value) | row[4].Equals(row[101]) ? "" : "⇐" + row[101].ToString();
                wr["5列"] += row[5].Equals(DBNull.Value) | row[102].Equals(DBNull.Value) | row[5].Equals(row[102]) ? "" : "⇐" + row[102].ToString();

                kihondt.Rows.Add(wr);

                //基本情報2
                wr = kihondt.NewRow();
                wr["0列"] = row[6].Equals(DBNull.Value) ? "" : "役職名";
                wr["1列"] = row[7].Equals(DBNull.Value) ? "" : "入社年月日";
                wr["2列"] = row[8].Equals(DBNull.Value) ? "" : "退職年月日";
                wr["3列"] = row[9].Equals(DBNull.Value) ? "" : "支給区分";
                wr["4列"] = row[10].Equals(DBNull.Value) ? "" : "週労働数";
                wr["5列"] = row[11].Equals(DBNull.Value) ? "" : "勤務時間";
                kihondt.Rows.Add(wr);

                wr = kihondt.NewRow();
                wr["0列"] = row[6].Equals(DBNull.Value) ? "" : row[6].ToString();
                wr["1列"] = row[7].Equals(DBNull.Value) ? "" : row[7].ToString();
                wr["2列"] = row[8].Equals(DBNull.Value) ? "" : row[8].ToString();
                wr["3列"] = row[9].Equals(DBNull.Value) ? "" : row[9].ToString();
                wr["4列"] = row[10].Equals(DBNull.Value) ? "" : row[10].ToString();
                wr["5列"] = row[11].Equals(DBNull.Value) ? "" : row[11].ToString().Replace(".000", "");

                wr["0列"] += row[6].Equals(DBNull.Value) | row[103].Equals(DBNull.Value) | row[6].Equals(row[103]) ? "" : "⇐" + row[103].ToString();
                wr["1列"] += row[7].Equals(DBNull.Value) | row[104].Equals(DBNull.Value) | row[7].Equals(row[104]) ? "" : "⇐" + row[104].ToString();
                wr["2列"] += row[8].Equals(DBNull.Value) | row[105].Equals(DBNull.Value) | row[8].Equals(row[105]) ? "" : "⇐" + row[105].ToString();
                wr["3列"] += row[9].Equals(DBNull.Value) | row[106].Equals(DBNull.Value) | row[9].Equals(row[106]) ? "" : "⇐" + row[106].ToString();
                wr["4列"] += row[10].Equals(DBNull.Value) | row[107].Equals(DBNull.Value) | row[10].Equals(row[107]) ? "" : "⇐" + row[107].ToString();
                wr["5列"] += row[11].Equals(DBNull.Value) | row[108].Equals(DBNull.Value) | row[11].Equals(row[108]) ? "" : "⇐" + row[108].ToString().Replace(".000", "");

                kihondt.Rows.Add(wr);


                int i0 = 0;
                int i1 = 0;
                int i2 = 0;
                int i3 = 0;
                int i4 = 0;
                int i5 = 0;
                int i6 = 0;
                int i7 = 0;
                int i8 = 0;
                int i9 = 0;

                //1行目
                i0 = row[109].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[12]) - Convert.ToInt32(row[109]));
                i1 = row[110].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[13]) - Convert.ToInt32(row[110]));
                i2 = row[111].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[14]) - Convert.ToInt32(row[111]));
                i3 = row[112].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[15]) - Convert.ToInt32(row[112]));
                i4 = row[113].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[16]) - Convert.ToInt32(row[113]));
                i5 = row[114].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[17]) - Convert.ToInt32(row[114]));
                i6 = row[115].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[18]) - Convert.ToInt32(row[115]));
                i7 = row[116].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[19]) - Convert.ToInt32(row[116]));
                i8 = row[117].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[20]) - Convert.ToInt32(row[117]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[12]) == 0 & i0 == 0 ? "" : "本給";
                wr["1列"] = Convert.ToInt32(row[13]) == 0 & i1 == 0 ? "" : "職務技能給";
                wr["2列"] = Convert.ToInt32(row[14]) == 0 & i2 == 0 ? "" : "調整手当";
                wr["3列"] = Convert.ToInt32(row[15]) == 0 & i3 == 0 ? "" : "特別手当";
                wr["4列"] = Convert.ToInt32(row[16]) == 0 & i4 == 0 ? "" : "皆勤手当";
                wr["5列"] = Convert.ToInt32(row[17]) == 0 & i5 == 0 ? "" : "役職手当";
                wr["6列"] = Convert.ToInt32(row[18]) == 0 & i6 == 0 ? "" : "現場手当";
                wr["7列"] = Convert.ToInt32(row[19]) == 0 & i7 == 0 ? "" : "免許手当";
                wr["8列"] = Convert.ToInt32(row[20]) == 0 & i8 == 0 ? "" : "離島手当";
                wr["9列"] = "";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[12]) == 0 ? "" : Convert.ToInt32(row[12]).ToString();
                wr["1列"] = Convert.ToInt32(row[13]) == 0 ? "" : Convert.ToInt32(row[13]).ToString();
                wr["2列"] = Convert.ToInt32(row[14]) == 0 ? "" : Convert.ToInt32(row[14]).ToString();
                wr["3列"] = Convert.ToInt32(row[15]) == 0 ? "" : Convert.ToInt32(row[15]).ToString();
                wr["4列"] = Convert.ToInt32(row[16]) == 0 ? "" : Convert.ToInt32(row[16]).ToString();
                wr["5列"] = Convert.ToInt32(row[17]) == 0 ? "" : Convert.ToInt32(row[17]).ToString();
                wr["6列"] = Convert.ToInt32(row[18]) == 0 ? "" : Convert.ToInt32(row[18]).ToString();
                wr["7列"] = Convert.ToInt32(row[19]) == 0 ? "" : Convert.ToInt32(row[19]).ToString();
                wr["8列"] = Convert.ToInt32(row[20]) == 0 ? "" : Convert.ToInt32(row[20]).ToString();
                wr["9列"] = "";
                dispdt.Rows.Add(wr);

                //2行目
                i0 = row[118].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[21]) - Convert.ToInt32(row[118]));
                i1 = row[119].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[22]) - Convert.ToInt32(row[119]));
                i2 = row[120].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[23]) - Convert.ToInt32(row[120]));
                i3 = row[121].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[24]) - Convert.ToInt32(row[121]));
                i4 = row[122].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[25]) - Convert.ToInt32(row[122]));
                i5 = row[123].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[26]) - Convert.ToInt32(row[123]));
                i6 = row[124].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[27]) - Convert.ToInt32(row[124]));

                i8 = row[125].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[28]) - Convert.ToInt32(row[125]));
                i9 = row[126].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[29]) - Convert.ToInt32(row[126]));


                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[21]) == 0 & i0 == 0 ? "" : "扶養手当";
                wr["1列"] = Convert.ToInt32(row[22]) == 0 & i1 == 0 ? "" : "転勤手当";
                wr["2列"] = Convert.ToInt32(row[23]) == 0 & i2 == 0 ? "" : "通勤非課税";
                wr["3列"] = Convert.ToInt32(row[24]) == 0 & i3 == 0 ? "" : "通勤課税";
                wr["4列"] = Convert.ToInt32(row[25]) == 0 & i4 == 0 ? "" : "登録手当";
                wr["5列"] = Convert.ToInt32(row[26]) == 0 & i5 == 0 ? "" : "通信手当";
                wr["6列"] = Convert.ToInt32(row[27]) == 0 & i6 == 0 ? "" : "車両手当";
                wr["7列"] = "";
                wr["8列"] = Convert.ToInt32(row[28]) == 0 & i8 == 0 ? "" : "退職積立金";
                wr["9列"] = Convert.ToInt32(row[29]) == 0 & i9 == 0 ? "" : "持株奨励金"; ;

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += i9 > 0 ? "　(+" + i9.ToString() + ")" : i9 < 0 ? "　(" + i9.ToString() + ")" : "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[21]) == 0 ? "" : Convert.ToInt32(row[21]).ToString();
                wr["1列"] = Convert.ToInt32(row[22]) == 0 ? "" : Convert.ToInt32(row[22]).ToString();
                wr["2列"] = Convert.ToInt32(row[23]) == 0 ? "" : Convert.ToInt32(row[23]).ToString();
                wr["3列"] = Convert.ToInt32(row[24]) == 0 ? "" : Convert.ToInt32(row[24]).ToString();
                wr["4列"] = Convert.ToInt32(row[25]) == 0 ? "" : Convert.ToInt32(row[25]).ToString();
                wr["5列"] = Convert.ToInt32(row[26]) == 0 ? "" : Convert.ToInt32(row[26]).ToString();
                wr["6列"] = Convert.ToInt32(row[27]) == 0 ? "" : Convert.ToInt32(row[27]).ToString();
                wr["7列"] = "";
                wr["8列"] = Convert.ToInt32(row[28]) == 0 ? "" : Convert.ToInt32(row[28]).ToString();
                wr["9列"] = Convert.ToInt32(row[29]) == 0 ? "" : Convert.ToInt32(row[29]).ToString(); 
                dispdt.Rows.Add(wr);

                //3行目
                i0 = row[127].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[30]) - Convert.ToInt32(row[127]));
                i1 = row[128].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[31]) - Convert.ToInt32(row[128]));
                i2 = row[129].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[32]) - Convert.ToInt32(row[129]));
                i3 = row[130].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[33]) - Convert.ToInt32(row[130]));
                i4 = row[131].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[34]) - Convert.ToInt32(row[131]));
                i5 = row[132].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[35]) - Convert.ToInt32(row[132]));
                i6 = row[133].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[36]) - Convert.ToInt32(row[133]));
                i7 = row[134].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[37]) - Convert.ToInt32(row[134]));
                i8 = row[135].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[38]) - Convert.ToInt32(row[135]));


                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[30]) == 0 & i0 == 0 ? "" : "延長手当";
                wr["1列"] = Convert.ToInt32(row[31]) == 0 & i1 == 0 ? "" : "法休出手当";
                wr["2列"] = Convert.ToInt32(row[32]) == 0 & i2 == 0 ? "" : "所休出手当";
                wr["3列"] = Convert.ToInt32(row[33]) == 0 & i3 == 0 ? "" : "残業手当";
                wr["4列"] = Convert.ToInt32(row[34]) == 0 & i4 == 0 ? "" : "60超残手当";
                wr["5列"] = Convert.ToInt32(row[35]) == 0 & i5 == 0 ? "" : "深夜手当";
                wr["6列"] = Convert.ToInt32(row[36]) == 0 & i6 == 0 ? "" : "回数手当１";
                wr["7列"] = Convert.ToInt32(row[37]) == 0 & i7 == 0 ? "" : "回数手当２";
                wr["8列"] = Convert.ToInt32(row[38]) == 0 & i8 == 0 ? "" : "臨時手当";
                wr["9列"] = "";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[30]) == 0 ? "" : Convert.ToInt32(row[30]).ToString();
                wr["1列"] = Convert.ToInt32(row[31]) == 0 ? "" : Convert.ToInt32(row[31]).ToString();
                wr["2列"] = Convert.ToInt32(row[32]) == 0 ? "" : Convert.ToInt32(row[32]).ToString();
                wr["3列"] = Convert.ToInt32(row[33]) == 0 ? "" : Convert.ToInt32(row[33]).ToString();
                wr["4列"] = Convert.ToInt32(row[34]) == 0 ? "" : Convert.ToInt32(row[34]).ToString();
                wr["5列"] = Convert.ToInt32(row[35]) == 0 ? "" : Convert.ToInt32(row[35]).ToString();
                wr["6列"] = Convert.ToInt32(row[36]) == 0 ? "" : Convert.ToInt32(row[36]).ToString();
                wr["7列"] = Convert.ToInt32(row[37]) == 0 ? "" : Convert.ToInt32(row[37]).ToString();
                wr["8列"] = Convert.ToInt32(row[38]) == 0 ? "" : Convert.ToInt32(row[38]).ToString();
                wr["9列"] = "";
                dispdt.Rows.Add(wr);

                //4行目
                i0 = row[136].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[39]) - Convert.ToInt32(row[136]));
                i1 = row[137].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[40]) - Convert.ToInt32(row[137]));
                i2 = row[138].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[41]) - Convert.ToInt32(row[138]));
                i3 = row[139].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[42]) - Convert.ToInt32(row[139]));




                i8 = row[140].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[43]) - Convert.ToInt32(row[140]));
                i9 = row[141].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[44]) - Convert.ToInt32(row[141]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[39]) == 0 & i0 == 0 ? "" : "臨作業手当";
                wr["1列"] = Convert.ToInt32(row[40]) == 0 & i1 == 0 ? "" : "正月期末";
                wr["2列"] = Convert.ToInt32(row[41]) == 0 & i2 == 0 ? "" : "前払金(+)";
                wr["3列"] = Convert.ToInt32(row[42]) == 0 & i2 == 0 ? "" : "臨休業手当";
                wr["4列"] = "";
                wr["5列"] = "";
                wr["6列"] = "";
                wr["7列"] = "";
                wr["8列"] = Convert.ToInt32(row[43]) == 0 & i8 == 0 ? "" : "欠勤控除";
                wr["9列"] = Convert.ToInt32(row[44]) == 0 & i9 == 0 ? "" : "支給合計額";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += "";
                    wr["5列"] += "";
                    wr["6列"] += "";
                    wr["7列"] += "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += i9 > 0 ? "　(+" + i9.ToString() + ")" : i9 < 0 ? "　(" + i9.ToString() + ")" : "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[39]) == 0 ? "" : Convert.ToInt32(row[39]).ToString();
                wr["1列"] = Convert.ToInt32(row[40]) == 0 ? "" : Convert.ToInt32(row[40]).ToString();
                wr["2列"] = Convert.ToInt32(row[41]) == 0 ? "" : Convert.ToInt32(row[41]).ToString();
                wr["3列"] = Convert.ToInt32(row[42]) == 0 ? "" : Convert.ToInt32(row[42]).ToString(); 
                wr["4列"] = "";
                wr["5列"] = "";
                wr["6列"] = "";
                wr["7列"] = "";
                wr["8列"] = Convert.ToInt32(row[43]) == 0 ? "" : Convert.ToInt32(row[43]).ToString();
                wr["9列"] = Convert.ToInt32(row[44]) == 0 ? "" : Convert.ToInt32(row[44]).ToString();
                dispdt.Rows.Add(wr);

                //5行目
                i0 = row[142].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[45]) - Convert.ToInt32(row[142]));
                i1 = row[143].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[46]) - Convert.ToInt32(row[143]));
                i2 = row[144].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[47]) - Convert.ToInt32(row[144]));
                i3 = row[145].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[48]) - Convert.ToInt32(row[145]));
                i4 = row[146].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[49]) - Convert.ToInt32(row[146]));
                i5 = row[147].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[50]) - Convert.ToInt32(row[147]));
                i6 = row[148].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[51]) - Convert.ToInt32(row[148]));
                i7 = row[149].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[52]) - Convert.ToInt32(row[149]));
                i8 = row[150].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[53]) - Convert.ToInt32(row[150]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[45]) == 0 & i0 == 0 ? "" :  "健保";
                wr["1列"] = Convert.ToInt32(row[46]) == 0 & i1 == 0 ? "" :  "介保";
                wr["2列"] = Convert.ToInt32(row[47]) == 0 & i2 == 0 ? "" :  "厚年";
                wr["3列"] = Convert.ToInt32(row[48]) == 0 & i3 == 0 ? "" :  "雇保";
                wr["4列"] = Convert.ToInt32(row[49]) == 0 & i4 == 0 ? "" :  "所得税";
                wr["5列"] = Convert.ToInt32(row[50]) == 0 & i5 == 0 ? "" :  "住民税";
                wr["6列"] = Convert.ToInt32(row[51]) == 0 & i6 == 0 ? "" :  "財形積立";
                wr["7列"] = Convert.ToInt32(row[52]) == 0 & i7 == 0 ? "" :  "生命保険";
                wr["8列"] = Convert.ToInt32(row[53]) == 0 & i8 == 0 ? "" :  "友の会";
                wr["9列"] = "";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[45]) == 0 ? "" : Convert.ToInt32(row[45]).ToString();
                wr["1列"] = Convert.ToInt32(row[46]) == 0 ? "" : Convert.ToInt32(row[46]).ToString();
                wr["2列"] = Convert.ToInt32(row[47]) == 0 ? "" : Convert.ToInt32(row[47]).ToString();
                wr["3列"] = Convert.ToInt32(row[48]) == 0 ? "" : Convert.ToInt32(row[48]).ToString();
                wr["4列"] = Convert.ToInt32(row[49]) == 0 ? "" : Convert.ToInt32(row[49]).ToString();
                wr["5列"] = Convert.ToInt32(row[50]) == 0 ? "" : Convert.ToInt32(row[50]).ToString();
                wr["6列"] = Convert.ToInt32(row[51]) == 0 ? "" : Convert.ToInt32(row[51]).ToString();
                wr["7列"] = Convert.ToInt32(row[52]) == 0 ? "" : Convert.ToInt32(row[52]).ToString();
                wr["8列"] = Convert.ToInt32(row[53]) == 0 ? "" : Convert.ToInt32(row[53]).ToString();
                wr["9列"] = "";
                dispdt.Rows.Add(wr);

                //6行目
                i0 = row[151].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[54]) - Convert.ToInt32(row[151]));
                i1 = row[152].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[55]) - Convert.ToInt32(row[152]));
                i2 = row[153].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[56]) - Convert.ToInt32(row[153]));
                i3 = row[154].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[57]) - Convert.ToInt32(row[154]));
                i4 = row[155].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[58]) - Convert.ToInt32(row[155]));
                i5 = row[156].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[59]) - Convert.ToInt32(row[156]));
                i6 = row[157].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[60]) - Convert.ToInt32(row[157]));
                i7 = row[158].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[61]) - Convert.ToInt32(row[158]));
                //i8 = row[159].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[62]) - Convert.ToInt32(row[159]));
                i9 = row[160].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[63]) - Convert.ToInt32(row[160]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[54]) == 0 & i0 == 0 ? "" : "固定他１";
                wr["1列"] = Convert.ToInt32(row[55]) == 0 & i1 == 0 ? "" : "固定他２";
                wr["2列"] = Convert.ToInt32(row[56]) == 0 & i2 == 0 ? "" : "積立金";
                wr["3列"] = Convert.ToInt32(row[57]) == 0 & i3 == 0 ? "" : "前払金(-)";
                wr["4列"] = Convert.ToInt32(row[58]) == 0 & i4 == 0 ? "" : "変動他１";
                wr["5列"] = Convert.ToInt32(row[59]) == 0 & i5 == 0 ? "" : "変動他２";
                wr["6列"] = Convert.ToInt32(row[60]) == 0 & i6 == 0 ? "" : "差押金";
                wr["7列"] = Convert.ToInt32(row[61]) == 0 & i7 == 0 ? "" : "年調過不足額";
                wr["8列"] = ""; //Convert.ToInt32(row[62]) == 0 & i8 == 0 ? "" : "定額減税額";
                wr["9列"] = Convert.ToInt32(row[63]) == 0 & i9 == 0 ? "" : "控除合計額";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += ""; // i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += i9 > 0 ? "　(+" + i9.ToString() + ")" : i9 < 0 ? "　(" + i9.ToString() + ")" : "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[54]) == 0 ? "" : Convert.ToInt32(row[54]).ToString();
                wr["1列"] = Convert.ToInt32(row[55]) == 0 ? "" : Convert.ToInt32(row[55]).ToString();
                wr["2列"] = Convert.ToInt32(row[56]) == 0 ? "" : Convert.ToInt32(row[56]).ToString();
                wr["3列"] = Convert.ToInt32(row[57]) == 0 ? "" : Convert.ToInt32(row[57]).ToString();
                wr["4列"] = Convert.ToInt32(row[58]) == 0 ? "" : Convert.ToInt32(row[58]).ToString();
                wr["5列"] = Convert.ToInt32(row[59]) == 0 ? "" : Convert.ToInt32(row[59]).ToString();
                wr["6列"] = Convert.ToInt32(row[60]) == 0 ? "" : Convert.ToInt32(row[60]).ToString();
                wr["7列"] = Convert.ToInt32(row[61]) == 0 ? "" : Convert.ToInt32(row[61]).ToString();
                wr["8列"] = ""; // Convert.ToInt32(row[62]) == 0 ? "" : Convert.ToInt32(row[62]).ToString();
                wr["9列"] = Convert.ToInt32(row[63]) == 0 ? "" : Convert.ToInt32(row[63]).ToString();
                dispdt.Rows.Add(wr);

                //7行目
                i0 = row[161].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[64]) - Convert.ToInt32(row[161]));
                i1 = row[162].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[65]) - Convert.ToInt32(row[162]));
                i2 = row[163].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[66]) - Convert.ToInt32(row[163]));
                i3 = row[164].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[67]) - Convert.ToInt32(row[164]));
                i4 = row[165].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[68]) - Convert.ToInt32(row[165]));
                i5 = row[166].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[69]) - Convert.ToInt32(row[166]));
                i6 = row[167].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[60]) - Convert.ToInt32(row[167]));
                i7 = row[168].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[71]) - Convert.ToInt32(row[168]));
                i8 = row[169].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[72]) - Convert.ToInt32(row[169]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToDecimal(row[64]) == 0 & i0 == 0 ? "" : "延長時間";
                wr["1列"] = Convert.ToDecimal(row[65]) == 0 & i1 == 0 ? "" : "法休時間";
                wr["2列"] = Convert.ToDecimal(row[66]) == 0 & i2 == 0 ? "" : "所休時間";
                wr["3列"] = Convert.ToDecimal(row[67]) == 0 & i3 == 0 ? "" : "残業時間";
                wr["4列"] = Convert.ToDecimal(row[68]) == 0 & i4 == 0 ? "" : "60超残Ｈ";
                wr["5列"] = Convert.ToDecimal(row[69]) == 0 & i5 == 0 ? "" : "深夜時間";
                wr["6列"] = Convert.ToDecimal(row[60]) == 0 & i6 == 0 ? "" : "遅刻回数";
                wr["7列"] = Convert.ToDecimal(row[71]) == 0 & i7 == 0 ? "" : "遅刻時間";
                wr["8列"] = Convert.ToDecimal(row[72]) == 0 & i8 == 0 ? "" : "時給";
                wr["9列"] = "";

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToDecimal(row[64]) == 0 ? "" : Convert.ToDecimal(row[64]).ToString("#,##0.#");
                wr["1列"] = Convert.ToDecimal(row[65]) == 0 ? "" : Convert.ToDecimal(row[65]).ToString("#,##0.#");
                wr["2列"] = Convert.ToDecimal(row[66]) == 0 ? "" : Convert.ToDecimal(row[66]).ToString("#,##0.#");
                wr["3列"] = Convert.ToDecimal(row[67]) == 0 ? "" : Convert.ToDecimal(row[67]).ToString("#,##0.#");
                wr["4列"] = Convert.ToDecimal(row[68]) == 0 ? "" : Convert.ToDecimal(row[68]).ToString("#,##0.#");
                wr["5列"] = Convert.ToDecimal(row[69]) == 0 ? "" : Convert.ToDecimal(row[69]).ToString("#,##0.#");
                wr["6列"] = Convert.ToDecimal(row[60]) == 0 ? "" : Convert.ToDecimal(row[60]).ToString("#,##0.#");
                wr["7列"] = Convert.ToDecimal(row[71]) == 0 ? "" : Convert.ToDecimal(row[71]).ToString("#,##0.#");
                wr["8列"] = Convert.ToDecimal(row[72]) == 0 ? "" : Convert.ToDecimal(row[72]).ToString("#,##0.##");
                wr["9列"] = "";
                dispdt.Rows.Add(wr);

                //8行目
                i0 = row[170].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[73]) - Convert.ToInt32(row[170]));
                i1 = row[171].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[74]) - Convert.ToInt32(row[171]));
                i2 = row[172].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[75]) - Convert.ToInt32(row[172]));
                i3 = row[173].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[76]) - Convert.ToInt32(row[173]));
                i4 = row[174].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[77]) - Convert.ToInt32(row[174]));
                i5 = row[175].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[78]) - Convert.ToInt32(row[175]));
                i6 = row[176].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[79]) - Convert.ToInt32(row[176]));
                i7 = row[177].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[80]) - Convert.ToInt32(row[177]));
                //i8 = row[186].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[89]) - Convert.ToInt32(row[186])); //+9

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToDecimal(row[73]) == 0 & i0 == 0 ? "" : "所定";
                wr["1列"] = Convert.ToInt32(row[74]) == 0 & i1 == 0 ? "" : "法休";
                wr["2列"] = Convert.ToInt32(row[75]) == 0 & i2 == 0 ? "" : "所休";
                wr["3列"] = Convert.ToDecimal(row[76]) == 0 & i3 == 0 ? "" : "有給";
                wr["4列"] = Convert.ToDecimal(row[77]) == 0 & i4 == 0 ? "" : "特休";
                wr["5列"] = Convert.ToInt32(row[78]) == 0 & i5 == 0 ? "" : "無特";
                wr["6列"] = Convert.ToInt32(row[79]) == 0 & i6 == 0 ? "" : "振休";
                wr["7列"] = Convert.ToInt32(row[80]) == 0 & i7 == 0 ? "" : "公休";
                wr["8列"] = ""; // Convert.ToInt32(row[89]) == 0 & i8 == 0 ? "" : "振込口座額"; //+8
                wr["9列"] = "";
                //d

                //前月差を表示
                if (!checkBox1.Checked)
                {
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += ""; // i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToDecimal(row[73]) == 0 ? "" : Convert.ToDecimal(row[73]).ToString("#,##0.#");
                wr["1列"] = Convert.ToInt32(row[74]) == 0 ? "" : Convert.ToInt32(row[74]).ToString("#,##0.#");
                wr["2列"] = Convert.ToInt32(row[75]) == 0 ? "" : Convert.ToInt32(row[75]).ToString("#,##0.#");
                wr["3列"] = Convert.ToDecimal(row[76]) == 0 ? "" : Convert.ToDecimal(row[76]).ToString("#,##0.#");
                wr["4列"] = Convert.ToDecimal(row[77]) == 0 ? "" : Convert.ToDecimal(row[77]).ToString("#,##0.#");
                wr["5列"] = Convert.ToInt32(row[78]) == 0 ? "" : Convert.ToInt32(row[78]).ToString("#,##0.#");
                wr["6列"] = Convert.ToInt32(row[79]) == 0 ? "" : Convert.ToInt32(row[79]).ToString("#,##0.#");
                wr["7列"] = Convert.ToInt32(row[80]) == 0 ? "" : Convert.ToInt32(row[80]).ToString("#,##0.#");
                wr["8列"] = ""; // Convert.ToInt32(row[89]) == 0 ? "" : Convert.ToInt32(row[89]).ToString("#,##0.#");
                wr["9列"] = "";
                dispdt.Rows.Add(wr);

                //9行目
                //先月差額を格納
                i0 = row[178].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[81]) - Convert.ToInt32(row[178]));
                i1 = row[179].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[82]) - Convert.ToInt32(row[179]));
                i2 = row[180].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[83]) - Convert.ToInt32(row[180]));
                i3 = row[181].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[84]) - Convert.ToInt32(row[181]));
                i4 = row[182].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[85]) - Convert.ToInt32(row[182]));
                i5 = row[183].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[86]) - Convert.ToInt32(row[183]));
                i6 = row[184].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[87]) - Convert.ToInt32(row[184]));
                i7 = row[185].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[88]) - Convert.ToInt32(row[185]));
                i8 = row[159].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[62]) - Convert.ToInt32(row[159])); //-27
                i9 = row[188].Equals(DBNull.Value) ? 0 : (Convert.ToInt32(row[91]) - Convert.ToInt32(row[188]));

                //列名表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[81]) == 0 & i0 == 0 ? "" : "調休";
                wr["1列"] = Convert.ToInt32(row[82]) == 0 & i1 == 0 ? "" : "届欠";
                wr["2列"] = Convert.ToInt32(row[83]) == 0 & i2 == 0 ? "" : "無届";
                wr["3列"] = Convert.ToInt32(row[84]) == 0 & i3 == 0 ? "" : "回数１";
                wr["4列"] = Convert.ToInt32(row[85]) == 0 & i4 == 0 ? "" : "回数２";
                wr["5列"] = Convert.ToInt32(row[86]) == 0 & i5 == 0 ? "" : "通勤1日単価"; ;
                wr["6列"] = Convert.ToInt32(row[87]) == 0 & i6 == 0 ? "" : "標準報酬月額";
                wr["7列"] = Convert.ToDecimal(row[88]) == 0 & i7 == 0 ? "" : "有給残日数";
                wr["8列"] = Convert.ToInt32(row[62]) == 0 & i8 == 0 ? "" : "定額減税額"; //-27
                wr["9列"] = Convert.ToInt32(row[91]) == 0 & i9 == 0 ? "" : "差引支給額";

                //前月差を表示
                if (!checkBox1.Checked)
                { 
                    wr["0列"] += i0 > 0 ? "　(+" + i0.ToString() + ")" : i0 < 0 ? "　(" + i0.ToString() + ")" : "";
                    wr["1列"] += i1 > 0 ? "　(+" + i1.ToString() + ")" : i1 < 0 ? "　(" + i1.ToString() + ")" : "";
                    wr["2列"] += i2 > 0 ? "　(+" + i2.ToString() + ")" : i2 < 0 ? "　(" + i2.ToString() + ")" : "";
                    wr["3列"] += i3 > 0 ? "　(+" + i3.ToString() + ")" : i3 < 0 ? "　(" + i3.ToString() + ")" : "";
                    wr["4列"] += i4 > 0 ? "　(+" + i4.ToString() + ")" : i4 < 0 ? "　(" + i4.ToString() + ")" : "";
                    wr["5列"] += i5 > 0 ? "　(+" + i5.ToString() + ")" : i5 < 0 ? "　(" + i5.ToString() + ")" : "";
                    wr["6列"] += i6 > 0 ? "　(+" + i6.ToString() + ")" : i6 < 0 ? "　(" + i6.ToString() + ")" : "";
                    wr["7列"] += i7 > 0 ? "　(+" + i7.ToString() + ")" : i7 < 0 ? "　(" + i7.ToString() + ")" : "";
                    wr["8列"] += i8 > 0 ? "　(+" + i8.ToString() + ")" : i8 < 0 ? "　(" + i8.ToString() + ")" : "";
                    wr["9列"] += i9 > 0 ? "　(+" + i9.ToString() + ")" : i9 < 0 ? "　(" + i9.ToString() + ")" : "";
                }
                dispdt.Rows.Add(wr);

                //今月表示
                wr = dispdt.NewRow();
                wr["0列"] = Convert.ToInt32(row[81]) == 0 ? "" : Convert.ToInt32(row[81]).ToString();
                wr["1列"] = Convert.ToInt32(row[82]) == 0 ? "" : Convert.ToInt32(row[82]).ToString();
                wr["2列"] = Convert.ToInt32(row[83]) == 0 ? "" : Convert.ToInt32(row[83]).ToString();
                wr["3列"] = Convert.ToInt32(row[84]) == 0 ? "" : Convert.ToInt32(row[84]).ToString();
                wr["4列"] = Convert.ToInt32(row[85]) == 0 ? "" : Convert.ToInt32(row[85]).ToString();
                wr["5列"] = Convert.ToInt32(row[86]) == 0 ? "" : Convert.ToInt32(row[86]).ToString();
                wr["6列"] = Convert.ToInt32(row[87]) == 0 ? "" : Convert.ToInt32(row[87]).ToString();
                wr["7列"] = Convert.ToDecimal(row[88]) == 0 ? "" : Convert.ToDecimal(row[88]).ToString();
                wr["8列"] = Convert.ToInt32(row[62]) == 0 ? "" : Convert.ToInt32(row[62]).ToString(); //+1
                wr["9列"] = Convert.ToInt32(row[91]) == 0 ? "" : Convert.ToInt32(row[91]).ToString();   
                dispdt.Rows.Add(wr);

                textBox2.Text = row[221].Equals(DBNull.Value) ? "" : row[221].ToString(); //備考
                textBox4.Text = row[222].Equals(DBNull.Value) ? "" : row[222].ToString(); //kintone情報
            }

            dataGridView2.DataSource = dispdt;
            dataGridView3.DataSource = kihondt;

            //dataGridView2.Columns[0].DefaultCellStyle.Format = "#,0";
            if (dispdt.Rows.Count == 0) return;
            //基本情報
            //1行目
            dataGridView3.Rows[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView3.Rows[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //2行目
            dataGridView3.Rows[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView3.Rows[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView3.Rows[0].DefaultCellStyle.BackColor = Color.Olive;
            dataGridView3.Rows[2].DefaultCellStyle.BackColor = Color.Olive;

            dataGridView3.Columns[0].Width = 135;
            dataGridView3.Columns[1].Width = 200;
            dataGridView3.Columns[2].Width = 135;
            dataGridView3.Columns[3].Width = 135;
            dataGridView3.Columns[4].Width = 200;
            dataGridView3.Columns[5].Width = 350;


            //1行目
            dataGridView2.Rows[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //2行目
            dataGridView2.Rows[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //3行目
            dataGridView2.Rows[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //4行目
            dataGridView2.Rows[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //5行目
            dataGridView2.Rows[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //6行目
            dataGridView2.Rows[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //7行目
            dataGridView2.Rows[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //8行目
            dataGridView2.Rows[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //9行目
            dataGridView2.Rows[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Rows[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dataGridView2.Rows[0].DefaultCellStyle.BackColor = Color.Honeydew;
            dataGridView2.Rows[2].DefaultCellStyle.BackColor = Color.Honeydew;
            dataGridView2.Rows[4].DefaultCellStyle.BackColor = Color.Honeydew;
            dataGridView2.Rows[6].DefaultCellStyle.BackColor = Color.Honeydew;

            dataGridView2.Rows[8].DefaultCellStyle.BackColor = Color.Beige;
            dataGridView2.Rows[10].DefaultCellStyle.BackColor = Color.Beige;

            dataGridView2.Rows[12].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView2.Rows[14].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView2.Rows[16].DefaultCellStyle.BackColor = Color.LightBlue;

            if (checkBox1.Checked)
            {
                dataGridView2.Columns[0].Width = 80;
                dataGridView2.Columns[1].Width = 80;
                dataGridView2.Columns[2].Width = 80;
                dataGridView2.Columns[3].Width = 80;
                dataGridView2.Columns[4].Width = 80;
                dataGridView2.Columns[5].Width = 80;
                dataGridView2.Columns[6].Width = 80;
                dataGridView2.Columns[7].Width = 80;
                dataGridView2.Columns[8].Width = 80;
                dataGridView2.Columns[9].Width = 80;
            }
            else
            {
                dataGridView2.Columns[0].Width = 120;
                dataGridView2.Columns[1].Width = 120;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[3].Width = 110;
                dataGridView2.Columns[4].Width = 110;
                dataGridView2.Columns[5].Width = 110;
                dataGridView2.Columns[6].Width = 110;
                dataGridView2.Columns[7].Width = 110;
                dataGridView2.Columns[8].Width = 120;
                dataGridView2.Columns[9].Width = 130;
            }

            dataGridView4.Columns[0].Width = 70;
            dataGridView4.Columns[1].Width = 150;
            dataGridView4.Columns[2].Width = 60;
            dataGridView4.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[2].DefaultCellStyle.Format = "#,0";
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            if (drv == null)
            {
                //MessageBox.Show("test");
                return;
            }
            SyousaiData(drv[0].ToString());
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem == null) return;
            ym = new DateTime(Convert.ToInt16(comboBox1.SelectedItem), Convert.ToInt16(comboBox2.SelectedItem), 1);
            ymex = ym.AddMonths(-1);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null) return;
            ym = new DateTime(Convert.ToInt16(comboBox1.SelectedItem), Convert.ToInt16(comboBox2.SelectedItem), 1);
            ymex = ym.AddMonths(-1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str = comboBox1.SelectedItem.ToString() + comboBox2.SelectedItem.ToString();
            if (maxymd < Convert.ToInt32(str))
            {
                MessageBox.Show(maxymd + "を超えたデータはありません");
                return;
            }

            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetMeisai();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetMeisai();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //ボタン無効化・カーソル変更
                button1.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                GetMeisai();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }
    }
}
