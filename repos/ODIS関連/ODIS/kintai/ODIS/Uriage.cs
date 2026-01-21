using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Npgsql;
using Microsoft.VisualBasic;
using System.Data.SqlClient;

namespace ODIS.ODIS
{
    public partial class Uriage : Form
    {
        //private DataTable dt = new DataTable();

        public Uriage()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            cburiageari.Checked = true;

            //データ取得
            //GetUriageData();

            //初期値設定
            Clear();
        }

        private void Clear()
        {
            //文字列で絞込
            textBox1.Text = "";
            textBox2.Text = "";

            //一列表示
            checkBox1.Checked = false;

            //コード表示
            checkBox2.Checked = false;

            //日付設定
            DateTime today = DateTime.Today;
            if (today.Day >= 15)
            {
                dateTimePicker1.Value = today;
                dateTimePicker2.Value = today;
            }
            else
            {
                dateTimePicker1.Value = today.AddMonths(-1);
                dateTimePicker2.Value = today.AddMonths(-1);
            }


            //金額
            cbizyou.Checked = false;
            cbika.Checked = false;

            izyou.Value = 0;
            ika.Value = 0;

            izyou.Visible = false;
            ika.Visible = false;
            izyo.Visible = false;
            ik.Visible = false;
        }

        private void GetUriageData()
        {
            if (cburiageari.Checked == false && cburiagenashi.Checked == false)
            {
                MessageBox.Show("売上計上有無のチェックはどっちかは入れでください。");
                return;
            }

            //処理速度計算
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            //除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                foreach (string s in ar2)
                {
                    result += " and (reskey not like '%" + s + "%' and reskey not like '%" + Com.isOneByteChar(s) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' and reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' and reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' and reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //金額
            if (cbizyou.Checked)
            {
                result += " and 売上額 >= '" + izyou.Value + "' ";
            }

            if (cbika.Checked)
            {
                result += " and 売上額 <= '" + ika.Value + "'";
            }

            string sDate = dateTimePicker1.Value.ToString("yyyyMM");
            string eDate = dateTimePicker2.Value.ToString("yyyyMM");


            int nRet;

            NpgsqlConnection conn = new NpgsqlConnection(Com.POSConstr);
            conn.Open();
            DataTable dt = new DataTable();

            //string sql = "select * from kpcp01.\"GetUriageDataSearchAll\" where 売上年月 between '" + sDate + "' and '" + eDate + "'" + result;
            string sql = "select * from kpcp01.\"CostomGyoumuKanri_serch\" where 売上年月 between '" + sDate + "' and '" + eDate + "'" + result;

            if (cburiageari.Checked)
            {
                if (cburiagenashi.Checked)
                {
                    //どっちもチェック
                }
                else
                {
                    //無はいらない
                    sql += " and 売上チェック = '1'";
                }
            }
            else
            {
                if (cburiagenashi.Checked)
                {
                    //有はいらない。無はいる
                    sql += " and 売上チェック <> '1'";

                }
                else
                {
                    //ここまでこない。
                }
            }
            

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(sql, conn);
            nRet = adapter.Fill(dt);

            conn.Close();

            //解放
            adapter.Dispose();
            conn.Dispose();


            DataTable Disp = new DataTable();

            decimal uriageAll = 0;

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //列表示
            if (checkBox1.Checked)
            {
                //コード表示
                if (checkBox2.Checked)
                {
                    //一列表示
                    //コード表示
                    //0
                    Disp.Columns.Add("売上年月", typeof(string));
                    Disp.Columns.Add("連番", typeof(string));
                    Disp.Columns.Add("契約区分", typeof(string));
                    Disp.Columns.Add("作業区分", typeof(string));
                    Disp.Columns.Add("部門CD", typeof(string));//
                    Disp.Columns.Add("部門", typeof(string));

                    //1
                    Disp.Columns.Add("取引先CD", typeof(string));//
                    Disp.Columns.Add("取引先名", typeof(string));
                    Disp.Columns.Add("工事CD", typeof(string));//
                    Disp.Columns.Add("工事枝CD", typeof(string));//
                    Disp.Columns.Add("工事名", typeof(string));
                    Disp.Columns.Add("工事種別", typeof(string));
                    Disp.Columns.Add("契約項目", typeof(string));
                    Disp.Columns.Add("作業内容区分", typeof(string));
                    Disp.Columns.Add("担当名", typeof(string));

                    //2
                    Disp.Columns.Add("売上額", typeof(string));
                    Disp.Columns.Add("実施額", typeof(string));
                    Disp.Columns.Add("人件費", typeof(string));

                    //3
                    Disp.Columns.Add("支払業者CD", typeof(string));//
                    Disp.Columns.Add("支払業者名", typeof(string));
                    Disp.Columns.Add("支払額", typeof(string));

                    //4
                    Disp.Columns.Add("入力日", typeof(string));
                    Disp.Columns.Add("売上チェック", typeof(string));
                    Disp.Columns.Add("作業対象区分", typeof(string));

                    //5
                    Disp.Columns.Add("数量", typeof(string));
                    Disp.Columns.Add("単位", typeof(string));
                    Disp.Columns.Add("単価", typeof(string));
                    Disp.Columns.Add("消費税額", typeof(string));
                    Disp.Columns.Add("消費税区分", typeof(string));
                    Disp.Columns.Add("課税区分", typeof(string));

                    //6
                    Disp.Columns.Add("頁No", typeof(string));
                    Disp.Columns.Add("印刷順", typeof(string));
                    Disp.Columns.Add("まとめNo", typeof(string));
                    Disp.Columns.Add("確認書", typeof(string));

                    //7
                    Disp.Columns.Add("細目コード", typeof(string));//
                    Disp.Columns.Add("細目名", typeof(string));
                    Disp.Columns.Add("材料・物品", typeof(string));
                    Disp.Columns.Add("外注費", typeof(string));
                    Disp.Columns.Add("その他費目", typeof(string));
                    //Disp.Columns.Add("合計額", typeof(string));

                    //8
                    Disp.Columns.Add("更新者", typeof(string));
                    Disp.Columns.Add("更新日時", typeof(string));
                    Disp.Columns.Add("部門区分", typeof(string));
                    Disp.Columns.Add("支払回数", typeof(string));
                    Disp.Columns.Add("実施回数", typeof(string));
                    Disp.Columns.Add("備考", typeof(string));


                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月"] = row["売上年月"];
                        nr["連番"] = row["連番"];
                        nr["契約区分"] = row["契約区分"];
                        nr["作業区分"] = row["作業区分"];
                        nr["部門CD"] = row["部門コード"];
                        nr["部門"] = row["部門"];

                        nr["取引先CD"] = row["オーナーコード"];
                        nr["取引先名"] = row["オーナー名"];
                        nr["工事CD"] = row["工事コード"];
                        nr["工事枝CD"] = row["工事枝"];
                        nr["工事名"] = row["工事"];
                        nr["工事種別"] = row["工事種別"];
                        nr["契約項目"] = row["契約項目"];
                        nr["作業内容区分"] = row["作業内容区分"];
                        nr["担当名"] = row["担当者氏名"];

                        nr["売上額"] = string.Format("{0:#,##0}", row["売上額"]);
                        nr["実施額"] = string.Format("{0:#,##0}", row["実施額"]);
                        nr["人件費"] = string.Format("{0:#,##0}", row["人件費"]);

                        nr["支払業者CD"] = row["業者コード"];
                        nr["支払業者名"] = row["支払業者"];
                        nr["支払額"] = string.Format("{0:#,##0}", row["支払金額"]);

                        nr["入力日"] = row["売上入力日"];
                        nr["売上チェック"] = row["売上チェック"];
                        nr["作業対象区分"] = row["作業対象区分"];

                        nr["数量"] = string.Format("{0:#,##0}", row["数量"]);
                        nr["単位"] = row["単位"];
                        nr["単価"] = string.Format("{0:#,##0}", row["単価"]);
                        nr["消費税額"] = string.Format("{0:#,##0}", row["消費税額"]);
                        nr["消費税区分"] = row["消費税区分"];
                        nr["課税区分"] = row["課税区分"];

                        nr["頁No"] = row["頁No"];
                        nr["印刷順"] = row["印刷順"];
                        nr["まとめNo"] = row["まとめNo"];
                        nr["確認書"] = row["確認書"];

                        nr["細目コード"] = row["細目コード"];
                        nr["細目名"] = row["細目名"];
                        nr["材料・物品"] = row["材料・物品"];
                        nr["外注費"] = row["外注費"];
                        nr["その他費目"] = row["その他費目"];

                        nr["更新者"] = row["更新者"];
                        nr["更新日時"] = row["更新日時"];
                        nr["部門区分"] = row["部門区分"];
                        nr["支払回数"] = row["支払回数"];
                        nr["実施回数"] = row["実施回数"];
                        nr["備考"] = row["備考"];

                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
                else
                {
                    //一列表示
                    //コード非表示
                    //0
                    Disp.Columns.Add("売上年月", typeof(string));
                    Disp.Columns.Add("連番", typeof(string));
                    Disp.Columns.Add("契約区分", typeof(string));
                    Disp.Columns.Add("作業区分", typeof(string));
                    Disp.Columns.Add("部門", typeof(string));

                    //1
                    Disp.Columns.Add("取引先名", typeof(string));
                    Disp.Columns.Add("工事名", typeof(string));
                    Disp.Columns.Add("工事種別", typeof(string));
                    Disp.Columns.Add("契約項目", typeof(string));
                    Disp.Columns.Add("作業内容区分", typeof(string));
                    Disp.Columns.Add("担当名", typeof(string));

                    //2
                    Disp.Columns.Add("売上額", typeof(string));
                    Disp.Columns.Add("実施額", typeof(string));
                    Disp.Columns.Add("人件費", typeof(string));

                    //3
                    Disp.Columns.Add("支払業者名", typeof(string));
                    Disp.Columns.Add("支払額", typeof(string));

                    //4
                    Disp.Columns.Add("入力日", typeof(string));
                    Disp.Columns.Add("売上チェック", typeof(string));
                    Disp.Columns.Add("作業対象区分", typeof(string));

                    //5
                    Disp.Columns.Add("数量", typeof(string));
                    Disp.Columns.Add("単位", typeof(string));
                    Disp.Columns.Add("単価", typeof(string));
                    Disp.Columns.Add("消費税額", typeof(string));
                    Disp.Columns.Add("消費税区分", typeof(string));
                    Disp.Columns.Add("課税区分", typeof(string));

                    //6
                    Disp.Columns.Add("頁No", typeof(string));
                    Disp.Columns.Add("印刷順", typeof(string));
                    Disp.Columns.Add("まとめNo", typeof(string));
                    Disp.Columns.Add("確認書", typeof(string));

                    //7
                    Disp.Columns.Add("細目名", typeof(string));
                    Disp.Columns.Add("材料・物品", typeof(string));
                    Disp.Columns.Add("外注費", typeof(string));
                    Disp.Columns.Add("その他費目", typeof(string));
                    //Disp.Columns.Add("合計額", typeof(string));

                    //8
                    Disp.Columns.Add("更新者", typeof(string));
                    Disp.Columns.Add("更新日時", typeof(string));
                    Disp.Columns.Add("部門区分", typeof(string));
                    Disp.Columns.Add("支払回数", typeof(string));
                    Disp.Columns.Add("実施回数", typeof(string));
                    Disp.Columns.Add("備考", typeof(string));


                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        nr["売上年月"] = row["売上年月"];
                        nr["連番"] = row["連番"];
                        nr["契約区分"] = row["契約区分"];
                        nr["作業区分"] = row["作業区分"];
                        nr["部門"] = row["部門"];

                        nr["取引先名"] = row["オーナー名"];
                        nr["工事名"] = row["工事"];
                        nr["工事種別"] = row["工事種別"];
                        nr["契約項目"] = row["契約項目"];
                        nr["作業内容区分"] = row["作業内容区分"];
                        nr["担当名"] = row["担当者氏名"];

                        nr["売上額"] = string.Format("{0:#,##0}", row["売上額"]);
                        nr["実施額"] = string.Format("{0:#,##0}", row["実施額"]);
                        nr["人件費"] = string.Format("{0:#,##0}", row["人件費"]);

                        nr["支払業者名"] = row["支払業者"];
                        nr["支払額"] = string.Format("{0:#,##0}", row["支払金額"]);

                        nr["入力日"] = row["売上入力日"];
                        nr["売上チェック"] = row["売上チェック"];
                        nr["作業対象区分"] = row["作業対象区分"];

                        nr["数量"] = string.Format("{0:#,##0}", row["数量"]);
                        nr["単位"] = row["単位"];
                        nr["単価"] = string.Format("{0:#,##0}", row["単価"]);
                        nr["消費税額"] = string.Format("{0:#,##0}", row["消費税額"]);
                        nr["消費税区分"] = row["消費税区分"];
                        nr["課税区分"] = row["課税区分"];

                        nr["頁No"] = row["頁No"];
                        nr["印刷順"] = row["印刷順"];
                        nr["まとめNo"] = row["まとめNo"];
                        nr["確認書"] = row["確認書"];

                        nr["細目名"] = row["細目名"];
                        nr["材料・物品"] = row["材料・物品"];
                        nr["外注費"] = row["外注費"];
                        nr["その他費目"] = row["その他費目"];

                        nr["更新者"] = row["更新者"];
                        nr["更新日時"] = row["更新日時"];
                        nr["部門区分"] = row["部門区分"];
                        nr["支払回数"] = row["支払回数"];
                        nr["実施回数"] = row["実施回数"];
                        nr["備考"] = row["備考"];

                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
            }
            else
            {
                Disp.Columns.Add("売上年月\n連番\n契約区分\n作業区分\n部門", typeof(string));
                Disp.Columns.Add("取引先名\n工事名\n工事種別\n契約項目\n作業内容区分\n担当名", typeof(string));
                Disp.Columns.Add("売上額\n消費税額\n実施額\n人件費", typeof(string));
                Disp.Columns.Add("数量\n単価\n消費税区分\n課税区分", typeof(string));
                Disp.Columns.Add("支払業者\n支払額", typeof(string));
                Disp.Columns.Add("入力日\n売上チェック\n作業対象区分", typeof(string));
                Disp.Columns.Add("頁No\n印刷順\nまとめNo\n確認書", typeof(string));
                Disp.Columns.Add("細目名\n材料・物品\n外注費\nその他費目", typeof(string));
                Disp.Columns.Add("更新者\n更新日時\n支払回数\n実施回数\n備考", typeof(string));

                if (checkBox2.Checked)
                {
                    //複数列表示
                    //コード表示
                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        string kakuninsyo = row["確認書"].Equals("1") ? "有" : "無";
                        string sagyoukubun = row["作業対象区分"].Equals("1") ? "作業有" : "作業無";
                        string uriagec = row["売上チェック"].Equals("1") ? "売上計上有" : "売上計上無";

                        nr["売上年月\n連番\n契約区分\n作業区分\n部門"] = row["売上年月"] + "\n" + row["連番"] + "\n" + row["契約区分"] + "\n" + row["作業区分"] + "\n" + row["部門"];
                        nr["取引先名\n工事名\n工事種別\n契約項目\n作業内容区分\n担当名"] = row["オーナーコード"] + " " + row["オーナー名"] + "\n" + row["工事コード"] + "-" + row["工事枝"] + " " + row["工事"] + "\n" + row["工事種別"] + "\n" + row["契約項目"] + "\n" + row["作業内容区分"] + "\n" + row["担当者氏名"];
                        nr["売上額\n消費税額\n実施額\n人件費"] = string.Format("{0:#,##0}", row["売上額"]) + "\n" + string.Format("{0:#,##0}", row["消費税額"]) + "\n" + string.Format("{0:#,##0}", row["実施額"]) + "\n" + string.Format("{0:#,##0}", row["人件費"]);
                        nr["数量\n単価\n消費税区分\n課税区分"] = string.Format("{0:#,##0}", row["数量"]) + "" + row["単位"] + "\n" + string.Format("{0:#,##0}", row["単価"]) + "\n" + row["消費税区分"] + "\n" + row["課税区分"] + "\n";
                        nr["支払業者\n支払額"] = row["業者コード"] + "\n" + row["支払業者"] + "\n" + string.Format("{0:#,##0}", row["支払金額"]);
                        nr["入力日\n売上チェック\n作業対象区分"] = row["売上入力日"] + "\n" + uriagec + "\n" + sagyoukubun;
                        nr["頁No\n印刷順\nまとめNo\n確認書"] = row["頁No"] + "\n" + row["印刷順"] + "\n" + row["まとめNo"] + "\n" + kakuninsyo;
                        nr["細目名\n材料・物品\n外注費\nその他費目"] = row["細目コード"] + "\n" + row["細目名"] + "\n" + string.Format("{0:#,##0}", row["材料・物品"]) + "\n" + string.Format("{0:#,##0}", row["外注費"]) + "\n" + string.Format("{0:#,##0}", row["その他費目"]);
                        nr["更新者\n更新日時\n支払回数\n実施回数\n備考"] = row["更新者"] + "\n" + row["更新日時"] + "\n" + row["支払回数"] + "\n" + row["実施回数"] + "\n" + row["備考"];
                        Disp.Rows.Add(nr);
                        
                        uriageAll += Convert.ToDecimal(row["売上額"]);
                    }
                }
                else
                {
                    //複数列表示
                    //コード非表示

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow nr = Disp.NewRow();
                        string kakuninsyo = row["確認書"].Equals("1") ? "確認書有" : "確認書無";
                        string sagyoukubun = row["作業対象区分"].Equals("1") ? "作業有" : "作業無";
                        string uriagec = row["売上チェック"].Equals("1") ? "売上計上有" : "売上計上無";

                        nr["売上年月\n連番\n契約区分\n作業区分\n部門"] = row["売上年月"] + "\n" + row["連番"] + "\n" + row["契約区分"] + "\n" + row["作業区分"] + "\n" + row["部門"];
                        nr["取引先名\n工事名\n工事種別\n契約項目\n作業内容区分\n担当名"] = row["オーナー名"] + "\n" + row["工事"] + "\n" + row["工事種別"] + "\n" + row["契約項目"] + "\n" + row["作業内容区分"] + "\n" + row["担当者氏名"];
                        nr["売上額\n消費税額\n実施額\n人件費"] = string.Format("{0:#,##0}", row["売上額"]) + "\n" + string.Format("{0:#,##0}", row["消費税額"]) + "\n" + string.Format("{0:#,##0}", row["実施額"]) + "\n" + string.Format("{0:#,##0}", row["人件費"]);
                        nr["数量\n単価\n消費税区分\n課税区分"] = "数量:" + string.Format("{0:#,##0}", row["数量"]) + "" + row["単位"] + "\n単価:" + string.Format("{0:#,##0}", row["単価"]) + "\n" + row["消費税区分"] + "\n" + row["課税区分"] + "\n";
                        nr["支払業者\n支払額"] = row["支払業者"] + "\n" + string.Format("{0:#,##0}", row["支払金額"]);
                        nr["入力日\n売上チェック\n作業対象区分"] = row["売上入力日"] + "\n" + uriagec + "\n" + sagyoukubun;
                        nr["頁No\n印刷順\nまとめNo\n確認書"] = row["頁No"] + "\n" + row["印刷順"] + "\n" + row["まとめNo"] + "\n" + kakuninsyo;
                        nr["細目名\n材料・物品\n外注費\nその他費目"] = row["細目名"] + "\n" + string.Format("{0:#,##0}", row["材料・物品"]) + "\n" + string.Format("{0:#,##0}", row["外注費"]) + "\n" + string.Format("{0:#,##0}", row["その他費目"]);
                        nr["更新者\n更新日時\n支払回数\n実施回数\n備考"] = row["更新者"] + "\n" + row["更新日時"] + "\n" + row["支払回数"] + "\n" + row["実施回数"] + "\n" + row["備考"];
                        Disp.Rows.Add(nr);

                        uriageAll += Convert.ToDecimal(row["売上額"]);

                    }
                }
            }

            string ct = dt.Rows.Count.ToString();

            //TODO 共有クラスへ移動
            Com.InHistory("11_現売上検索", sDate + "～" + eDate + "【" + result + "】", ct);

            //データグリッドビューの高さ指定　※セット前にすること！
            if (checkBox1.Checked)
            {
                dataGridView1.RowTemplate.Height = 20;
            }
            else
            {
                dataGridView1.RowTemplate.Height = 105;
            }


            label1.Text = ct + " 件";
            label7.Text = "\\" + uriageAll.ToString("#,0") + "円";
            dataGridView1.DataSource = Disp;

            if (checkBox1.Checked)
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

                if (checkBox2.Checked)
                {
                    //金額右寄せ
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    //金額右寄せ
                    dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else
            {
                // セル内で文字列を折り返す
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                //金額右寄せ
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //列幅変更
                dataGridView1.Columns[0].Width = 190;
                dataGridView1.Columns[1].Width = 530;
                dataGridView1.Columns[2].Width = 120;
                dataGridView1.Columns[3].Width = 120;
                dataGridView1.Columns[4].Width = 160;
                dataGridView1.Columns[5].Width = 220;
                dataGridView1.Columns[6].Width = 140;
                dataGridView1.Columns[7].Width = 120;
                dataGridView1.Columns[8].Width = 220;
            }

            //ストップ
            sw.Stop();

            //処理速度表示
            label2.Text = sw.Elapsed.TotalSeconds.ToString("F") + " 秒";

            System.GC.Collect();
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //ボタン無効化・カーソル変更
                button1.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                //データ処理
                GetUriageData();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ処理
            GetUriageData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                //押されたキーが 0～9でない場合は、イベントをキャンセルする
                e.Handled = true;

                MessageBox.Show("半角数字で入力ください。");
            }
            else
            {

            }
        }

        private void cbizyou_CheckedChanged(object sender, EventArgs e)
        {
            if (cbizyou.Checked)
            {
                izyou.Visible = true;
                izyo.Visible = true;
            }
            else
            {
                izyou.Visible = false;
                izyo.Visible = false;
            }
        }

        private void cbika_CheckedChanged(object sender, EventArgs e)
        {
            if (cbika.Checked)
            {
                ika.Visible = true;
                ik.Visible = true;
            }
            else
            {
                ika.Visible = false;
                ik.Visible = false;
            }
        }
    }
}
