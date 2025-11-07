using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using System.Data.SqlClient;

namespace ODIS.ODIS
{
    public partial class DetailsS_KK : Form
    {
        public DetailsS_KK()
        {
            InitializeComponent();
        }

        public DetailsS_KK(string group, string ys, string ye, string bumonCD, string genbaCD, string bumon, string genba, bool hikiate)
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            GetData(group, ys, ye, bumonCD, genbaCD, bumon, genba, hikiate);
        }


        private void GetData(string group, string ys, string ye, string bumonCD, string genbaCD, string bumon, string genba, bool hikiate)
        {

            // ym 例：202304
            string ymplus = Convert.ToDateTime(ys.Insert(4, "/") + "/01").AddMonths(1).ToString("yyyyMM"); //例：202305
            string yplus = Convert.ToDateTime(ys.Insert(4, "/") + "/01").AddMonths(1).ToString("MM"); //例：05

            //当月末設定
            string str = Convert.ToDateTime(ys.Insert(4, "/") + "/01").AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd");

            if (group == "従業員数")
            {
                SqlConnection Cn;
                SqlCommand Cmd;
                SqlDataAdapter da;

                DataTable sumdt = new DataTable();

                try
                {
                    using (Cn = new SqlConnection(Com.SQLConstr))
                    {
                        using (Cmd = Cn.CreateCommand())
                        {
                            Cmd.CommandType = CommandType.StoredProcedure;

                            Cmd.Parameters.Add(new SqlParameter("date", SqlDbType.VarChar));
                            Cmd.Parameters["date"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("bumon", SqlDbType.VarChar));
                            Cmd.Parameters["bumon"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("genba", SqlDbType.VarChar));
                            Cmd.Parameters["genba"].Direction = ParameterDirection.Input;

                            Cmd.Parameters["date"].Value = str;
                            Cmd.Parameters["bumon"].Value = bumonCD;
                            Cmd.Parameters["genba"].Value = genbaCD;

                            Cmd.CommandText = "[dbo].[従業員固定給与取得_期間指定]";
                            da = new SqlDataAdapter(Cmd);
                            da.Fill(sumdt);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("エラー" + ex.ToString());
                    throw;
                }
                
                dataGridView1.DataSource = sumdt;

                this.ct.Text = sumdt.Rows.Count.ToString() + " 名";

            }
            else if (group == "人件費")
            {
                //年間に使用
                string ymsplus = ymplus;
                string ymeplus = Convert.ToDateTime(ye.Insert(4, "/") + "/01").AddMonths(1).ToString("yyyyMM"); //例：202305

                DataTable dt = new DataTable();

                string zinsql = "";

                if (hikiate)
                {
                    //引当無
                    if (ys == ye)
                    {
                        zinsql = "select * from dbo.z人件費詳細_引当無('" + ymplus + "','" + yplus + "','" + str + "') where 組織CD = '" + bumonCD + "' and 現場CD = '" + genbaCD + "' ";
                    }
                    else
                    {
                        zinsql = "select * from dbo.z人件費詳細_年間_引当無('" + ymsplus + "','" + ymeplus + "') where 組織CD = '" + bumonCD + "' and 現場CD = '" + genbaCD + "' ";
                    }
                }
                else
                {
                    //引当有
                    if (ys == ye)
                    {
                        zinsql = "select * from dbo.z人件費詳細_引当有('" + ymplus + "','" + yplus + "','" + str + "') where 組織CD = '" + bumonCD + "' and 現場CD = '" + genbaCD + "' ";
                    }
                    else
                    {
                        zinsql = "select * from dbo.z人件費詳細_年間_引当有('" + ymsplus + "','" + ymeplus + "') where 組織CD = '" + bumonCD + "' and 現場CD = '" + genbaCD + "' ";
                    }
                }

                dt = Com.GetDB(zinsql);

                Int64 sum = 0;

                foreach (DataRow row in dt.Rows)
                {
                   sum += Convert.ToInt64(row["人件費総額"]);
                }

                this.total.Text = sum.ToString("C");
                this.ct.Text = dt.Rows.Count.ToString() + " 件";

                dataGridView1.DataSource = dt;

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else if (group == "固定売上" || group == "臨時売上" || group == "売上")
            {
                string keiyaku = "";
                if (group == "固定売上")
                {
                    keiyaku = "契約";
                }
                else if (group == "臨時売上")
                {
                    keiyaku = "臨時";
                }
                else if (keiyaku == "" && (genbaCD == "19000" && bumonCD == "22010" || genbaCD == "29000" && bumonCD == "32010" || genbaCD == "39000" && bumonCD == "42010" || genbaCD == "59000" && bumonCD == "62010" || genbaCD == "69000" && bumonCD == "72010")) //施設
                {
                    //施設の全現場で売上の場合は臨時にする
                    keiyaku = "臨時";
                }
                else if (keiyaku == "" && (genbaCD != "19000" && bumonCD == "22010" || genbaCD != "29000" && bumonCD == "32010" || genbaCD != "39000" && bumonCD == "42010" || genbaCD != "59000" && bumonCD == "62010" || genbaCD != "69000" && bumonCD == "72010")) //施設
                {
                    //施設の全現場以外の場合は契約にする
                    keiyaku = "契約";
                }
                else
                {
                    keiyaku = "";
                }

                DataTable dt = new DataTable();
                string sql = "select * from pプロステージ売上データ where uriageym between '" + ys + "' and '" + ye + "' and 契約区分 like '%" + keiyaku + "%' and ";

                //技術企画とエンジ
                if (bumonCD.Substring(1, 3) == "202" || bumonCD.Substring(1, 3) == "102")
                {
                    sql += " bumoncode like '" + bumonCD + "%' ";
                }
                //植栽の臨時で全現場だった場合
                else if (keiyaku == "臨時" && genbaCD == "19000" && bumonCD == "24055")
                {
                    sql += " bumoncode like '" + bumonCD + "%' ";
                }
                //施設の臨時で全現場だった場合
                else if ((keiyaku == "臨時" || keiyaku == "") && (genbaCD == "19000" && bumonCD == "22010" || genbaCD == "29000" && bumonCD == "32010" || genbaCD == "39000" && bumonCD == "42010" || genbaCD == "59000" && bumonCD == "62010" || genbaCD == "69000" && bumonCD == "72010")) //施設
                {
                    //施設、全現場、臨時売上または売上
                    sql += " bumoncode like '" + bumonCD + "%' ";
                }
                else if (bumonCD == "24055" && genbaCD == "10101") //PPP/PFI　国際センター
                {
                    sql += " bumoncode like '" + bumonCD + "%' and koujicode = '10101'";
                }
                else if (bumonCD == "24055" && genbaCD != "10101") //PPP/PFI　植栽
                {
                    sql += " bumoncode like '" + bumonCD + "%' and koujicode <> '10101'";
                }
                else if (bumonCD == "44050" && genbaCD == "30363") //PPP/PFI　うるま市IT事業支援センター対応
                {
                    sql += " (( bumoncode like '" + bumonCD + "%' and koujicode = '30363') or (bumoncode = '24051' and koujicode in ('18000','18001'))) ";
                }
                else if (bumonCD == "24051" && genbaCD == "10246") //PPP/PFI　IT津梁パーク
                {
                    sql += " (( bumoncode like '" + bumonCD + "%' and koujicode = '10246') or (bumoncode = '24051' and koujicode in ('10274','10267','10260'))) ";
                }
                else if (bumonCD == "21060" && genbaCD == "19000") //客室　全現場
                {
                    sql += " bumoncode like '" + bumonCD + "%' and koujicode in ('18000','18001','19000')";
                }
                else
                {
                    sql += " bumoncode like '" + bumonCD + "%' and koujicode = '" + genbaCD + "'";
                }

                sql += " order by 契約区分, 作業区分, torihikisakiname, keiyakukoumoku, uriagekingaku ";

                dt = Com.GetDB(sql);

                //dataGridView1.DataSource = dt;

                decimal sum = 0;
                DataTable Disp = new DataTable();

                Disp.Columns.Add("年月", typeof(string));
                Disp.Columns.Add("契約区分", typeof(string));
                Disp.Columns.Add("作業区分", typeof(string));
                Disp.Columns.Add("契約項目", typeof(string));
                Disp.Columns.Add("金額", typeof(decimal));
                Disp.Columns.Add("取引先名", typeof(string));
                Disp.Columns.Add("現場名", typeof(string));
                Disp.Columns.Add("部門CD", typeof(string));
                Disp.Columns.Add("現場CD", typeof(string));

                foreach (DataRow row in dt.Rows)
                {
                    DataRow nr = Disp.NewRow();
                    nr["年月"] = row["uriageym"];
                    nr["契約区分"] = row["契約区分"];
                    nr["作業区分"] = row["作業区分"];
                    nr["契約項目"] = row["keiyakukoumoku"];
                    nr["金額"] = row["uriagekingaku"];
                    nr["取引先名"] = row["torihikisakiname"];
                    nr["現場名"] = row["koujiname"];
                    nr["部門CD"] = row["bumoncode"];
                    nr["現場CD"] = row["koujicode"];

                    sum += Convert.ToDecimal(row["uriagekingaku"]);

                    Disp.Rows.Add(nr);
                }

                this.total.Text = sum.ToString("C");
                this.ct.Text = Disp.Rows.Count.ToString() + " 件";

                dataGridView1.DataSource = Disp;

                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 100;
                dataGridView1.Columns[2].Width = 100;
                dataGridView1.Columns[3].Width = 250;
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[5].Width = 250;


            }
            else if (group == "諸経費" )
            {
                DataTable dt = new DataTable();

                string sql = "";
                if (ys == ye)
                {
                    sql = "select * from dbo.c管理計数詳細取得_経費('" + ys + "', '', '" + genbaCD + "', '" + bumonCD + "') order by 伝票日付, 科目コード, 取引先名, 金額";
                }
                else
                {
                    sql = "select * from dbo.c管理計数詳細取得_経費_年間('" + ys + "', '" + ye + "', '" + genbaCD + "', '" + bumonCD + "') order by 伝票日付, 科目コード, 取引先名, 金額";
                }

                dt = Com.GetDB(sql);

                decimal sum = 0;
                DataTable Disp = new DataTable();

                Disp.Columns.Add("伝票日付", typeof(string));
                Disp.Columns.Add("科目CD", typeof(string));
                Disp.Columns.Add("科目名", typeof(string));
                Disp.Columns.Add("金額", typeof(decimal));
                Disp.Columns.Add("取引先", typeof(string));
                //Disp.Columns.Add("消費税", typeof(decimal));
                Disp.Columns.Add("摘要", typeof(string));
                Disp.Columns.Add("伝票番号", typeof(string));
                //Disp.Columns.Add("工種コード", typeof(string));
                //Disp.Columns.Add("工種名", typeof(string));

                foreach (DataRow row in dt.Rows)
                {
                    DataRow nr = Disp.NewRow();
                    nr["伝票日付"] = row["伝票日付"];
                    nr["科目CD"] = row["科目コード"];
                    nr["科目名"] = row["科目名"];
                    nr["金額"] = row["金額"];
                    nr["取引先"] = row["取引先名"];
                    //nr["消費税"] = Convert.ToDecimal(row["消費税"]);
                    nr["摘要"] = row["摘要文"];
                    nr["伝票番号"] = row["伝票番号"];
                    //nr["工種コード"] = row["工種コード"];
                    //nr["工種名"] = row["工種名"];
                    sum += Convert.ToDecimal(row["金額"]);
                    //zeisum += Convert.ToDecimal(row["消費税"]);

                    Disp.Rows.Add(nr);
                }

                this.total.Text = sum.ToString("C");
                this.ct.Text = Disp.Rows.Count.ToString() + " 件";

                dataGridView1.DataSource = Disp;

                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 100;
                dataGridView1.Columns[2].Width = 150;
                dataGridView1.Columns[3].Width = 120;
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[5].Width = 500;
            }

            this.bumon.Text = bumon;
            this.genba.Text = genba;
            if (ys == ye)
            {
                this.month.Text = ys;
            }
            else
            {
                this.month.Text = "年間";
            }
            this.koumoku.Text = group;


            //検索履歴登録
            Com.InHistory("51_管理計数_詳細表示", group + " " + ys + " " + bumon + " " + genba, dataGridView1.RowCount.ToString());
        }

            

   }
}
