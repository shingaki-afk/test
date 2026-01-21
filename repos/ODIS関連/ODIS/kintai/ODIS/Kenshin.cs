using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Kenshin : Form
    {

        private DataTable dt = new DataTable();
        private DataTable dtlist = new DataTable();
        private Boolean flg = false;
        private string tikuflg = "";
        private System.Data.SqlTypes.SqlString SqlStrNull = System.Data.SqlTypes.SqlString.Null;

        public Kenshin()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;
            
            //フォントサイズの変更
            //dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);

            //行ヘッダを非表示
            dataGridView1.RowHeadersVisible = false;

            // 選択モードを行単位での選択のみにする
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            comboBox1.Items.Add("2021");
            comboBox1.Items.Add("2022");
            comboBox1.SelectedIndex = 1;

            koumoku.Items.Add("01_深夜業");
            koumoku.Items.Add("02_定期");
            
            koumoku.SelectedIndex = 0;

            checkedListBox3.Items.Add("加入者");
            checkedListBox3.Items.Add("未");
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }

            kikan.Items.Add("健康づくり財団");
            kikan.Items.Add("那覇市立病院");
            kikan.Items.Add("中部地区医師会");
            kikan.Items.Add("石垣島徳洲会");
            kikan.Items.Add("北部○○");


            ninngen.Items.Add("");
            ninngen.Items.Add("○");

            nyuugan.Items.Add("");
            nyuugan.Items.Add("○");

            shikyuu.Items.Add("");
            shikyuu.Items.Add("○");

            nijikenshin.Items.Add("３項目");
            nijikenshin.Items.Add("４項目");

            //フラグ設定
            if (Program.loginname == "高江洲　華子" || Program.loginname == "我那覇　早苗" || Program.loginname == "喜屋武　大祐" || Program.loginname == "太田　朋宏" || Program.loginname == "新垣　聖悟")
            {
                flg = true;
            }
            else if (Program.loginname == "田村　美由紀" || Program.loginname == "石田　稚菜")
            {
                tikuflg = "yaeyama";
                flg = true;
            }
            else if (Program.loginname == "大城　佐代子" || Program.loginname == "大城　森一")
            {
                tikuflg = "hokubu";
                flg = true;
            }
            else
            {
                flg = false;
                tableLayoutPanel2.Visible = false;
                tableLayoutPanel3.Visible = false;
            }

            //SetTiku();
            SetBumon();
            GetData();




                Com.InHistory("健診入力", "", "");
        }

        //private void GetList()
        //{
        //    //グリッド表示クリア
        //    dataGridView2.DataSource = "";

        //    //テーブルクリア
        //    dtlist.Clear();

        //    string sql = "select 担当区分, count(*) as 受診対象者, sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end) as 受診者, Convert(money, sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end)) *100 / Convert(money, count(*)) as 受診率 from dbo.k健康診断データ取得('2021', '01_定期', '2021/08/31') group by 担当区分 order by 担当区分";
            

        //}
        private void GetData()
        {
            //DataClear();

            //グリッド表示クリア
            dataGridView1.DataSource = "";
            dataGridView2.DataSource = "";

            //テーブルクリア
            dt.Clear();
            dtlist.Clear();

            string sql = "select * from dbo.k健康診断データ取得('2021','01_定期','2021/08/31')";
            //string sqllist = "select 担当区分, count(*) as 受診対象者, sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end) as 受診者, Convert(money, sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end)) *100 / Convert(money, count(*)) as 受診率 from dbo.k健康診断データ取得('2021', '01_定期', '2021/08/31') ";

            string sqllist = "select 担当区分, sum(受診対象者) as 受診対象者,sum(受診者) as 受診者, Convert(money, sum(受診者)) / Convert(money, sum(受診対象者)) * 100 as 受診率";
            sqllist += " from(select case when 担当区分 = '13_広域' then '20_浦添' when 担当区分 = '21_役員室' then '20_浦添' when 担当区分 = '22_営業' then '20_浦添'";
            sqllist += " when 担当区分 = '23_経営企画' then '20_浦添' when 担当区分 = '24_総務' then '20_浦添' else 担当区分 end as 担当区分, count(*) as 受診対象者, ";
            sqllist += " sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end) as 受診者";
            sqllist += " from dbo.k健康診断データ取得('2021', '01_定期', '2021/08/31') ";
            //sqllist += " group by 担当区分 ) temp group by 担当区分";

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

            //部門
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 担当区分 <> '" + checkedListBox2.Items[i].ToString() + "'";
                    //flg = true;
                }
            }

            //社保
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    if (checkedListBox3.Items[i].ToString() == "加入者")
                    {
                        result += " and 健康保険整理番号 is null ";
                    }
                    else
                    {
                        result += " and 健康保険整理番号 is not null ";
                    }
                    
                    //flg = true;
                }
            }

            //集計結果　退職者外す。
            //string result2 = result + " and 退職年月日 is null ";


            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = " where " + result.Remove(0, 4);
            }

            //先頭が「and」の場合、削除する
            //if (result2.StartsWith(" and"))
            //{
            //    result2 = " where " + result2.Remove(0, 4);
            //}

            sql += result + " order by 組織CD, 現場CD, カナ名 ";
            //sqllist += result2 + " group by 担当区分 order by 担当区分 ";
            sqllist += result + " group by 担当区分 union all select '合計' as 担当区分, count(*) as 受診対象者,  sum(case when datediff(day, 受診日, GETDATE()) > 0 then 1 else 0 end) as 受診者 from dbo.k健康診断データ取得('2021', '01_定期', '2021/08/31') " + result + " ) temp group by 担当区分";

            dt = Com.GetDB(sql);
            dtlist = Com.GetDB(sqllist);

            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = dtlist;

            //幅
            dataGridView1.Columns["社員番号"].Width = 60;
            dataGridView1.Columns["氏名"].Width = 120;
            dataGridView1.Columns["カナ名"].Width = 100;
            dataGridView1.Columns["生年月日"].Width = 70;
            dataGridView1.Columns["地区名"].Width = 60;
            dataGridView1.Columns["組織CD"].Width = 60;
            dataGridView1.Columns["組織名"].Width = 90;
            dataGridView1.Columns["現場CD"].Width = 60;
            dataGridView1.Columns["現場名"].Width = 150;
            dataGridView1.Columns["社員区分名"].Width = 60;
            dataGridView1.Columns["入社年月日"].Width = 70;
            dataGridView1.Columns["健康保険整理番号"].Width = 60;
            dataGridView1.Columns["年齢"].Width = 60;
            dataGridView1.Columns["健診種類"].Width = 60;
            //dataGridView1.Columns["退職年月日"].Width = 70;
            //dataGridView1.Columns["予約日"].Width = 70;


            dataGridView1.Columns["受診日"].Width = 70;
            dataGridView1.Columns["受診結果"].Width = 70;
            dataGridView1.Columns["受診機関"].Width = 120;

            dataGridView1.Columns["人間ドック"].Width = 40;
            dataGridView1.Columns["乳がん"].Width = 40;
            dataGridView1.Columns["子宮頸がん"].Width = 40;
            dataGridView1.Columns["二次健診"].Width = 70;
            dataGridView1.Columns["二次結果"].Width = 70;

            dataGridView1.Columns["備考"].Width = 150;

            //非表示
            if (flg)
            {

            }
            else
            {
                dataGridView1.Columns["生年月日"].Visible = false;
                dataGridView1.Columns["現場CD"].Visible = false;
                dataGridView1.Columns["組織CD"].Visible = false;
                dataGridView1.Columns["入社年月日"].Visible = false;
                dataGridView1.Columns["年齢"].Visible = false;

                dataGridView1.Columns["人間ドック"].Visible = false;
                dataGridView1.Columns["乳がん"].Visible = false;
                dataGridView1.Columns["子宮頸がん"].Visible = false;
                dataGridView1.Columns["二次健診"].Visible = false;
                dataGridView1.Columns["二次結果"].Visible = false;
            }

            dataGridView1.Columns["性別"].Visible = false;
            dataGridView1.Columns["地区CD"].Visible = false;
            dataGridView1.Columns["組織CD"].Visible = false;
            
            dataGridView1.Columns["住所"].Visible = false;
            dataGridView1.Columns["対象年"].Visible = false;
            dataGridView1.Columns["項目"].Visible = false;
            dataGridView1.Columns["reskey"].Visible = false;
            dataGridView1.Columns["担当区分"].Visible = false;
            dataGridView1.Columns["担当事務"].Visible = false;
            

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                GetData();
            }
        }

        //private void label4_Click(object sender, EventArgs e)
        //{
        //    if (checkedListBox1.CheckedItems.Count == 0)
        //    {
        //        for (int i = 0; i < checkedListBox1.Items.Count; i++)
        //        {
        //            checkedListBox1.SetItemChecked(i, true);
        //        }
        //    }
        //    else
        //    {
        //        for (int i = 0; i < checkedListBox1.Items.Count; i++)
        //        {
        //            checkedListBox1.SetItemChecked(i, false);
        //        }
        //    }

        //    SetBumon();
        //    GetData(); 
        //}

        private void label23_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
            }

            GetData();
        }

        //private void SetTiku()
        //{
        //    checkedListBox1.Items.Clear();

        //    DataTable dt = new DataTable();
        //    string sql = "select distinct 地区名 from dbo.k健康診断データ取得('2021','01_定期','2021/08/31') where 地区名 <> 'dummy' ";
        //    dt = Com.GetDB(sql);

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        checkedListBox1.Items.Add(row["地区名"]);
        //    }

        //    for (int i = 0; i < checkedListBox1.Items.Count; i++)
        //    {
        //        checkedListBox1.SetItemChecked(i, true);
        //    }
        //}

        private void SetBumon()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox2.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.k健康診断データ取得('2021','01_定期','2021/08/31') where 担当区分 <> 'dummy' ";

            //for (int i = 0; i < checkedListBox1.Items.Count; i++)
            //{
            //    if (!checkedListBox1.GetItemChecked(i)) sql += " and 地区名 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            //}

            sql += " order by 担当区分 ";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["担当区分"]);
            }

            if (tikuflg == "yaeyama")
            {
                checkedListBox2.SetItemChecked(7, true);
            }
            else if (tikuflg == "hokubu")
            {
                checkedListBox2.SetItemChecked(8, true);
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }


        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetBumon();
            GetData();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            GetMeisaiData();
        }

        private void DataClear()
        {
            //クリア処理
            no.Text = "";
            name.Text = "";
            kanamei.Text = "";
            tiku.Text = "";
            soshiki.Text = "";
            genba.Text = "";
            skubun.Text = "";

            nendo.Text = "";
            koumokulbl.Text = "";
            zyushinday.Value = null;
            kekkaday.Value = null;
            //zyushinday.Text = "";
            //kekkaday.Text = "";
            kikan.Text = "";
            ninngen.Text = "";
            nyuugan.Text = "";
            shikyuu.Text = "";
            nijikenshin.Text = "";
            nijikenshinday.Value = null;
            //nijikenshinday.valu = "";
            bikou.Text = "";
        }
        private void GetMeisaiData()
        {

            //ヘッダは対象外
            if (dataGridView1.CurrentCell != null)
            {
                DataClear();

                DataGridViewRow dgr = dataGridView1.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;

                no.Text = drv["社員番号"].ToString();
                name.Text = drv["氏名"].ToString();
                kanamei.Text = drv["カナ名"].ToString();
                tiku.Text = drv["地区名"].ToString();
                soshiki.Text = drv["組織名"].ToString();
                genba.Text = drv["現場名"].ToString();
                skubun.Text = drv["社員区分名"].ToString();

                nendo.Text = drv["対象年"].ToString();
                koumokulbl.Text = drv["項目"].ToString();
                //yoyaku.Value = drv[""].ToString();
                //zyushinday.Value = drv["受診日"].ToString();
                zyushinday.Value = drv["受診日"].Equals(DBNull.Value) ? null : drv["受診日"].ToString();
                kekkaday.Value = drv["受診結果"].Equals(DBNull.Value) ? null : drv["受診結果"].ToString();

                //kekkaday.Value = drv["受診結果"].ToString();
                kikan.Text = drv["受診機関"].ToString();
                ninngen.Text = drv["人間ドック"].ToString();
                nyuugan.Text = drv["乳がん"].ToString();
                shikyuu.Text = drv["子宮頸がん"].ToString();

                nijikenshin.Text = drv["二次健診"].ToString();
                //nijikenshinday.Value = drv["二次結果"].ToString();
                nijikenshinday.Value = drv["二次結果"].Equals(DBNull.Value) ? null : drv["二次結果"].ToString();
                bikou.Text = drv["備考"].ToString();

                if (!flg && drv["健康保険整理番号"].ToString() != "")
                {
                    //健康診断担当以外は社保入力不可
                    tableLayoutPanel1.Visible = false;
                }
                else
                {
                    tableLayoutPanel1.Visible = true;
                }


            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int ri = dataGridView1.CurrentCell.RowIndex;

            SetData();
            GetData();

            //
            dataGridView1.CurrentCell = dataGridView1[1, ri];

            GetMeisaiData();
        }

        private void SetData()
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
                    Cmd.CommandText = "[dbo].[k健康診断更新]";

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("対象年", SqlDbType.VarChar)); Cmd.Parameters["対象年"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("項目", SqlDbType.VarChar)); Cmd.Parameters["項目"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("受診日", SqlDbType.Date)); Cmd.Parameters["受診日"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("受診結果", SqlDbType.Date)); Cmd.Parameters["受診結果"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("受診機関", SqlDbType.VarChar)); Cmd.Parameters["受診機関"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("人間ドック", SqlDbType.VarChar)); Cmd.Parameters["人間ドック"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("乳がん", SqlDbType.VarChar)); Cmd.Parameters["乳がん"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("子宮頸がん", SqlDbType.VarChar)); Cmd.Parameters["子宮頸がん"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("二次健診", SqlDbType.VarChar)); Cmd.Parameters["二次健診"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("二次結果", SqlDbType.Date)); Cmd.Parameters["二次結果"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar)); Cmd.Parameters["備考"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = no.Text;
                    Cmd.Parameters["対象年"].Value = nendo.Text;
                    Cmd.Parameters["項目"].Value = koumokulbl.Text;
                    Cmd.Parameters["受診日"].Value = zyushinday.Value == null ? DBNull.Value : zyushinday.Value; 
                    Cmd.Parameters["受診結果"].Value = kekkaday.Value == null ? DBNull.Value : kekkaday.Value;

                    //if (kekkaday.Value == null)
                    //{
                    //    Cmd.Parameters["受診結果"].Value = DBNull.Value;
                    //}
                    //else
                    //{
                    //    Cmd.Parameters["受診結果"].Value = kekkaday.Value;
                    //}

                    Cmd.Parameters["受診機関"].Value = kikan.Text;
                    Cmd.Parameters["人間ドック"].Value = ninngen.Text;
                    Cmd.Parameters["乳がん"].Value = nyuugan.Text;
                    Cmd.Parameters["子宮頸がん"].Value = shikyuu.Text;
                    Cmd.Parameters["二次健診"].Value = nijikenshin.Text;
                    Cmd.Parameters["二次結果"].Value = nijikenshinday == null ? SqlStrNull : nijikenshinday.Value;
                    Cmd.Parameters["備考"].Value = bikou.Text;
                    //Cmd.Parameters["氏名"].Value = shimei.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\健康診断.xlsx"); return;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\17_総務部\02_人事労務\健康診断関連\定期\2021\②名簿"); return;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\23_労働安全衛生\１２：那覇地区、労働安全衛生委員会\第５０期　労働安全衛生委員会\健康診断\定期健康診断"); return;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\健康診断・ストレスチェック（年間）再編集.xlsx"); return;
        }

        private void label6_Click(object sender, EventArgs e)
        {
            if (checkedListBox3.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, false);
                }
            }

            GetData();
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
