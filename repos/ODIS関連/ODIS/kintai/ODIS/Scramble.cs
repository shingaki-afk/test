using C1.C1Excel;
using Microsoft.VisualBasic;
using ODIS.ODIS;
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
    public partial class Scramble : Form
    {
        //private SqlConnection Cn;
        //private SqlDataAdapter da;
        //private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public Scramble()
        {

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            DataTable dt = Com.GetDB("select distinct 担当区分 from dbo.t通勤管理一覧取得 order by 担当区分");
            foreach (DataRow drw in dt.Rows)
            {
                comboBox1.Items.Add(drw["担当区分"].ToString());
            }

            comboBox1.Items.Add("全体");

            comboBox1.SelectedIndex = comboBox1.FindString(Program.loginbusyo);

            //comboBox1.SelectedIndex = 4;

            //通勤方法
            tuukinhouhou.Items.Add("1 車");
            tuukinhouhou.Items.Add("2 バイク");
            //tuukinhouhou.Items.Add("3 徒歩・自転車");
            tuukinhouhou.Items.Add("4 バス・モノレール");
            tuukinhouhou.Items.Add("5 送迎(会社)");
            tuukinhouhou.Items.Add("6 送迎(知人・親族)");
            tuukinhouhou.Items.Add("7 業務車両");
            tuukinhouhou.Items.Add("8 徒歩");
            tuukinhouhou.Items.Add("9 自転車");

            //通勤手当区分
            tuukinkubun.Items.Add("");
            tuukinkubun.Items.Add("1 実費精算");

            if (Program.loginname != "喜屋武　大祐")
            {
                tuukinkubun.Enabled = false;
                tuukinhouhou.Enabled = false;
                katakyori.Enabled = false;
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

            Com.InHistory("64_通勤管理", "", "");
        }



        private void GetData()
        {
            if (comboBox1.SelectedItem == null) return;

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            //テーブルクリア
            dt.Clear();

            string sql = "select * from dbo.t通勤管理一覧取得";

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');
            string result = "";


            if (comboBox1.SelectedItem.ToString() == "全体")
            {
                if (checkBox2.Checked)
                {
                    result += " where ";
                    //result += " エラーと警告 <> '' or 通勤手当区分 = '2 規程外ルール' or 通勤方法 = '8 不明' or ";
                    result += " (通勤方法 = '1 車' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '6 送迎(知人・親族)' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '2 バイク' and(免許証 < GETDATE() or 自賠責 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 自賠責 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '7 業務車両' and(免許証 < GETDATE() or 免許証 is null)) or ";
                    result += " (通勤方法 = '9 自転車' and(任意保険 < GETDATE() or 任意保険 is null)) ";
                }
                else if (checkBox1.Checked)
                {
                    result += " where ";
                    result += " エラーと警告 <> '' or 通勤手当区分 = '2 規程外ルール' or 通勤方法 = '8 不明' or ";
                    result += " (通勤方法 = '1 車' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '6 送迎(知人・親族)' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '2 バイク' and(免許証 < GETDATE() or 自賠責 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 自賠責 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '7 業務車両' and(免許証 < GETDATE() or 免許証 is null)) or ";
                    result += " (通勤方法 = '9 自転車' and(任意保険 < GETDATE() or 任意保険 is null)) ";
                }
            }
            else
            {
                result += " where 担当区分 = '" + comboBox1.SelectedItem.ToString() + "' ";

                if (checkBox2.Checked)
                {
                    result += " and( ";
                    //result += " エラーと警告 <> '' or 通勤手当区分 = '2 規程外ルール' or 通勤方法 = '8 不明' or ";
                    result += " (通勤方法 = '1 車' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '6 送迎(知人・親族)' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '2 バイク' and(免許証 < GETDATE() or 自賠責 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 自賠責 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '7 業務車両' and(免許証 < GETDATE() or 免許証 is null)) or ";
                    result += " (通勤方法 = '9 自転車' and(任意保険 < GETDATE() or 任意保険 is null))) ";
                }
                else if (checkBox1.Checked)
                {
                    result += " and( ";
                    result += " エラーと警告 <> '' or 通勤手当区分 = '2 規程外ルール' or 通勤方法 = '8 不明' or ";
                    result += " (通勤方法 = '1 車' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '6 送迎(知人・親族)' and(免許証 < GETDATE() or 車検証 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 車検証 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '2 バイク' and(免許証 < GETDATE() or 自賠責 < GETDATE() or 任意保険 < GETDATE() or 免許証 is null or 自賠責 is null or 任意保険 is null)) or ";
                    result += " (通勤方法 = '7 業務車両' and(免許証 < GETDATE() or 免許証 is null)) or ";
                    result += " (通勤方法 = '9 自転車' and(任意保険 < GETDATE() or 任意保険 is null))) ";
                }
            }

            if (ar[0] != "")
                {
                    foreach (string s in ar)
                    {
                        result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                    }
                }

                //先頭が「and」の場合、削除する
                if (result.StartsWith(" and"))
            {
                result = " where " + result.Remove(0, 4);
            }

            sql += result + " order by 組織CD, 現場CD, カナ名 ";



            dt = Com.GetDB(sql);

            dataGridView1.DataSource = dt;

            //dataGridView1.Columns[0].ReadOnly = true;
            //dataGridView1.Columns["通勤非課税"].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns["通勤課税"].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns["回数１単価"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["片道距離"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns["片道料金"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns["通勤非課税"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns["通勤課税"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns["回数１単価"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns["社員番号"].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns["氏名"].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns["組織名"].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns["現場名"].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns["給与区分"].DefaultCellStyle.BackColor = Color.AliceBlue;
            //dataGridView1.Columns["通勤非課税"].DefaultCellStyle.BackColor = Color.AliceBlue;
            //dataGridView1.Columns["通勤課税"].DefaultCellStyle.BackColor = Color.AliceBlue;
            //dataGridView1.Columns["回数１単価"].DefaultCellStyle.BackColor = Color.AliceBlue;

            dataGridView1.Columns["社員番号"].Width = 60;
            dataGridView1.Columns["氏名"].Width = 120;
            dataGridView1.Columns["組織名"].Width = 90;
            dataGridView1.Columns["現場名"].Width = 150;
            dataGridView1.Columns["給与区分"].Width = 60;
            dataGridView1.Columns["通勤手当区分"].Width = 90;
            dataGridView1.Columns["通勤方法"].Width = 90;
            dataGridView1.Columns["免許証"].Width = 70;
            dataGridView1.Columns["車検証"].Width = 70;
            dataGridView1.Columns["自賠責"].Width = 70;
            dataGridView1.Columns["任意保険"].Width = 70;
            dataGridView1.Columns["片道距離"].Width = 30;
            //dataGridView1.Columns["片道料金"].Width = 30;
            dataGridView1.Columns["備考"].Width = 100;
            //dataGridView1.Columns["通勤非課税"].Width = 40;
            //dataGridView1.Columns["通勤課税"].Width = 40;
            //dataGridView1.Columns["回数１単価"].Width = 40;
            dataGridView1.Columns["メーカー"].Width = 50;
            dataGridView1.Columns["車名"].Width = 50;
            dataGridView1.Columns["色"].Width = 50;
            dataGridView1.Columns["車両番号"].Width = 100;
            dataGridView1.Columns["担当区分"].Visible = false; //削除
            dataGridView1.Columns["エラーと警告"].Width = 100;
            //dataGridView1.Columns["非課税算出"].Width = 40;
            //dataGridView1.Columns["課税算出"].Width = 40;
            dataGridView1.Columns["組織CD"].Visible = false;　//削除
            dataGridView1.Columns["現場CD"].Visible = false;　//削除
            dataGridView1.Columns["カナ名"].Visible = false; //削除
            dataGridView1.Columns["管理No"].Visible = false;　//削除
            dataGridView1.Columns["reskey"].Visible = false;　//削除
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            GridDisp(e, 1);
        }

        private void GridDisp(DataGridViewCellFormattingEventArgs e, int i)
        {

            //セルの列を確認
            DateTime val;
            if (e.Value != null && (e.ColumnIndex == 7 || e.ColumnIndex == 8 || e.ColumnIndex == 9 || e.ColumnIndex == 10) && DateTime.TryParse(e.Value.ToString(), out val))
            {
                //4つの期限項目で日付が入っている
                if (val < DateTime.Now)
                {
                    //有効期限切れ
                    e.CellStyle.BackColor = Color.Red;
                }
                else if (val < DateTime.Now.AddMonths(+1))
                {
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (val < DateTime.Now.AddMonths(+2))
                {
                    e.CellStyle.BackColor = Color.SpringGreen;
                }
            }
            else
            {
                //メイン画面
                if (e.ColumnIndex == 7 || e.ColumnIndex == 8 || e.ColumnIndex == 9 || e.ColumnIndex == 10)
                {
                    if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "1 車" || Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "2 バイク" || Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "4 実費精算" || Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "7 業務車両" || Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "9 自転車")
                    {
                        if (e.ColumnIndex == 9 && Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "1 車")
                        {
                            e.CellStyle.BackColor = Color.LightGray;
                        }
                        else if (e.ColumnIndex == 9 && Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "4 実費精算")
                        {
                            e.CellStyle.BackColor = Color.LightGray;
                        }
                        else if (e.ColumnIndex == 8 && Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "2 バイク")
                        {
                            e.CellStyle.BackColor = Color.LightGray;
                        }
                        else if ((e.ColumnIndex == 8 || e.ColumnIndex == 9 || e.ColumnIndex == 10) && Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "7 業務車両")
                        {
                            e.CellStyle.BackColor = Color.LightGray;
                        }
                        else if ((e.ColumnIndex == 7 || e.ColumnIndex == 8 || e.ColumnIndex == 9) && Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "9 自転車")
                        {
                            e.CellStyle.BackColor = Color.LightGray;
                        }
                        else
                        {
                            e.CellStyle.BackColor = Color.Red;
                        }
                    }
                    else if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[6].Value) == "")
                    {
                        //通勤方法が入力されてない場合
                        e.CellStyle.BackColor = Color.Red;
                    }
                    else
                    {
                        e.CellStyle.BackColor = Color.LightGray;
                    }
                }

            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SetTuukin();

            GetData();
        }

        //通勤管理テーブル情報登録更新
        private void SetTuukin()
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
                    Cmd.CommandText = "[dbo].[t通勤管理テーブル登録更新]";

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.Char)); Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("氏名", SqlDbType.VarChar)); Cmd.Parameters["氏名"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("管理No", SqlDbType.VarChar)); Cmd.Parameters["管理No"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("通勤方法", SqlDbType.VarChar)); Cmd.Parameters["通勤方法"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("免許証", SqlDbType.Date)); Cmd.Parameters["免許証"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("車検証", SqlDbType.Date)); Cmd.Parameters["車検証"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("自賠責", SqlDbType.Date)); Cmd.Parameters["自賠責"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("任意保険", SqlDbType.Date)); Cmd.Parameters["任意保険"].Direction = ParameterDirection.Input;
                    //Cmd.Parameters.Add(new SqlParameter("通勤手当区分", SqlDbType.VarChar)); Cmd.Parameters["通勤手当区分"].Direction = ParameterDirection.Input;
                    //Cmd.Parameters.Add(new SqlParameter("片道距離", SqlDbType.Decimal)); Cmd.Parameters["片道距離"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("メーカー", SqlDbType.VarChar)); Cmd.Parameters["メーカー"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("車名", SqlDbType.VarChar)); Cmd.Parameters["車名"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("色", SqlDbType.VarChar)); Cmd.Parameters["色"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("車両番号", SqlDbType.VarChar)); Cmd.Parameters["車両番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("許可証No", SqlDbType.VarChar)); Cmd.Parameters["許可証No"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar)); Cmd.Parameters["備考"].Direction = ParameterDirection.Input;
                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar)); Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["社員番号"].Value = syainno.Text;
                    Cmd.Parameters["氏名"].Value = shimei.Text;

                    if (this.kanrino.Text == "")
                    {
                        DataTable tuukinid = new DataTable();
                        tuukinid = Com.GetDB("select max(管理No) from dbo.t通勤管理テーブル where 社員番号 = '" + syainno.Text + "'");
                        if (tuukinid.Rows[0][0] == DBNull.Value)
                        {
                            Cmd.Parameters["管理No"].Value = 1;
                        }
                        else
                        {
                            Cmd.Parameters["管理No"].Value = Convert.ToInt16(tuukinid.Rows[0][0]) + 1;
                        }
                    }
                    else
                    {
                        Cmd.Parameters["管理No"].Value = kanrino.Text;
                    }

                    if (tuukinhouhou.SelectedItem.ToString() == "")
                    {
                        Cmd.Parameters["通勤方法"].Value = DBNull.Value;
                    }
                    else
                    {
                        Cmd.Parameters["通勤方法"].Value = tuukinhouhou.SelectedItem.ToString();
                    }

                    if (menkyonew.Text == "")
                    {
                        Cmd.Parameters["免許証"].Value = DBNull.Value;
                    }
                    else
                    {
                        Cmd.Parameters["免許証"].Value = menkyonew.Text;
                    }

                    if (syakennew.Text == "")
                    {
                        Cmd.Parameters["車検証"].Value = DBNull.Value;
                    }
                    else
                    {
                        Cmd.Parameters["車検証"].Value = syakennew.Text;
                    }

                    if (zibainew.Text == "")
                    {
                        Cmd.Parameters["自賠責"].Value = DBNull.Value;
                    }
                    else
                    {
                        Cmd.Parameters["自賠責"].Value = zibainew.Text;
                    }

                    if (ninninew.Text == "")
                    {
                        Cmd.Parameters["任意保険"].Value = DBNull.Value;
                    }
                    else
                    {
                        Cmd.Parameters["任意保険"].Value = ninninew.Text;
                    }


                    //if (tuukinkubun.SelectedItem.ToString() == "")
                    //{
                    //    Cmd.Parameters["通勤手当区分"].Value = DBNull.Value;
                    //}
                    //else
                    //{
                    //    Cmd.Parameters["通勤手当区分"].Value = tuukinkubun.SelectedItem.ToString();
                    //}

                    //if (katakyori.Text == "")
                    //{
                    //    Cmd.Parameters["片道距離"].Value = DBNull.Value;
                    //}
                    //else
                    //{
                    //    Cmd.Parameters["片道距離"].Value = katakyori.Text;
                    //}

                    Cmd.Parameters["メーカー"].Value = meka.Text;
                    Cmd.Parameters["車名"].Value = syamei.Text;
                    Cmd.Parameters["色"].Value = iro.Text;
                    Cmd.Parameters["車両番号"].Value = syaryou.Text;

                    Cmd.Parameters["許可証No"].Value = kyoka.Text;
                    Cmd.Parameters["備考"].Value = bikou.Text;
                    
                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //【参考】
            //1回のクリックでエディットモードにする処置  
            //DataGridView dgv = sender as DataGridView;

            //if (e.ColumnIndex >= 0)
            //{
            //    if (dgv.Columns[e.ColumnIndex] is DataGridViewTextBoxColumn)
            //    {
            //        SendKeys.Send("{F2}");
            //    }
            //}
        }


        private void GetData(string no, string howto, string carname)
        {
            DataTable dt = new DataTable();
            //dt = Com.GetDB("select * from dbo.t通勤管理テーブル a left join dbo.t通勤手当元データ b on a.社員番号 = b.社員番号 and b.適用終了日 = '9999/12/31' where a.社員番号 = '" + no + "' and b.通勤方法　= '" + howto + "' and a.管理No = '" + carname + "'");
            dt = Com.GetDB("select * from dbo.t通勤管理テーブル a left join dbo.t通勤手当元データ b on a.社員番号 = b.社員番号 and b.適用終了日 = '9999/12/31' where a.社員番号 = '" + no + "' and a.管理No = '" + carname + "'");

            if (dt.Rows.Count == 0) return;

            syainno.Text = dt.Rows[0]["社員番号"].ToString();
            shimei.Text = dt.Rows[0]["氏名"].ToString();
            kanrino.Text = dt.Rows[0]["管理No"].ToString();

            //TODO 2台もち対応
            if (dt.Rows[0]["管理No"].ToString() == "2" && dt.Rows[0]["通勤方法"].ToString() == "2 バイク")
            {
                tuukinhouhou.SelectedIndex = tuukinhouhou.FindString("2 バイク");
            }
            else if (dt.Rows[0]["管理No"].ToString() == "2" && dt.Rows[0]["通勤方法"].ToString() == "1 車")
            {
                tuukinhouhou.SelectedIndex = tuukinhouhou.FindString("1 車");
            }
            else 
            {
                //TODO 　通勤方法削除
                //tuukinhouhou.SelectedIndex = tuukinhouhou.FindString(dt.Rows[0]["通勤方法1"].ToString());
                tuukinhouhou.SelectedIndex = tuukinhouhou.FindString(dt.Rows[0]["通勤方法"].ToString());
            }

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
            tuukinkubun.SelectedIndex = tuukinkubun.FindString(dt.Rows[0]["通勤手当区分"].ToString());

            katakyori.Text = dt.Rows[0]["片道距離"].ToString();
            meka.Text = dt.Rows[0]["メーカー"].ToString();
            syamei.Text = dt.Rows[0]["車名"].ToString();
            iro.Text = dt.Rows[0]["色"].ToString();
            syaryou.Text = dt.Rows[0]["車両番号"].ToString();

            bikou.Text = dt.Rows[0]["備考"].ToString();
            kyoka.Text = dt.Rows[0]["許可証No"].ToString();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("入力値エラー");
            e.Cancel = false;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            comboBox1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            comboBox1.Enabled = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //ヘッダは対象外
            if (dataGridView1.CurrentCell != null)
            {
                //TODO クリア
                //KazokuClear();
                DataGridViewRow dgr = dataGridView1.CurrentRow;
                if (dgr == null) return;
                DataRowView drv = (DataRowView)dgr.DataBoundItem;
                GetData(drv[0].ToString(), drv[6].ToString(), drv[22].ToString());
            }
        }



        #region 和暦対応
        private void wareki_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wareki.SelectedItem == null || wareki.SelectedItem.ToString() == "") return;

            //今日の西暦になってしまう！
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
                if (wareki.SelectedItem.ToString()?.Substring(0, 2) == "平成")
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
            if (wareki2.SelectedItem == null || wareki2.SelectedItem.ToString() == "") return;

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
            if (wareki3.SelectedItem == null || wareki3.SelectedItem.ToString() == "") return;

            //if (zibainew.Text == "") zibainew.Value = DateTime.Today;

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
            if (wareki4.SelectedItem == null || wareki4.SelectedItem.ToString() == "") return;

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

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.ToString() == "全体") return;

            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button2.Enabled = false;

            //新しいワークブックを作成します。
            C1XLBook c1XLBook1 = new C1XLBook();

            //ブックをロードします
                c1XLBook1.Load(@"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\40_更新案内.xlsx");
            //リストシート
            XLSheet ls = c1XLBook1.Sheets["List"];
            
            //対象データ取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.k更新案内取得('" + DateTime.Today.ToString("yyyy/MM/dd") + "','" + comboBox1.SelectedItem  + "')  order by 組織CD, 現場CD, カナ名");


            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("案内対象者はいません。");
                //マウスカーソルをデフォルトにする
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button2.Enabled = true;
                return;
            }

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (j == 7 || j == 8 || j == 9 || j == 10)
                    {
                        if (dt.Rows[i][j].ToString() == "")
                        {
                            ls[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                        }
                        else
                        {
                            ls[i + 1, j + 1].Value = Convert.ToDateTime(dt.Rows[i][j]).ToString("yyyy年MM月dd日");

                        }

                    }
                    else
                    {
                        ls[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                    }

                }
            }

            //出勤簿テンプレートシート
            XLSheet ws = c1XLBook1.Sheets["更新案内"];

            //DateTime dtime = DateTime.Parse(comboBox1.SelectedItem + "/01");
            //int MM = dtime.Month;
            ////YYYY年M月分 (M月給与)
            ws[0, 7].Value = DateTime.Now.ToString("yyyy年MM月dd日");

            ////日付設定
            //for (int i = 0; i <= 30; i++)
            //{
            //    if (dtime.AddDays(i).Month == MM)
            //    {
            //        ws[i + 11, 20].Value = dtime.AddDays(i).ToString();
            //    }
            //}

            for (int i = 1; i <= rows; i++)
            {
                XLSheet newSheet = ws.Clone();
                newSheet.Name = i.ToString(); // クローンをリネーム
                newSheet[0, 0].Value = i;      // 値の変更
                c1XLBook1.Sheets.Add(newSheet);       // クローンをブックに追加
            }

            // テンプレートシートを削除
            c1XLBook1.Sheets.Remove("更新案内");

            string localPass = @"C:\ODIS\TsuuKin\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            c1XLBook1.Save(exlName + ".xlsx");

            msg.Text = "まだぐるぐるします。";

            //Excel Change PDF           
            Microsoft.Office.Interop.Excel.Application m_MyExcel = new Microsoft.Office.Interop.Excel.Application();  //エクセルオブジェクト
            m_MyExcel.Visible = false; //エクセルを非表示
            m_MyExcel.DisplayAlerts = false; //アラート非表示
            Microsoft.Office.Interop.Excel.Workbook m_MyBook; //ブックオブジェクト
                                                              //Microsoft.Office.Interop.Excel.Worksheet m_MySheet; //シートオブジェクト


            //ブックを開く
            m_MyBook = m_MyExcel.Workbooks.Open(Filename: exlName + ".xlsx");

            msg.Text = "もう少しですが、こっから長いです。。";

            //PDF保存
            m_MyBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, exlName + ".pdf");


            m_MyBook.Close(false);
            m_MyExcel.Quit();


            //excel出力
            //System.Diagnostics.Process.Start(@"c:\temp\test2.xlsx");
            //PDF出力
            System.Diagnostics.Process.Start(exlName + ".pdf");

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;

            msg.Text = "";
        }

        private void tuukinhouhou_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tuukinhouhou.SelectedItem.ToString() == "1 車")
            {
                lbl_menkyo.Enabled = true;
                lbl_syaken.Enabled = true;
                lbl_zibai.Enabled = false;
                lbl_ninni.Enabled = true;
            }
            else if (tuukinhouhou.SelectedItem.ToString() == "2 バイク")
            {
                lbl_menkyo.Enabled = true;
                lbl_syaken.Enabled = false;
                lbl_zibai.Enabled = true;
                lbl_ninni.Enabled = true;
            }
            else if (tuukinhouhou.SelectedItem.ToString() == "7 業務車両")
            {
                lbl_menkyo.Enabled = true;
                lbl_syaken.Enabled = false;
                lbl_zibai.Enabled = false;
                lbl_ninni.Enabled = false;
            }
            else if (tuukinhouhou.SelectedItem.ToString() == "9 自転車")
            {
                lbl_menkyo.Enabled = false;
                lbl_syaken.Enabled = false;
                lbl_zibai.Enabled = false;
                lbl_ninni.Enabled = true;
            }
            //else if (tuukinhouhou.SelectedItem.ToString() == "9 その他" || tuukinhouhou.SelectedItem.ToString() == "")
            //{
            //    lbl_menkyo.Enabled = true;
            //    lbl_syaken.Enabled = true;
            //    lbl_zibai.Enabled = true;
            //    lbl_ninni.Enabled = true;
            //}
            else
            {
                lbl_menkyo.Enabled = false;
                lbl_syaken.Enabled = false;
                lbl_zibai.Enabled = false;
                lbl_ninni.Enabled = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //カーソル変更
            Cursor.Current = Cursors.WaitCursor;

            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                GetData();
            }
        }
    }
}
