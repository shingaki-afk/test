using ODIS;
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
    public partial class CheckList : Form
    {
        private TargetDays td = new TargetDays();
        private DataTable dt = new DataTable();

        //諸単価、諸手当、社保関連
        private string type = "";

        //諸単価用 北部、八重山、施設、エンジ、他
        private string syubetu = "";

        private string tiku = "";

        //諸単価 取得する処理年月
        private string year = "";
        private string month = "";

        //諸単価 タイトル
        private string title = "";

        public CheckList()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //TODO 宮古島、久米島未対応
            if (Program.dispZinzi == 7)
            {
                tiku = "八重山";
            }
            else if (Program.dispZinzi == 8)
            {
                tiku = "北部";
            }
            else if (Program.dispZinzi == 99)
            {
                tiku = "";
            }

            label1.Text = td.StartYMD.ToString("yyyy/MM/dd") + "～" + td.EndYMD.ToString("yyyy/MM/dd");

            if (getsyotanka())
            {
                //当月分あればボタン権限反映
                tankabtn();

                year = td.StartYMD.AddMonths(1).ToString("yyyy");
                month = td.StartYMD.AddMonths(1).ToString("MM");

                //総務権限→人事権限
                if (Program.dispZinzi > 0)
                { 
                    button1.Enabled = true;
                }
                button3.Enabled = true;

            }
            else
            {
                //当月まだであればボタン権限はそのまま全て無効
                label2.Text = td.StartYMD.AddMonths(1).ToString("yyyy年MM月分の給与計算がまだです。") + Environment.NewLine;
                label2.Text += "「諸手当控除チェックリスト」と「雇用体系変更者」は、給与計算後に使用できます。";
            }

            //comboBox1.Items.Add("【情報】休業手当");
            //comboBox1.Items.Add("【情報】休業手当_組織現場集計");
            ////comboBox1.Items.Add("【情報】休業申請_雇保加入");
            ////comboBox1.Items.Add("【情報】休業申請_雇保未加入");
            ////comboBox1.Items.Add("【情報】雇用保険と所定労働日数");
            //comboBox1.Items.Add("【情報】休校に伴う休業一覧");

        }



        //権限で諸単価ボタンの有効無効を設定
        private void tankabtn()
        {
            switch (Program.dispZinzi)
            {
                case 1:　//その他
                    break;
                case 2:
                    break;
                case 3: //PPP/PFI・OIC
                    button10.Enabled = true;
                    button11.Enabled = true;
                    break;
                case 4: //現業・客室
                    button8.Enabled = true;
                    button9.Enabled = true;
                    button15.Enabled = true;
                    button17.Enabled = true;
                    break;
                case 5: //エンジ・施設
                    button12.Enabled = true;
                    button13.Enabled = true;
                    button15.Enabled = true;
                    button17.Enabled = true;
                    break;
                case 6:
                    break;
                case 7: //八重山
                    button7.Enabled = true;
                    if (Program.loginname == "石川　尚吾") button14.Enabled = true;
                    break;
                case 8: //北部
                    button6.Enabled = true;
                    break;
                case 9: //OIC
                    button10.Enabled = true;
                    break;
                case 10: //宮古
                    button15.Enabled = true;
                    break;
                case 11: //久米島
                    button17.Enabled = true;
                    break;
                case 99: //役員・管理者
                    button6.Enabled = true;
                    button7.Enabled = true;
                    button8.Enabled = true;
                    button9.Enabled = true;
                    button10.Enabled = true;
                    button11.Enabled = true;
                    button12.Enabled = true;
                    button13.Enabled = true;
                    button14.Enabled = true;
                    button15.Enabled = true;
                    button17.Enabled = true;
                    break;
            }
        }




        //当月給与計算処理が完了しているかどうか
        private bool getsyotanka()
        {
            return true;
        }


        //諸手当チェックリスト
        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            GetInit();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[s諸手当控除チェックリスト]";

                        Cmd.CommandTimeout = 600;

                        Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                        Cmd.Parameters["year"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["year"].Value = td.StartYMD.AddMonths(1).ToString("yyyy");

                        Cmd.Parameters.Add(new SqlParameter("month", SqlDbType.VarChar));
                        Cmd.Parameters["month"].Direction = ParameterDirection.Input;
                        
                        Cmd.Parameters["month"].Value = td.StartYMD.AddMonths(1).ToString("MM");

                        Cmd.Parameters.Add(new SqlParameter("tiku", SqlDbType.VarChar));
                        Cmd.Parameters["tiku"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["tiku"].Value = tiku;

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
            type = "諸手当";
            syubetu = tiku == "" ? "全地区" : tiku;
            title = "諸手当_控除チェックリスト";

            dgwstyle(type);

            Com.InHistory("40_チェックリスト", "", "");
        }

        //社保関連 給与区分変更者
        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            GetInit();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[k給与区分変更者社保チェックリスト]";

                        Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                        Cmd.Parameters["year"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["year"].Value = td.StartYMD.AddMonths(1).ToString("yyyy");

                        Cmd.Parameters.Add(new SqlParameter("month", SqlDbType.VarChar));
                        Cmd.Parameters["month"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["month"].Value = td.StartYMD.AddMonths(1).ToString("MM");

                        Cmd.Parameters.Add(new SqlParameter("year2", SqlDbType.VarChar));
                        Cmd.Parameters["year2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["year2"].Value = td.StartYMD.ToString("yyyy");

                        Cmd.Parameters.Add(new SqlParameter("month2", SqlDbType.VarChar));
                        Cmd.Parameters["month2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["month2"].Value = td.StartYMD.ToString("MM"); ;

                        Cmd.Parameters.Add(new SqlParameter("tiku", SqlDbType.VarChar));
                        Cmd.Parameters["tiku"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["tiku"].Value = tiku ;

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
            type = "社保関連";
            syubetu = tiku == "" ? "全地区" : tiku;
            title = "【" + td.StartYMD.AddMonths(1).ToString("yyyy年MM月度") + "】給与区分変更者リスト";

            dgwstyle(type);
            Com.InHistory("給与区分変更者社保チェックリスト", "", "");
        }

        //社保関連 雇用保険新規加入者
        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            GetInit();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[k雇用保険新規加入者チェックリスト]";

                        Cmd.Parameters.Add(new SqlParameter("start", SqlDbType.VarChar));
                        Cmd.Parameters["start"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["start"].Value = td.StartYMD.ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("end", SqlDbType.VarChar));
                        Cmd.Parameters["end"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["end"].Value = td.EndYMD.ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("tiku", SqlDbType.VarChar));
                        Cmd.Parameters["tiku"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["tiku"].Value = tiku;

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
            type = "社保関連";
            syubetu = tiku == "" ? "全地区" : tiku;
            title = "【" + td.StartYMD.AddMonths(1).ToString("yyyy年MM月度") + "】雇用保険新規加入者リスト";

            dgwstyle(type);

            Com.InHistory("雇用保険新規加入予定者", "", "");
        }

        //社保関連 65才70才75才到達
        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            GetInit();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[65才70才75才到達リスト]";


                        Cmd.Parameters.Add(new SqlParameter("start", SqlDbType.VarChar));
                        Cmd.Parameters["start"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["start"].Value = td.StartYMD.AddYears(-65).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("end", SqlDbType.VarChar));
                        Cmd.Parameters["end"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["end"].Value = td.EndYMD.AddYears(-65).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("start2", SqlDbType.VarChar));
                        Cmd.Parameters["start2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["start2"].Value = td.StartYMD.AddYears(-70).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("end2", SqlDbType.VarChar));
                        Cmd.Parameters["end2"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["end2"].Value = td.EndYMD.AddYears(-70).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("start3", SqlDbType.VarChar));
                        Cmd.Parameters["start3"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["start3"].Value = td.StartYMD.AddYears(-75).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("end3", SqlDbType.VarChar));
                        Cmd.Parameters["end3"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["end3"].Value = td.EndYMD.AddYears(-75).AddMonths(2).ToString("yyyy/MM/dd");

                        Cmd.Parameters.Add(new SqlParameter("tiku", SqlDbType.VarChar));
                        Cmd.Parameters["tiku"].Direction = ParameterDirection.Input;
                        Cmd.Parameters["tiku"].Value = tiku;

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
            type = "社保関連";
            syubetu = tiku==""?"全地区":tiku;
            title = "【" + td.StartYMD.AddMonths(1).ToString("yyyy年MM月度") + "】65才70才75才到達リスト";

            dgwstyle(type);

            Com.InHistory("65才70才75才到達リスト", "", "");
        }

        //諸単価
        private void button6_Click(object sender, EventArgs e)
        {
            GetInit();

            syubetu = ((Button)sender).Text;

            //ALL
            //OIC
            string selectstr = "select 社員番号, 氏名, 組織名, 現場名, 区分, 時給, 残業, 深夜, 所休, 法休, 延長 from dbo.s諸単価一覧";
            string orderstr = " order by 組織CD, 現場CD, 給与支給区分, 役職CD, カナ名";
            if (syubetu == "ALL")
            {
                dt = Com.GetDB(selectstr + " where 役職CD > '0130' " + orderstr);
            }
            else if (syubetu == "OIC")
            {
                dt = Com.GetDB(selectstr + " where 役職CD > '0130' and 現場名 = 'ＪＩＣＡ沖縄センター'" + orderstr);
            }
            else if (syubetu == "エンジ")
            {
                dt = Com.GetDB(selectstr + " where 担当区分 = '01_エンジ' " + orderstr);
            }
            else
            { 
                dt = Com.GetDB(selectstr + " where 役職CD > '0130' and 担当区分 = '" + syubetu + "'" + orderstr);
            }

            dataGridView1.DataSource = dt;

            type = "諸単価";
            dgwstyle(type);
        }

        //データ取得前処理
        private void GetInit()
        {
            dt.Clear();
            dt.Columns.Clear();
            dataGridView1.DataSource = null;

            //ソート解除 ※エラー防止対応
            BindingSource bs = new BindingSource();
            bs.DataSource = dt;
            bs.Sort = "";
        }

        //グリッドビューの表示設定
        private void dgwstyle(string type)
        {
            if (type == "諸単価")
            {
                dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 130;
                dataGridView1.Columns[2].Width = 120;
                dataGridView1.Columns[3].Width = 250;
                dataGridView1.Columns[4].Width = 100;

                dataGridView1.Columns[5].Width = 70;
                dataGridView1.Columns[6].Width = 70;
                dataGridView1.Columns[7].Width = 70;
                dataGridView1.Columns[8].Width = 70;
                dataGridView1.Columns[9].Width = 70;
                dataGridView1.Columns[10].Width = 70;
            }
            else if (type == "社保関連")
            {
                //dataGridView1.Columns[5].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
                //dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 120;
                dataGridView1.Columns[2].Width = 85;
                dataGridView1.Columns[3].Width = 95;
                dataGridView1.Columns[4].Width = 185;
                dataGridView1.Columns[5].Width = 140;
                dataGridView1.Columns[6].Width = 60;
                dataGridView1.Columns[7].Width = 115;
                dataGridView1.Columns[8].Width = 130;
                dataGridView1.Columns[9].Width = 65;
                dataGridView1.Columns[10].Width = 65;
                dataGridView1.Columns[11].Width = 30;
                dataGridView1.Columns[12].Width = 30;
            }
            else if (type == "諸手当")
            {
                dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[11].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[12].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[13].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[14].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[15].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[16].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 120;
                dataGridView1.Columns[2].Width = 100;
                dataGridView1.Columns[3].Width = 90;
                dataGridView1.Columns[4].Width = 120;
                dataGridView1.Columns[5].Width = 90;
                dataGridView1.Columns[6].Width = 190;

                dataGridView1.Columns[7].Width = 65;
                dataGridView1.Columns[8].Width = 65;
                dataGridView1.Columns[9].Width = 65;
                dataGridView1.Columns[10].Width = 65;
                dataGridView1.Columns[11].Width = 65;
                dataGridView1.Columns[12].Width = 65;
                dataGridView1.Columns[13].Width = 65;
                dataGridView1.Columns[14].Width = 65;
                dataGridView1.Columns[15].Width = 65;
                dataGridView1.Columns[16].Width = 65;
            }
            else if (type == "退職者")
            {
                dataGridView1.Columns[6].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[7].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 150;
                dataGridView1.Columns[2].Width = 150;
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].Width = 60;
                dataGridView1.Columns[5].Width = 60;
                dataGridView1.Columns[6].Width = 100;
                dataGridView1.Columns[7].Width = 100;
                dataGridView1.Columns[8].Width = 100;
                dataGridView1.Columns[9].Width = 100;
            }
        }

        //Excel出力ボタン
        private void button2_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("出力対象を選んでください。");
                return;
            }

            //ボタン無効化・カーソル変更
            button2.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            string fileName = "";
            string printwide = "";
            
            //印刷範囲の調整
            int adj = 0;
            int adj2 = 0;

            if (type == "諸単価")
            {
                fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\01_諸単価一覧表.xls";
                printwide = "L";
                adj2 = 40;
            }
            else if (type == "諸手当")
            {
                fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\02_諸手当控除チェックリストKJ.xls";
                printwide = "R";
                adj = 5;
                adj2 = 40;
            }
            else if (type == "社保関連")
            {
                fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\03_社保関連チェックリスト.xls";
                printwide = "N";
                adj = 5;
                adj2 = 18;
            }
            else if (type == "退職者")
            {
                fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\04_退職者チェック.xls";
                printwide = "J";
                adj = 5;
                adj2 = 27;
            }
            else
            {
                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button2.Enabled = true;

                return;
            }

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

            m_MySheet.Cells[3, 1] = "【" + syubetu + "】" + title;
            m_MySheet.Cells[3, 7] = DateTime.Today.ToString("作成日: yyyy/M/d");


            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;


            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (type == "諸手当")
                    {
                        m_MySheet.Cells[i + 6, j + 2] = dt.Rows[i][j].ToString();
                    }
                    else
                    {
                        m_MySheet.Cells[i + 6, j + 2] = dt.Rows[i][j].ToString().Replace("→", "\r\n" + "→").Replace(")　", ")　" + "\r\n").Replace("（那指）", "").Replace("（那指）", "").Replace("（八）", "").Replace("（北）", "").Replace("（那）", "");
                    }
                }
            }


            for (int i = 0; i < 2000; i += adj2)
            {
                if (rows <= i)
                {
                    m_MySheet.PageSetup.PrintArea = @"$A$1:$" + printwide + "$" + (i + adj).ToString();
                    break;
                }
            }

            string localPass = @"C:\ODIS\CHECKLIST\";
            string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_") + title;
            //string exlName = localPass + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒");

            //フォルダがなければ作成する
            if (!System.IO.File.Exists(localPass))
            {
                System.IO.Directory.CreateDirectory(localPass);
            }

            //excel保存 ローカルへ
            m_MyBook.SaveAs(exlName + ".xls");

            m_MyBook.Close(false);
            m_MyExcel.Quit();

            //excel出力
            System.Diagnostics.Process.Start(exlName + ".xls");


            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;

            //this.TopMost = false;

            //エクセルを表示
            //m_MyExcel.Visible = true;

            //Microsoft.VisualBasic.Interaction.AppActivate(fileName);
        }

        //退職者チェック
        private void button16_Click(object sender, EventArgs e)
        {
            //DataTable dt = new DataTable();
            GetInit();
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
                        //Cmd.CommandText = "select * from dbo.t退職者控除関連チェック where 地区名 like '%" + tiku + "%' or 地区名 like '%" + tiku2 + "%'";
                        Cmd.CommandText = "select * from dbo.t退職者控除関連チェック where 地区名 like '%" + tiku + "%' ";
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

            DataTable Disp = new DataTable();
            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("氏名", typeof(string));
            Disp.Columns.Add("組織", typeof(string));
            Disp.Columns.Add("退職年月日", typeof(string));
            Disp.Columns.Add("雇用", typeof(string));
            Disp.Columns.Add("社保", typeof(string));
            Disp.Columns.Add("住民税当月控除額", typeof(int));
            Disp.Columns.Add("住民税一括徴収額", typeof(int));
            Disp.Columns.Add("住民税控除合計額", typeof(int));
            Disp.Columns.Add("在籍年月", typeof(string));

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = Disp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["氏名"] = row["氏名"];
                nr["組織"] = row["組織名"];
                nr["退職年月日"] = row["退職年月日"];
                nr["雇用"] = row["雇用"];
                nr["社保"] = row["社保"];
                nr["住民税当月控除額"] = Convert.ToInt32(row["住民税当月控除額"]);
                nr["住民税一括徴収額"] = Convert.ToInt32(row["住民税一括徴収額"]);
                nr["住民税控除合計額"] = Convert.ToInt32(row["住民税控除合計額"]);
                nr["在籍年月"] = row["在籍年月"];
                Disp.Rows.Add(nr);
            }

            dataGridView1.DataSource = Disp;
            title = "【" + td.StartYMD.AddMonths(1).ToString("yyyy年MM月度") + "】退職者チェックリスト";
            type = "退職者";
            dgwstyle(type);

            Com.InHistory("退職者チェックリスト", "", "");
        }


        private void button15_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComboData();
        }

        private void GetComboData()
        {
            ////コンボボックス無効化・カーソル変更
            //comboBox1.Enabled = false;
            //Cursor.Current = Cursors.WaitCursor;

            //dt.Clear();
            //dt.Columns.Clear();
            //dataGridView1.DataSource = null;


            //switch (comboBox1.SelectedItem.ToString())
            //{
            //    //case "【情報】雇用保険と所定労働日数":
            //    //    dt = Com.GetDB("select * from dbo.k雇用保険と所定労働日数");
            //    //    break;
            //    case "【情報】休業手当":
            //        string sql = "select * from dbo.k休業手当防止額 ";
            //        dt = Com.GetDB(sql + "order by 組織CD, 現場CD");
            //        break;
            //    case "【情報】休業手当_組織現場集計":
            //        dt = Com.GetDB("select * from dbo.k休業手当表示_組織別現場別");
            //        break;
            //    //case "【情報】休業申請_雇保加入":
            //    //    dt = Com.GetDB("select * from dbo.休業手当_雇用保険加入 order by 組織CD, 現場CD, カナ氏名");
            //    //    break;
            //    //case "【情報】休業申請_雇保未加入":
            //    //    dt = Com.GetDB("select * from dbo.休業手当_雇用保険未加入 order by 組織CD, 現場CD, カナ氏名");
            //    //    break;
            //    case "【情報】休校に伴う休業一覧":
            //        dt = Com.GetDB("select * from dbo.k休校に伴う休業 order by 組織CD, 現場CD");
            //        break;
            //}


            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            //dataGridView1.DataSource = dt;

            //DispChange();

            ////全て入力した後に列幅を自動調節する
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            ////カーソル変更・メッセージキュー処理・コンボボックス有効化
            //Cursor.Current = Cursors.Default;
            //Application.DoEvents();
            //comboBox1.Enabled = true;

            //Com.InHistory(comboBox1.SelectedItem.ToString(), "", "");
        }

        private void DispChange()
        {
            //if (dt.Rows.Count == 0) return;

            //if (comboBox1.SelectedItem.ToString() == "【情報】休業手当_組織現場集計")
            //{
            //    dataGridView1.Columns[4].DefaultCellStyle.Format = "#,0";
            //    dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    //for (int i = 2; i < 5; i++)
            //    //{
            //    //    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
            //    //    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    //}

            //}
            //else if (comboBox1.SelectedItem.ToString() == "【情報】休業手当")
            //{
            //    dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            //    dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            //    //for (int i = 8; i < 31; i++)
            //    //{
            //    //    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
            //    //    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    //}

            //}



        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button18_Click(object sender, EventArgs e)
        {
            GetInit();

            dt = Com.GetDB("select 社員番号, 氏名, 組織名, 現場名, 区分, 時給, 残業, 深夜, 所休, 法休, 延長 from dbo.s諸単価一覧_施設用  order by 組織CD, 現場CD, 給与支給区分, 役職CD, カナ名");

            dataGridView1.DataSource = dt;

            type = "諸単価";
            dgwstyle(type);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            GetInit();

            dt = Com.GetDB("select 社員番号, 氏名, 組織名, 現場名, 区分, 時給, 残業, 深夜, 所休, 法休, 延長 from dbo.s諸単価一覧_エンジ用  order by 組織CD, 現場CD, 給与支給区分, 役職CD, カナ名");

            dataGridView1.DataSource = dt;

            type = "諸単価";
            dgwstyle(type);
        }
    }
}
