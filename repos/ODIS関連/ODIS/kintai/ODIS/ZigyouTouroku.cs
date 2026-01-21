using ODIS;
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
    public partial class ZigyouTouroku : Form
    {
        //public DataTable mainDt;

        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public ZigyouTouroku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 10);
            dataGridView2.Font = new Font(dataGridView2.Font.Name, 10);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //行ヘッダーを非表示にする
            //dataGridView1.RowHeadersVisible = false;

            // セル内で文字列を折り返す
            //dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            //事業登録データ
            GetData();
            this.dataGridView1.Columns["残日数"].ReadOnly = true;
            this.dataGridView1.Columns["登録者"].ReadOnly = true;

            //ビル管データ
            BillData(true);

            Com.InHistory("631_事業登録・ビル管登録一覧", "", "");

            //dataGridView1.DataSource = GetZigyouData();

            ////金額右寄せ
            //dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            ////列幅変更
            //dataGridView1.Columns[0].Width = 40;
            //dataGridView1.Columns[1].Width = 150;
            //dataGridView1.Columns[2].Width = 230;
            //dataGridView1.Columns[3].Width = 60;
            //dataGridView1.Columns[4].Width = 200;
            //dataGridView1.Columns[5].Width = 100;
            //dataGridView1.Columns[6].Width = 80;
            //dataGridView1.Columns[7].Width = 100;
            //dataGridView1.Columns[8].Width = 100;
            //dataGridView1.Columns[9].Width = 100;


            #region 20180318 準備中
            ////フィルター設定
            //c1FlexGrid1.AllowFiltering = true;

            ////自動グリップボード機能を有効にする
            //c1FlexGrid1.AutoClipboard = true;

            //// グリッドのAllowMergingプロパティを設定
            //c1FlexGrid1.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free;

            ////列幅自動調整　バインド前でないといけない
            //c1FlexGrid1.AutoResize = true;

            ////dataGridView1.DataSource = dt;
            //c1FlexGrid1.DataSource = ListDt();

            //for (int i = 0; i < dt.Columns.Count + 1; i++)
            //{
            //    c1FlexGrid1.Cols[i].AllowMerging = true;
            //}

            #endregion
        }

        //
        private void BillData(bool flg)
        {
            DataTable dt = new DataTable();
            if (flg)
            { 
                dt = Com.GetDB("select * from dbo.bビル管登録一覧");
            }
            else
            {
                dt = Com.GetDB("select * from dbo.bビル管登録一覧 where 登録現場 is not null");
            }

            dataGridView2.DataSource = dt;

            dataGridView2.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[12].DefaultCellStyle.Format = "#,0";

            //[列(column),　行(row)]

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    if (dataGridView2[5, i].Value.ToString() == "")
            //    {
            //        dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;//売上
            //    }
            //}
        }

        private void GetData()
        {

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            //タブ文字
            //string tab = " + char(9) + ";
            string tab = " + '    ' + ";

            //string sql = "select 項目, 内容, 社員番号, 氏名, 備考3 as 表示順, 備考2 as 現場, 金額, 備考 from dbo.固定控除 where 内容 = '4_サンエー駐車場代' order by 備考2, 備考3";
            string sql = "select No, 法規制, 名称, 区分, 登録番号, 有効期限, DATEDIFF(day, GETDATE(), 有効期限) as 残日数, 社員番号, ";
            sql += "(select 氏名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 登録者, ";
            sql += "(select case when 在籍区分 = '1' then '' else '【退職】' end + 地区名" + tab + "組織名" + tab + "現場名 from dbo.社員基本情報 where 社員番号 = a.社員番号) as 登録者情報, ";
            sql += "備考, 毎年度報告書提出義務有無 from dbo.事業登録データ a order by No";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        //public DataTable GetZigyouData()
        //{
        //    SqlConnection Cn;
        //    SqlCommand Cmd;
        //    mainDt = new DataTable();
            
        //    SqlDataAdapter da;

        //    using (Cn = new SqlConnection(Common.constr))
        //    {
        //        Cn.Open();

        //        using (Cmd = Cn.CreateCommand())
        //        {
        //            //Cmd.CommandType = CommandType.StoredProcedure;
        //            Cmd.CommandText = "select * from dbo.事業登録";

        //            da = new SqlDataAdapter(Cmd);
        //            da.Fill(mainDt);
        //        }
        //    }

        //    DataTable Disp = new DataTable();

        //    Disp.Columns.Add("No", typeof(int));
        //    Disp.Columns.Add("法規制", typeof(string));
        //    Disp.Columns.Add("名称", typeof(string));
        //    Disp.Columns.Add("区分", typeof(string));
        //    Disp.Columns.Add("登録番号", typeof(string));
        //    Disp.Columns.Add("有効期限", typeof(DateTime));
        //    Disp.Columns.Add("残日数", typeof(Int32));
        //    Disp.Columns.Add("管理部門", typeof(string));
        //    Disp.Columns.Add("管理者", typeof(string));
        //    Disp.Columns.Add("対応状況", typeof(string));

        //    foreach (DataRow row in mainDt.Rows)
        //    {
        //        DataRow nr = Disp.NewRow();
        //        nr["No"] = Convert.ToInt16(row["No"]);
        //        nr["法規制"] = row["法規制"];
        //        nr["名称"] = row["名称"];
        //        nr["区分"] = row["区分"];
        //        nr["登録番号"] = row["登録番号"];

        //        DateTime dt;
        //        TimeSpan ts;
        //        if (DateTime.TryParse(row["有効期限"].ToString(), out dt))
        //        {
        //            nr["有効期限"] = dt;
        //            ts = dt.Subtract(DateTime.Now);
        //            nr["残日数"] = ts.Days;
        //        }
        //        else
        //        {
        //            nr["有効期限"] = new DateTime(0);
        //            nr["残日数"] = DBNull.Value;
        //        }

        //        //DateTime kigen = Convert.ToDateTime(row["有効期限"]);
        //        //TimeSpan ts = DateTime.Now.Subtract(kigen);

        //        //nr["残日数"] = str;
        //        nr["管理部門"] = row["管理部門"];
        //        nr["管理者"] = row["管理者"];
        //        nr["対応状況"] = row["対応フラグ"];
        //        Disp.Rows.Add(nr);
        //    }

        //    return Disp;
        //}

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //セルの列を確認
            decimal val = 0;
            if (e.Value != null && e.ColumnIndex == 6 && decimal.TryParse(e.Value.ToString(), out val))
            {
                //セルの値により、背景色を変更する
                if (val < 100 && val > 30)
                {
                    //有効期限 100日きった。
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                    //e.CellStyle.BackColor = Color.Yellow;
                }
                else if (val <= 30)
                {
                    //有効期限 30日きった。
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                    //e.CellStyle.BackColor = Color.Crimson;
                }
            }
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

            //セルの列を確認
            //decimal val = 0;
            //if (e.Value != null && e.ColumnIndex == 12 && decimal.TryParse(e.Value.ToString(), out val))
            //{
            //    //セルの値により、背景色を変更する
            //    if (val > 0)
            //    {
            //        //有効期限 100日きった。
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Yellow;
            //        //e.CellStyle.BackColor = Color.Yellow;
            //    }
            //    else 
            //    {
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
            //    }
            //}
            //[列(column),　行(row)]
            //dataGridView2[1, e.RowIndex].Value;


            //if (str == "A4売上45"
            //dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PaleGreen; //売上

            if (e.ColumnIndex == 5)
            {
                //MessageBox.Show(e.Value.ToString());
                if (e.Value != null && e.Value.ToString().Length > 0)
                {
                    dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                }
                else
                {
                    dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.CornflowerBlue;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }

            dt.Clear();
            GetData();
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            //新しい行のセルでなく、セルの内容が変更されている時だけ検証する
            if (e.RowIndex == dgv.NewRowIndex || !dgv.IsCurrentCellDirty)
            {
                return;
            }

            if (dgv.Columns[e.ColumnIndex].Name == "有効期限" &&　e.FormattedValue.ToString() == "")
            {
                //行にエラーテキストを設定
                dgv.Rows[e.RowIndex].ErrorText = "値が入力されていません。";
                //入力した値をキャンセルして元に戻すには、次のようにする
                //dgv.CancelEdit();
                //キャンセルする
                e.Cancel = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                //先月情報含めない
                BillData(false);
            }
            else
            {
                //先月情報含める
                BillData(true); ;
            }
        }


    }
}
