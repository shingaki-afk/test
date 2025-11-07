using ODIS.ODIS;
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
    public partial class shikaku : Form
    {
        //全データ格納テーブル
        private DataTable dt = new DataTable();

        public shikaku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //データ取得
            GetData();

            Com.InHistory("22_資格検索", "", "");
        }

        //データ取得
        private void GetData()
        {
            dt.Clear();
            
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        if (!checkBox1.Checked)
                        {
                            Cmd.CommandText = "select * from dbo.s資格一覧 where 在籍区分 <> '9'";
                        }
                        else
                        { 
                            Cmd.CommandText = "select * from dbo.s資格一覧";
                        }
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ表示
            DataView();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void DataView()
        {
            dataGridView1.DataSource = "";

            DataTable Disp = new DataTable();

            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string conditions = "";

            string result = "";
            if (ar[0] != "")
            {
                conditions = "【含】" + textBox1.Text;

                foreach (string s in ar)
                {
                        result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //test 除き文字列
            string res2 = textBox2.Text.Trim().Replace("　", " ");
            string[] ar2 = res2.Split(' ');

            if (ar2[0] != "")
            {
                conditions += "　【除】" + textBox2.Text;
                foreach (string s in ar2)
                {
                    result += " and (reskey not like '%" + s + "%' or reskey not like '%" + Com.isOneByteChar(s) + "%' or reskey not like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey not like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey not like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey not like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");

            //グリッド表示クリア
            dataGridView1.DataSource = "";

            int ct = 0;

            Disp.Columns.Add("社員番号", typeof(string));
            Disp.Columns.Add("漢字氏名", typeof(string));

            Disp.Columns.Add("資格コード", typeof(string));
            Disp.Columns.Add("資格名", typeof(string));
            Disp.Columns.Add("取得日", typeof(string));
            Disp.Columns.Add("取得番号", typeof(string));
            Disp.Columns.Add("有効期限", typeof(string));
            Disp.Columns.Add("対象", typeof(string));

            Disp.Columns.Add("規程額", typeof(string));
            Disp.Columns.Add("手当額", typeof(string));
            Disp.Columns.Add("免許手当", typeof(string));

            Disp.Columns.Add("地区名", typeof(string));
            Disp.Columns.Add("組織名", typeof(string));
            Disp.Columns.Add("現場名", typeof(string));
            Disp.Columns.Add("役職名", typeof(string));

            Disp.Columns.Add("年齢", typeof(string));
            Disp.Columns.Add("在籍年月", typeof(string));
            Disp.Columns.Add("退職年月日", typeof(string));

            foreach (DataRow row in dtrow)
            {
                DataRow nr = Disp.NewRow();
                nr["社員番号"] = row["社員番号"];
                nr["漢字氏名"] = row["漢字氏名"];
                nr["地区名"] = row["地区名"];
                nr["組織名"] = row["組織名"];
                nr["現場名"] = row["現場名"];
                nr["役職名"] = row["役職名"];

                nr["年齢"] = row["年齢"];
                nr["在籍年月"] = row["在籍年月"];
                nr["退職年月日"] = row["退職年月日"];
                

                nr["資格コード"] = row["資格コード"];
                nr["資格名"] = row["資格名"];
                nr["取得日"] = row["取得日"];
                nr["取得番号"] = row["取得番号"];
                nr["有効期限"] = row["有効期限"];
                nr["対象"] = row["対象"]; 

                nr["規程額"] = row["規程額"];
                nr["手当額"] = row["手当額"];
                nr["免許手当"] = row["免許手当"];

                Disp.Rows.Add(nr);

                ct++;
            }

            //データグリッドビューの高さ指定　※セット前にすること！
            dataGridView1.RowTemplate.Height = 20;

            dataGridView1.DataSource = Disp;

            dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";

            label1.Text = ct.ToString();

            //検索履歴登録
            Com.InHistory("資格", conditions, dtrow.Length.ToString());

            // セル内で文字列を折り返えさない
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

            System.GC.Collect();

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                DataView();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                DataView();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            checkBox1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
            checkBox1.Enabled = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;
            //DataDisp(drv[0].ToString());
        }

        private void DataDisp(string s)
        {
            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.m免許手当合計規程額取得 where 社員番号 = '" + s + "'");
            dataGridView2.DataSource = dt;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //dataGridView1.Columns[8].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns[9].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns[10].DefaultCellStyle.Format = "#,0";
        }
    }
}
