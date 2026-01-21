using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class HardSoft : Form
    {
        public HardSoft()
        {
            InitializeComponent();

            Com.InHistory("66_ハードソフト情報管理","","");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ボタン無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            //データ表示
            GetData();

            //カーソル変更・メッセージキュー処理・ボタン有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void GetData()
        {
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

            string sql = "";
            sql += "select * from dbo.t端末検索 ";

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = "where " + result.Remove(0, 4);
            }

            sql += result;
            sql += " order by 組織CD,現場CD";
            DataTable dt = new DataTable();
            dt = Com.GetDB(sql);

            dt.Columns.Remove("reskey");
            c1FlexGrid1.DataSource = dt;

            //マージ設定
            c1FlexGrid1.Cols[1].AllowMerging = true;
            c1FlexGrid1.Cols[2].AllowMerging = true;
            c1FlexGrid1.Cols[3].AllowMerging = true;
            c1FlexGrid1.Cols[4].AllowMerging = true;
            c1FlexGrid1.Cols[5].AllowMerging = true;
            c1FlexGrid1.Cols[6].AllowMerging = true;
            c1FlexGrid1.Cols[7].AllowMerging = true;
            c1FlexGrid1.Cols[8].AllowMerging = true;
            c1FlexGrid1.Cols[9].AllowMerging = true;
            c1FlexGrid1.Cols[10].AllowMerging = true;
            c1FlexGrid1.Cols[11].AllowMerging = true;
            c1FlexGrid1.Cols[12].AllowMerging = true;
            c1FlexGrid1.Cols[13].AllowMerging = true;
            c1FlexGrid1.Cols[14].AllowMerging = true;
            c1FlexGrid1.Cols[15].AllowMerging = true;
            c1FlexGrid1.Cols[16].AllowMerging = true;
            c1FlexGrid1.Cols[17].AllowMerging = true;
            c1FlexGrid1.Cols[18].AllowMerging = true;
            c1FlexGrid1.Cols[19].AllowMerging = true;
            c1FlexGrid1.Cols[20].AllowMerging = true;

            //選択はしない
            c1FlexGrid1.Select(-1, -1);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //ボタン無効化・カーソル変更
                button1.Enabled = false;
                Cursor.Current = Cursors.WaitCursor;

                //データ処理
                GetData();

                //カーソル変更・メッセージキュー処理・ボタン有効化
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button1.Enabled = true;
            }
        }
    }
}
