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
    public partial class Soushikyuu : Form
    {
        //private string sDate = ""; 
        //private string eDate = "";

        public Soushikyuu()
        {

            if (Convert.ToInt16(Program.access) < 8)
            {
                MessageBox.Show("参照権限がありません。");
                Com.InHistory("49_年度別総支給額_権限制限", "", "");
                return;
            }

            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);
            dataGridView2.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //境界線
            splitContainer1.SplitterDistance = 220;
            //splitContainer1.SplitterDistance = 600;

            for (int i = 2012; i < 2025; i++)
            {
                checkedListBox1.Items.Add(i);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }

            //checkBox1.Checked = true;

            GetData();

            Com.InHistory("49_年度別総支給額", "", "");
        }

        private void GetData()
        {
            Cursor.Current = Cursors.WaitCursor;

            //sDate = yms.Value.ToString("yyyyMM");
            //eDate = yme.Value.ToString("yyyyMM");

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

            ////先頭が「and」の場合、削除する
            //if (result.StartsWith(" and"))
            //{
            //    result = result.Remove(0, 4);
            //}

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) result += " and 年度 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            if (checkBox1.Checked)
            {
                result += " and 退職年月日 is null "; 
            }

            DataTable dt = new DataTable();
            dt = Com.GetDB("select a.*, b.退職年月日 from dbo.s支給額_賞与込 a left join dbo.s社員基本情報 b on a.社員番号 = b.社員番号 where reskey like '%%' " + result + " order by 年度");

            dataGridView1.DataSource = dt;

            string sql2 = "";
            sql2 += "select Convert(varchar, 年度) as 年度 ";
            sql2 += ",sum(労働時間) as 労働時間,sum(基本給) as 基本給,sum(固定手当) as 固定手当,sum(変動手当) as 変動手当,sum(給与合計) as 給与合計 ";
            sql2 += ",sum(夏期賞与) as 夏期賞与,sum(冬期賞与) as 冬期賞与,sum(期末賞与) as 期末賞与,sum(賞与合計) as 賞与合計,sum(総支給額) as 総支給額 ";
            sql2 += "from dbo.s支給額_賞与込 a left join dbo.s社員基本情報 b on a.社員番号 = b.社員番号 ";
            sql2 += "where reskey like '%%' " + result ;
            sql2 += " group by 年度 union all select '合計' as 年度 ";
            sql2 += ",sum(労働時間) as 労働時間,sum(基本給) as 基本給,sum(固定手当) as 固定手当,sum(変動手当) as 変動手当,sum(給与合計) as 給与合計 ";
            sql2 += ",sum(夏期賞与) as 夏期賞与,sum(冬期賞与) as 冬期賞与,sum(期末賞与) as 期末賞与,sum(賞与合計) as 賞与合計,sum(総支給額) as 総支給額 ";
            sql2 += "from dbo.s支給額_賞与込 a left join dbo.s社員基本情報 b on a.社員番号 = b.社員番号 ";
            sql2 += "where reskey like '%%' " + result + " order by 年度";

            DataTable dt2 = new DataTable();
            dt2 = Com.GetDB(sql2);

            dataGridView2.DataSource = dt2;

            //0社員番号
            //1氏名
            //2役職名
            //3支給区分
            //4地区名
            //5組織名
            //6現場名
            //7開始年月
            //8最終年月

            //9労働時間平均
            //10月数
            //11固定平均
            //12変動平均
            //13支給平均
            //14reskey

            //右寄左寄
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView1.Columns[8].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[11].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[12].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[13].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[14].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[15].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[16].DefaultCellStyle.Format = "#,##0";
            dataGridView1.Columns[17].DefaultCellStyle.Format = "#,##0";

            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[2].Width = 120;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 90;
            dataGridView1.Columns[5].Width = 90;
            dataGridView1.Columns[6].Width = 90;
            dataGridView1.Columns[7].Width = 120;

            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].Width = 80;
            dataGridView1.Columns[14].Width = 80;
            dataGridView1.Columns[15].Width = 80;
            dataGridView1.Columns[16].Width = 80;
            dataGridView1.Columns[17].Width = 80;

            dataGridView1.Columns[18].Visible = false; //reskey
            dataGridView1.Columns[19].Visible = false; //reskey

            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView2.Columns[1].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[2].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[3].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[4].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[5].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[6].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[7].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[8].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[9].DefaultCellStyle.Format = "#,##0";
            dataGridView2.Columns[10].DefaultCellStyle.Format = "#,##0";

            dataGridView2.Columns[0].Width = 60;
            dataGridView2.Columns[1].Width = 160;
            dataGridView2.Columns[2].Width = 160;
            dataGridView2.Columns[3].Width = 160;
            dataGridView2.Columns[4].Width = 160;
            dataGridView2.Columns[5].Width = 160;
            dataGridView2.Columns[6].Width = 160;
            dataGridView2.Columns[7].Width = 160;
            dataGridView2.Columns[8].Width = 160;
            dataGridView2.Columns[9].Width = 160;
            dataGridView2.Columns[10].Width = 160;


            Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void yms_ValueChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void yme_ValueChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //GetData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

           
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) GetData();
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) GetData();
        }

        private void label23_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }
        }
    }
}
