using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SelectEmpNyu : Form
    {
        private string[] argumentValues; //Form1から受け取った引数
        public string[] ReturnValue;       //Form1に返す戻り値

        /// <summary>
        /// 従業員全データ
        /// </summary>
        private DataTable dt = new DataTable();

        public SelectEmpNyu(params string[] argumentValues)
        {
            //Form1から受け取ったデータをForm2インスタンスのメンバに格納
            this.argumentValues = argumentValues;
            InitializeComponent();
        }

        private void SelectEmp_Load(object sender, EventArgs e)
        {
            //Form1から送られてきたテキストをForm2で表示
            //this.ReceiveTextBox.Text = argumentValues[0];
            GetData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //戻り値をセット
            //this.ReturnValue = SendTextBox.Text;
            //this.ReturnValue[1] = label1.Text;
            //this.ReturnValue[2] = label2.Text;
            //this.ReturnValue[3] = label3.Text;
            //this.ReturnValue[4] = label4.Text;

            this.ReturnValue = new string[] { label1.Text, label2.Text, soshikicd.Text, label4.Text, genbacd.Text, label5.Text };

            this.Close();
        }

        static public string[] ShowMiniForm(string s)
        {
            SelectEmpNyu f = new SelectEmpNyu(s);
            f.ShowDialog();
            string[] receiveText = f.ReturnValue;
            f.Dispose();

            return receiveText;
        }

        private void GetData()
        {
            string res = textBox2.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Microsoft.VisualBasic.Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //先頭が「and」の場合、「where」にする
            //if (result.StartsWith(" and"))
            //{
            //    result = result.Remove(0, 4);
            //    result = " where " + result;
            //}

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    string sql = "select 社員番号, 漢字氏名, 地区名, 組織名, 現場名 from dbo.従業員情報_期間指定検索('" + argumentValues[0] + "') where ";

                    if (Convert.ToInt16(Program.access) >= 9)
                    {
                        Cmd.CommandText = sql + " 担当区分  <> 'dummy' " + result;
                    }
                    else if (Convert.ToInt16(Program.access) >= 5)
                    {
                        Cmd.CommandText = sql + " 担当区分  <> '21_役員室' " + result;
                    }
                    else
                    {
                        Cmd.CommandText = sql + " 役職CD >= '0120' " + result;
                    }

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Width = 70; 
            dataGridView1.Columns[1].Width = 150; 
            dataGridView1.Columns[2].Width = 70; 
            dataGridView1.Columns[3].Width = 70; 
            dataGridView1.Columns[4].Width = 200;  
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //マウスカーソルを砂時計にする
                Cursor.Current = Cursors.WaitCursor;
                button2.Enabled = false;

                dt.Clear();
                GetData();

                //マウスカーソルをデフォルトにする
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                button2.Enabled = true;
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            label1.Text = drv[0].ToString();
            label2.Text = drv[1].ToString();
            label4.Text = drv[2].ToString();
            label5.Text = drv[3].ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //マウスカーソルを砂時計にする
            Cursor.Current = Cursors.WaitCursor;
            button2.Enabled = false;

            dt.Clear();
            GetData();

            //マウスカーソルをデフォルトにする
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button2.Enabled = true;
        }
    }
}
