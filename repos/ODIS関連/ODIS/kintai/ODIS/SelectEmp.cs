using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SelectEmp : Form
    {
        private string[] argumentValues; //Form1から受け取った引数
        public string[] ReturnValue;       //Form1に返す戻り値

        /// <summary>
        /// 従業員全データ
        /// </summary>
        private DataTable dt = new DataTable();

        public SelectEmp(params string[] argumentValues)
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

            this.ReturnValue = new string[] { label1.Text, label2.Text, label3.Text, soshikicd.Text, label4.Text, genbacd.Text, label5.Text, label6.Text, label7.Text };

            this.Close();
        }

        static public string[] ShowMiniForm(string s)
        {
            SelectEmp f = new SelectEmp(s);
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
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
                result = " where " + result;
            }

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "select 社員番号, 漢字氏名, 地区名, 組織CD, 組織名, 現場CD, 現場名, 年齢, 在籍年月 from dbo.従業員情報_期間指定検索('" + argumentValues[0] + "')" + result;

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dt.Clear();
                GetData();
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            if (checkBox1.Checked)
            { 
                label1.Text = "";
                label2.Text = "";
                label3.Text = drv[2].ToString();
                soshikicd.Text = drv[3].ToString();
                label4.Text = drv[4].ToString();
                genbacd.Text = drv[5].ToString();
                label5.Text = drv[6].ToString();
                label6.Text = "";
                label7.Text = "";
            }
            else
            {
                label1.Text = drv[0].ToString();
                label2.Text = drv[1].ToString();
                label3.Text = drv[2].ToString();
                soshikicd.Text = drv[3].ToString();
                label4.Text = drv[4].ToString();
                genbacd.Text = drv[5].ToString();
                label5.Text = drv[6].ToString();
                label6.Text = drv[7].ToString();
                label7.Text = drv[8].ToString();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                label1.Text = "";
                label2.Text = "";
                label6.Text = "";
                label7.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dt.Clear();
            GetData();
        }
    }
}
