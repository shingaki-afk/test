using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SelectIdouDay : Form
    {
        private string[] argumentValues; //Form1から受け取った引数
        public string[] ReturnValue;       //Form1に返す戻り値

        /// <summary>
        /// 従業員全データ
        /// </summary>
        private DataTable dt = new DataTable();

        public SelectIdouDay(params string[] argumentValues)
        {
            //Form1から受け取ったデータをForm2インスタンスのメンバに格納
            this.argumentValues = argumentValues;
            InitializeComponent();
        }

        private void SelectEmp_Load(object sender, EventArgs e)
        {
            //Form1から送られてきたテキストをForm2で表示
            //this.ReceiveTextBox.Text = argumentValues[0];
            //GetData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //戻り値をセット
            //this.ReturnValue = new string[] { label1.Text, label2.Text, soshikicd.Text, label4.Text, genbacd.Text, label5.Text };
            this.ReturnValue = new string[] { Convert.ToDateTime(idoudaynew.Value).ToString("yyyy/MM/dd") };
            this.Close();
        }

        static public string[] ShowMiniForm(string s)
        {
            SelectIdouDay f = new SelectIdouDay(s);
            f.ShowDialog();
            string[] receiveText = f.ReturnValue;
            f.Dispose();

            return receiveText;
        }


    }
}
