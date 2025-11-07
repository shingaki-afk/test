using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class RingiMini : Form
    {
        private string[] argumentValues; //Form1から受け取った引数
        public string[] ReturnValue;       //Form1に返す戻り値

        /// <summary>
        /// 従業員全データ
        /// </summary>
        private DataTable dt = new DataTable();

        //TODO 毎年変更が必要!!!!  会社今期　
        private string genzaiki = "49";

        public RingiMini(params string[] argumentValues)
        {
            //Form1から受け取ったデータをForm2インスタンスのメンバに格納
            this.argumentValues = argumentValues;
            InitializeComponent();
        }

        static public string[] ShowMiniForm(string[] s)
        {
            RingiMini f = new RingiMini(s);
            f.ShowDialog();
            string[] receiveText = f.ReturnValue;
            f.Dispose();

            return receiveText;
        }

        private void SelectEmp_Load(object sender, EventArgs e)
        {
            //Form1から送られてきたテキストをForm2で表示
            //this.ReceiveTextBox.Text = argumentValues[0];

            //画面を最大化、最小化、閉じる、のみに設定
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            //I : 新規登録
            //S : 修正
            //K : 決裁
            if (argumentValues[0] == "I")
            {
                this.Text = "新規登録画面";

                //フォームの高さ設定
                this.Height = 256;

                label13.Visible = false;
                textBox8.Visible = false;

                //新規登録時
                comboBox1.Items.Add("");
                comboBox1.Items.Add("01 講習・研修");
                comboBox1.Items.Add("02 募集");
                comboBox1.Items.Add("03 資機材");
                comboBox1.Items.Add("04 車輌・電子機器");
                comboBox1.Items.Add("05 ユニフォーム・備品");
                comboBox1.Items.Add("06 退職・手当金");
                comboBox1.Items.Add("07 誤計算");
                comboBox1.Items.Add("08 接待交際");
                comboBox1.Items.Add("09 宣伝広告協賛");
                comboBox1.Items.Add("10 その他");
                comboBox1.SelectedIndex = 0;

                textBox3.Text = Program.loginname;
                //textBox4.Text = Program.syainno;
                textBox5.Text = Program.tiku;
                textBox6.Text = Program.soshiki;
                textBox7.Text = DateTime.Now.ToString("yyyy/MM/dd");

            }
            else if (argumentValues[0] == "S") //修正時
            {
                this.Text = "修正画面";
                //フォームの高さ設定
                this.Height = 256;

                //TODO
                comboBox1.Items.Add("");
                comboBox1.Items.Add("01 講習・研修");
                comboBox1.Items.Add("02 募集");
                comboBox1.Items.Add("03 資機材");
                comboBox1.Items.Add("04 車輌・電子機器");
                comboBox1.Items.Add("05 ユニフォーム・備品");
                comboBox1.Items.Add("06 退職・手当金");
                comboBox1.Items.Add("07 誤計算");
                comboBox1.Items.Add("08 接待交際");
                comboBox1.Items.Add("09 宣伝広告協賛");
                comboBox1.Items.Add("10 その他");
                comboBox1.SelectedIndex = 0;

                //対象のデータを取得
                DataTable dtK = new DataTable();
                dtK = Com.GetDB("select 稟議番号, 地区名, 組織名, 氏名, 登録日, 申請額, 区分, 目的, 結果, コメント, 社員番号 from dbo.稟議データ where 稟議番号 = '" + argumentValues[1] + "'");

                textBox7.Text = dtK.Rows[0]["登録日"].ToString();
                textBox8.Text = dtK.Rows[0]["稟議番号"].ToString();
                textBox3.Text = dtK.Rows[0]["氏名"].ToString();
                textBox4.Text = dtK.Rows[0]["社員番号"].ToString();
                textBox5.Text = dtK.Rows[0]["地区名"].ToString();
                textBox6.Text = dtK.Rows[0]["組織名"].ToString();

                textBox1.Text = dtK.Rows[0]["目的"].ToString();
                numericUpDown1.Value = Convert.ToInt64(dtK.Rows[0]["申請額"]);
                comboBox1.SelectedItem = dtK.Rows[0]["区分"].ToString();

                this.button2.Text = "修正";
            }
            else if (argumentValues[0] == "K") //決裁時
            {
                this.Text = "決裁画面";

                //フォームの高さ設定
                this.Height = 375;

                this.button2.Visible = false;

                comboBox2.Items.Add("");
                comboBox2.Items.Add("承認");
                comboBox2.Items.Add("否認");
                comboBox2.Items.Add("保留");
                comboBox2.Items.Add("取消");

                //対象のデータを取得
                DataTable dtK = new DataTable();
                dtK = Com.GetDB("select 稟議番号, 地区名, 組織名, 氏名, 登録日, 申請額, 区分, 目的, 結果, コメント, 社員番号 from dbo.稟議データ where 稟議番号 = '" + argumentValues[1] + "'");

                textBox7.Text = dtK.Rows[0]["登録日"].ToString();
                textBox8.Text = dtK.Rows[0]["稟議番号"].ToString();
                textBox3.Text = dtK.Rows[0]["氏名"].ToString();
                textBox4.Text = dtK.Rows[0]["社員番号"].ToString();
                textBox5.Text = dtK.Rows[0]["地区名"].ToString();
                textBox6.Text = dtK.Rows[0]["組織名"].ToString();

                textBox1.Text = dtK.Rows[0]["目的"].ToString();
                numericUpDown1.Value = Convert.ToInt64(dtK.Rows[0]["申請額"]);
                comboBox1.SelectedItem = dtK.Rows[0]["区分"].ToString();

                textBox1.Enabled = false;
                textBox1.BackColor = System.Drawing.Color.White;
                textBox1.Font = new System.Drawing.Font(label1.Font, System.Drawing.FontStyle.Bold);

                numericUpDown1.Enabled = false;
                numericUpDown1.BackColor = System.Drawing.Color.White;
                numericUpDown1.Font = new System.Drawing.Font(numericUpDown1.Font, System.Drawing.FontStyle.Bold);

                comboBox1.Enabled = false;
                comboBox1.BackColor = System.Drawing.Color.White;
                comboBox1.Font = new System.Drawing.Font(comboBox1.Font, System.Drawing.FontStyle.Bold);

                textBox2.Text = dtK.Rows[0]["コメント"].ToString();
                comboBox2.SelectedItem = dtK.Rows[0]["結果"].ToString();
            }

            //共通処理
            //桁区切りを表示する
            numericUpDown1.ThousandsSeparator = true;

        }

        //登録ボタン/修正ボタン
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("目的は入力必須です。"); return;
            }

            if (this.button2.Text == "修正")
            {
                DataSyuusei();

                //戻り値をセット
                this.ReturnValue = new string[] { textBox8.Text, textBox1.Text, numericUpDown1.Value.ToString("#,0"), comboBox1.SelectedItem.ToString() };
            }
            else
            {
                //稟議データテーブルへインサート
                //採番
                string str_tiku = "";
                string str_no = "";
                if (Program.tiku == "本社" || Program.tiku == "那覇")
                {
                    str_tiku = "0";
                }
                else if (Program.tiku == "八重山")
                {
                    str_tiku = "1";
                }
                else if (Program.tiku == "北部")
                {
                    str_tiku = "2";
                }
                else
                {
                    MessageBox.Show("おかしーす。システム管理者に問合願います。");
                    return;
                }

                dt = Com.GetDB("select max(right(稟議番号, 3)) + 1 from dbo.稟議データ where left(稟議番号, 4) = '" + genzaiki + "-" + str_tiku + "'");

                str_no = dt.Rows[0][0].ToString();

                if (str_no.Length == 0) str_no = "001";
                if (str_no.Length == 1) str_no = "00" + str_no;
                if (str_no.Length == 2) str_no = "0" + str_no;

                str_no = genzaiki + "-" + str_tiku + str_no;

                DataInsert(str_no);

                //戻り値をセット
                this.ReturnValue = new string[] { str_no, textBox1.Text, numericUpDown1.Value.ToString("#,0"), comboBox1.SelectedItem.ToString() };
            }

            this.Close();
        }

        // データ登録処理
        private void DataInsert(string no)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            string cmd = "";
            cmd = "[dbo].[Insert稟議]";
            
            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = cmd;

                    Cmd.Parameters.Add(new SqlParameter("稟議番号", SqlDbType.VarChar));
                    Cmd.Parameters["稟議番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                    Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("地区名", SqlDbType.VarChar));
                    Cmd.Parameters["地区名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("組織名", SqlDbType.VarChar));
                    Cmd.Parameters["組織名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("氏名", SqlDbType.VarChar));
                    Cmd.Parameters["氏名"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("登録日", SqlDbType.Date));
                    Cmd.Parameters["登録日"].Direction = ParameterDirection.Input;


                    Cmd.Parameters.Add(new SqlParameter("申請額", SqlDbType.Decimal));
                    Cmd.Parameters["申請額"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("目的", SqlDbType.VarChar));
                    Cmd.Parameters["目的"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("結果", SqlDbType.VarChar));
                    Cmd.Parameters["結果"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("コメント", SqlDbType.VarChar));
                    Cmd.Parameters["コメント"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("区分", SqlDbType.VarChar));
                    Cmd.Parameters["区分"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;


                    Cmd.Parameters["稟議番号"].Value = no;
                    Cmd.Parameters["社員番号"].Value = textBox4.Text;
                    Cmd.Parameters["地区名"].Value = textBox5.Text;
                    Cmd.Parameters["組織名"].Value = textBox6.Text;
                    Cmd.Parameters["氏名"].Value = textBox3.Text;
                    Cmd.Parameters["登録日"].Value = DateTime.Now.ToString("yyyy-MM-dd");
                    Cmd.Parameters["申請額"].Value = numericUpDown1.Value;
                    Cmd.Parameters["目的"].Value = textBox1.Text;
                    Cmd.Parameters["結果"].Value = "";
                    Cmd.Parameters["コメント"].Value = "";
                    Cmd.Parameters["区分"].Value = comboBox1.SelectedItem;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        //決裁ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            DataUpdate("");
            //戻り値をセット
            this.ReturnValue = new string[] { textBox4.Text, textBox8.Text, textBox3.Text, textBox1.Text, numericUpDown1.Value.ToString("#,0"), textBox2.Text, comboBox2.SelectedItem.ToString(), Convert.ToDateTime(textBox7.Text).ToString("yyyy/MM/dd") };

            //textBox4.Text = dtK.Rows[0]["社員番号"].ToString();
            //textBox8.Text = dtK.Rows[0]["稟議番号"].ToString();
            //textBox3.Text = dtK.Rows[0]["氏名"].ToString();
            //textBox1.Text = dtK.Rows[0]["目的"].ToString();
            //numericUpDown1.Value = Convert.ToInt64(dtK.Rows[0]["申請額"]);
            //textBox2.Text = dtK.Rows[0]["コメント"].ToString();
            //comboBox2.SelectedItem = dtK.Rows[0]["結果"].ToString();
            this.Close();
        }

        // データ修正
        private void DataSyuusei()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            string cmd = "";
            cmd = "[dbo].[Update修正稟議]";

            //textBox1.Text = dtK.Rows[0]["目的"].ToString();
            //numericUpDown1.Value = Convert.ToInt64(dtK.Rows[0]["申請額"]);
            //comboBox1.SelectedItem = dtK.Rows[0]["区分"].ToString();

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = cmd;

                    Cmd.Parameters.Add(new SqlParameter("稟議番号", SqlDbType.VarChar));
                    Cmd.Parameters["稟議番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("目的", SqlDbType.VarChar));
                    Cmd.Parameters["目的"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("申請額", SqlDbType.Decimal));
                    Cmd.Parameters["申請額"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("区分", SqlDbType.VarChar));
                    Cmd.Parameters["区分"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["稟議番号"].Value = textBox8.Text;
                    Cmd.Parameters["目的"].Value = textBox1.Text;
                    Cmd.Parameters["申請額"].Value = numericUpDown1.Value;
                    Cmd.Parameters["区分"].Value = comboBox1.SelectedItem;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }

        // データ更新処理　決裁
        private void DataUpdate(string no)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable DataTable = new DataTable();
            SqlDataReader dr;

            string cmd = "";
            cmd = "[dbo].[Update決裁稟議]";

            using (Cn = new SqlConnection(Com.SQLConstr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = cmd;

                    Cmd.Parameters.Add(new SqlParameter("稟議番号", SqlDbType.VarChar));
                    Cmd.Parameters["稟議番号"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("結果", SqlDbType.VarChar));
                    Cmd.Parameters["結果"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("コメント", SqlDbType.VarChar));
                    Cmd.Parameters["コメント"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;
                    
                    Cmd.Parameters["稟議番号"].Value = textBox8.Text;
                    Cmd.Parameters["結果"].Value = comboBox2.SelectedItem;
                    Cmd.Parameters["コメント"].Value = textBox2.Text;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                    }
                }
            }
        }
    }
}
