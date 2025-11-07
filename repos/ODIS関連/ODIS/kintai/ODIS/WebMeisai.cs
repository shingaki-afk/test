using C1.C1Excel;
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
    public partial class WebMeisai : Form
    {
        private TargetDays td = new TargetDays();

        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();


        public WebMeisai()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);
            dataGridView2.Font = new Font(dataGridView1.Font.Name, 12);

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            for (int i = 2023; i <= Convert.ToInt16(td.StartYMD.AddMonths(1).ToString("yyyy")); i++)
            {
                comboBox1.Items.Add(i.ToString());
                comboBox3.Items.Add(i.ToString());
            }

            comboBox2.Items.Add("01");
            comboBox2.Items.Add("02");
            comboBox2.Items.Add("03");
            comboBox2.Items.Add("04");
            comboBox2.Items.Add("05");
            comboBox2.Items.Add("06");
            comboBox2.Items.Add("07");
            comboBox2.Items.Add("08");
            comboBox2.Items.Add("09");
            comboBox2.Items.Add("10");
            comboBox2.Items.Add("11");
            comboBox2.Items.Add("12");

            comboBox4.Items.Add("01_夏");
            comboBox4.Items.Add("02_冬");
            comboBox4.Items.Add("11_期末");

            comboBox1.SelectedItem = td.StartYMD.AddMonths(1).ToString("yyyy");
            comboBox2.SelectedItem = td.StartYMD.AddMonths(1).ToString("MM");

            comboBox3.SelectedItem = td.StartYMD.AddMonths(1).ToString("yyyy");

            //comboBox3.Items.Add("給与");
            //comboBox3.Items.Add("賞与");

            comboBox4.SelectedIndex = 0;

            GetData();
        }

        private void GetData()
        {
            dataGridView1.DataSource = null;
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();

            string sql = "select * , (select 退職年月日 from dbo.s社員基本情報 b where b.社員番号 = a.社員番号) as 退職年月日 from dbo.web明細対象者 a ";

            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);
            dataGridView1.DataSource = dt;

            label7.Text = dt.Rows.Count.ToString() + "/200";

            //事務所、エンジ、全現場でweb明
            //細未設定
            DataTable dt2 = new DataTable();
            dt2 = Com.GetDB("select a.社員番号, a.氏名, a.組織名, a.現場名, a.生年月日, a.入社年月日, a.退職年月日, a.メール from dbo.s社員基本情報 a left join dbo.web明細対象者 b on a.社員番号 = b.社員番号 where a.在籍区分 <> '9' and 役職CD<> '0055' and b.社員番号 is null and (現場CD like '%9900' or 組織CD like '%202%' or 現場CD like '%9000')");
            dataGridView2.DataSource = dt2;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string quote = "\"";
            string separator = ",";
            string replace = "";

            string y = comboBox1.SelectedItem.ToString();
            string m = comboBox2.SelectedItem.ToString();

            string pass = @"\\daikensrv03\17_総務部\04_給与\毎月給与計算業務\WEB明細CSV\";
            string file = pass + "Web明細データ_給与" + y + "年" + m + "月" + ".csv";

            DataTable dt = new DataTable();

            dt = Com.GetDB("select * from dbo.web明細_給与データ取得('" + y + "', '" + m + "')");

            Com.OutPutCSV(dt, true, separator, quote, replace, file);

            //作成したフォルダ表示
            System.Diagnostics.Process.Start(pass);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string quote = "\"";
            string separator = ",";
            string replace = "";

            string y = comboBox3.SelectedItem.ToString();
            string m = comboBox4.SelectedItem.ToString().Substring(0,2);

            string pass = @"\\daikensrv03\17_総務部\04_給与\毎月給与計算業務\WEB明細CSV\";
            string file = pass + "Web明細データ_賞与" + y + "年" + comboBox4.SelectedItem.ToString() + ".csv";

            DataTable dt = new DataTable();

            dt = Com.GetDB("select * from dbo.web明細_賞与データ取得('" + y + "', '" + m + "')");

            Com.OutPutCSV(dt, true, separator, quote, replace, file);

            //作成したフォルダ表示
            System.Diagnostics.Process.Start(pass);
        }

        private void button3_Click(object sender, EventArgs e)
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

            GetData();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
