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
    public partial class TData : Form
    {
        public TData()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピーで
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            //フォントサイズの変更
            dataGridView1.Font = new Font(dataGridView1.Font.Name, 12);

            comboBox1.Items.Add("ALL");
            comboBox1.Items.Add("本社");
            comboBox1.Items.Add("那覇");
            comboBox1.Items.Add("八重山");
            comboBox1.Items.Add("北部");
            comboBox1.SelectedIndex = 0;

            comboBox2.Items.Add("ALL");
            comboBox2.Items.Add("2016");
            comboBox2.Items.Add("2017");
            comboBox2.Items.Add("2018");
            comboBox2.SelectedIndex = 0;

            comboBox3.Items.Add("ALL");
            comboBox3.Items.Add("01");
            comboBox3.Items.Add("02");
            comboBox3.Items.Add("03");
            comboBox3.Items.Add("04");
            comboBox3.Items.Add("05");
            comboBox3.Items.Add("06");
            comboBox3.Items.Add("07");
            comboBox3.Items.Add("08");
            comboBox3.Items.Add("09");
            comboBox3.Items.Add("10");
            comboBox3.Items.Add("11");
            comboBox3.Items.Add("12");
            comboBox3.SelectedIndex = 0;

        }

        private void GetData()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable dt = new DataTable();

            string y = comboBox2.SelectedItem.ToString();
            string m = comboBox3.SelectedItem.ToString();
            string p = comboBox1.SelectedItem.ToString();

            if (y == "ALL") y = "";
            if (m == "ALL") m = "";
            if (p == "ALL") p = "";

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandText = "select * from dbo.退職過去データ where 対象年月 like '" + y + "%' and 対象年月 like '%" + m + "' and 地区名 like '%" + p + "%' order by 地区名, 組織名, 現場名";
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
            label1.Text = dt.Rows.Count.ToString() + "人";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetData();
        }
    }
}
