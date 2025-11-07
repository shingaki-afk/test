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
    public partial class Kensyuu : Form
    {
        /// <summary>
        /// 従業員全データ
        /// </summary>
        private DataTable dt = new DataTable();

        public Kensyuu()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            // 選択モードを行単位での選択のみにする
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            hyouka.Items.Add("");
            hyouka.Items.Add("保留中");
            hyouka.Items.Add("有効");
            hyouka.Items.Add("無効");
            hyouka.SelectedIndex = 0;

            comboBox2.Items.Add("");
            comboBox2.Items.Add("社外研修");
            comboBox2.Items.Add("通信研修");
            comboBox2.Items.Add("トップリーグ");
            comboBox2.Items.Add("説明会");
            comboBox2.Items.Add("社内研修");
            comboBox2.Items.Add("プロジェクト");

            GetData();
            Com.InHistory("23_研修登録・更新", "", "");

            //フィルター設定
            c1FlexGrid1.AllowFiltering = true;

            //自動グリップボード機能を有効にする
            c1FlexGrid1.AutoClipboard = true;

            // グリッドのAllowMergingプロパティを設定
            c1FlexGrid1.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free;
            
            // 固定行数を設定
            //c1FlexGrid1.Rows.Fixed = 2;


        }

        private void GetData()
        {
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

            //退職者含めない
            if (!checkBox1.Checked)
            {
                result += " and 退職年月日 is null ";
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
                    if (checkBox2.Checked)
                    {
                        Cmd.CommandText = "select distinct b.* from dbo.研修テーブル a left join dbo.研修対象者検索 b on a.社員番号 = b.社員番号 where 評価 is null"; 
                    }
                    else
                    {
                        Cmd.CommandText = "select 社員番号, 漢字氏名, 組織名, 現場名 from dbo.[研修対象者検索] " + result;
                    }

                    

                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;

            //dataGridView1.Columns["カナ氏名"].Visible = false;
            //dataGridView1.Columns["組織CD"].Visible = false;
            //dataGridView1.Columns["現場CD"].Visible = false;
            //dataGridView1.Columns["reskey"].Visible = false;
        }

        private void GetSyousaiData(string no)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataAdapter da;
            DataTable Syousaidt = new DataTable();

            try
            {
                using (Cn = new SqlConnection(Com.SQLConstr))
                {
                    using (Cmd = Cn.CreateCommand())
                    {
                        if (checkBox2.Checked)
                        {
                            Cmd.CommandText = "select id, 日付, 種別, 講座等名称, 評価, 評価者名, 備考 from dbo.研修テーブル where 評価 is null and 社員番号 = " + no;
                        }
                        else
                        {
                            Cmd.CommandText = "select id, 日付, 種別, 講座等名称, 評価, 評価者名, 備考 from dbo.研修テーブル where 社員番号 = " + no;
                        }
                        
                        da = new SqlDataAdapter(Cmd);
                        da.Fill(Syousaidt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            c1FlexGrid1.DataSource = Syousaidt;
            //c1TrueDBGrid1.DataSource = Syousaidt;
            //dataGridView3.DataSource = Syousaidt;

            //マージ設定
            c1FlexGrid1.Cols[1].AllowMerging = true;
            c1FlexGrid1.Cols[2].AllowMerging = true;
            c1FlexGrid1.Cols[3].AllowMerging = true;
            c1FlexGrid1.Cols[4].AllowMerging = true;
            c1FlexGrid1.Cols[5].AllowMerging = true;
            c1FlexGrid1.Cols[6].AllowMerging = true;
            c1FlexGrid1.Cols[7].AllowMerging = true;

            //選択はしない
            c1FlexGrid1.Select(-1, -1);
            
        }



        private void button1_Click(object sender, EventArgs e)
        {
            dt.Clear();
            GetData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.CurrentRow;
            if (dgr == null) return;
            DataRowView drv = (DataRowView)dgr.DataBoundItem;

            no.Text = drv[0].ToString();
            name.Text = drv[1].ToString();
            soshikin.Text = drv[2].ToString();
            genban.Text = drv[3].ToString();
            GetSyousaiData(drv[0].ToString());
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dt.Clear();
                GetData();
            }
        }

        private void c1FlexGrid1_SelChange(object sender, EventArgs e)
        {

            int row = c1FlexGrid1.RowSel;

            if (c1FlexGrid1.Clip == "") return;

            if (row < 0) return;
            //if (c1FlexGrid1.GetData(row, 1).Equals(DBNull.Value)) return;

            id.Text = c1FlexGrid1.GetData(row, 1).ToString();
            dateTimePicker1.Value = Convert.ToDateTime(c1FlexGrid1.GetData(row, 2));
            comboBox2.SelectedItem = c1FlexGrid1.GetData(row, 3).ToString();
            kouzamei.Text = c1FlexGrid1.GetData(row, 4).ToString();
            hyouka.SelectedItem = c1FlexGrid1.GetData(row, 5).ToString();
            zyoushi.Text = c1FlexGrid1.GetData(row, 6).ToString();
            bikou.Text = c1FlexGrid1.GetData(row, 7).ToString();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("削除してよろしいですか？",
            "警告",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) return;

            SqlConnection Cn;
            SqlCommand Cmd;
            SqlDataReader dr;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandText = "delete from dbo.研修テーブル where id = '" + id.Text + "'";
                    using (dr = Cmd.ExecuteReader())
                    {
                        //TODO
                    }
                }
            }

            GetSyousaiData(no.Text);
        }

        //新規登録
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("講座種別を入力してください。");
                return;
            }

            if (kouzamei.Text == "")
            {
                MessageBox.Show("講座等名称を入力してください。");
                return;
            }
            else
            {
                InsertUpdate("I");
            }
        }

        //更新
        private void button2_Click(object sender, EventArgs e)
        {
            if (kouzamei.Text == "")
            {
                MessageBox.Show("講座等名称を入力してください。");
                return;
            }
            else if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("講座種別を入力してください。");
                return;
            }
            else
            { 
                InsertUpdate("U");
            }
        }

        private void InsertUpdate(string flg)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataReader dr;

            try
            {
                using (Cn = new SqlConnection(Common.constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = "[dbo].[研修データ登録更新]";

                        Cmd.Parameters.Add(new SqlParameter("id", SqlDbType.Int));
                        Cmd.Parameters["id"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("社員番号", SqlDbType.VarChar));
                        Cmd.Parameters["社員番号"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("日付", SqlDbType.Date));
                        Cmd.Parameters["日付"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("種別", SqlDbType.VarChar));
                        Cmd.Parameters["種別"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("講座等名称", SqlDbType.VarChar));
                        Cmd.Parameters["講座等名称"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("備考", SqlDbType.VarChar));
                        Cmd.Parameters["備考"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("評価", SqlDbType.VarChar));
                        Cmd.Parameters["評価"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("評価者名", SqlDbType.VarChar));
                        Cmd.Parameters["評価者名"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("更新者", SqlDbType.VarChar));
                        Cmd.Parameters["更新者"].Direction = ParameterDirection.Input;

                        Cmd.Parameters.Add(new SqlParameter("更新時間", SqlDbType.DateTime));
                        Cmd.Parameters["更新時間"].Direction = ParameterDirection.Input;

                        if (flg == "I")
                        {
                            Cmd.Parameters["id"].Value = 0;
                        }
                        else
                        {
                            Cmd.Parameters["id"].Value = Convert.ToInt32(id.Text);
                        }

                        Cmd.Parameters["社員番号"].Value = no.Text;
                        Cmd.Parameters["日付"].Value = dateTimePicker1.Value;
                        Cmd.Parameters["種別"].Value = comboBox2.SelectedItem.ToString();
                        Cmd.Parameters["講座等名称"].Value = kouzamei.Text;
                        Cmd.Parameters["備考"].Value = bikou.Text.ToString();
                        Cmd.Parameters["評価"].Value = hyouka.SelectedItem.ToString();
                        Cmd.Parameters["評価者名"].Value = zyoushi.Text.ToString();
                        Cmd.Parameters["更新者"].Value = Program.loginname;
                        Cmd.Parameters["更新時間"].Value = DateTime.Now;

                        using (dr = Cmd.ExecuteReader())
                        {
                            MessageBox.Show("登録・更新しました");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }

            GetSyousaiData(no.Text);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }
    }
}
