using Microsoft.VisualBasic;
using Npgsql;
using System;
using System.Data;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class GyoumuKanri : Form
    {
        private string result;
        private DataTable dt = new DataTable();
        private string ym;

        public GyoumuKanri()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            //dataGridView3.Font = new Font(dataGridView3.Font.Name, 10);

            // 選択モードを行単位での選択のみにする
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            comboBox1.Items.Add("202304");
            comboBox1.Items.Add("202305");
            comboBox1.Items.Add("202306");
            comboBox1.Items.Add("202307");
            comboBox1.Items.Add("202308");
            comboBox1.Items.Add("202309");
            comboBox1.Items.Add("202310");
            comboBox1.Items.Add("202311");
            comboBox1.Items.Add("202312");
            comboBox1.Items.Add("202401");
            comboBox1.Items.Add("202402");
            comboBox1.Items.Add("202403");
            comboBox1.Items.Add("202404");
            comboBox1.Items.Add("202405");
            comboBox1.Items.Add("202406");
            comboBox1.Items.Add("202407");
            comboBox1.Items.Add("202408");
            comboBox1.Items.Add("202409");
            comboBox1.Items.Add("202410");
            comboBox1.Items.Add("202411");
            comboBox1.Items.Add("202412");
            comboBox1.Items.Add("202501");
            comboBox1.Items.Add("202502");
            comboBox1.Items.Add("202503");

            comboBox1.SelectedIndex = comboBox1.FindString(DateTime.Now.AddDays(-5).ToString("yyyyMM"));

            //comboBox1.SelectedIndex = 5;
            ym = comboBox1.SelectedItem.ToString();

            checkedListBox1.Items.Add("契約固定");
            checkedListBox1.Items.Add("契約臨時");
            checkedListBox1.Items.Add("臨時");
            checkedListBox1.Items.Add("物品");
            checkedListBox1.SetItemChecked(0, true);
            checkedListBox1.SetItemChecked(1, true);
            checkedListBox1.SetItemChecked(2, true);
            checkedListBox1.SetItemChecked(3, true);

            checkedListBox2.Items.Add("自社");
            checkedListBox2.Items.Add("外注");
            checkedListBox2.SetItemChecked(0, true);
            checkedListBox2.SetItemChecked(1, true);

            cbbikou.Checked = true;
            cbkakunin.Checked = true;
            cbkeihi.Checked = true;
            cbshiharai.Checked = true;
            cbsyori.Checked = true;
            cbsyouhi.Checked = true;
            cbtani.Checked = true;

            //checkedListBox3.Items.Add("消費税関係");
            //checkedListBox3.Items.Add("単価関係");
            //checkedListBox3.SetItemChecked(0, true);
            //checkedListBox3.SetItemChecked(1, true);

            //checkedListBox4.Items.Add("1_売上");
            //checkedListBox4.Items.Add("2_実施");
            //checkedListBox4.SetItemChecked(0, true);
            //checkedListBox4.SetItemChecked(1, true);

            SetBumon();
            GetUriageData();

            Com.InHistory("14_業務管理台帳", "", "");
        }

        private void SetBumon()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct 部門コード, 部門 from kpcp01.\"CostomGyoumuKanri\" order by 部門コード";

            dt = Com.GetPosDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["部門コード"].ToString() + ' ' + row["部門"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        private void GetDisp()
        {
            //検索文字列処理
            ResultStr();

            DataRow[] dtrow;
            dtrow = dt.Select(result, "");


            DataTable Disp = new DataTable();
            //1
            Disp.Columns.Add("売上年月", typeof(string));
            Disp.Columns.Add("連番", typeof(string));
            //2
            Disp.Columns.Add("契約区分", typeof(string));
            Disp.Columns.Add("作業区分", typeof(string));
            //3
            Disp.Columns.Add("部門コード", typeof(string));
            Disp.Columns.Add("部門", typeof(string));
            //4
            Disp.Columns.Add("オーナーコード", typeof(string));
            Disp.Columns.Add("オーナー名", typeof(string));
            //5
            Disp.Columns.Add("工事コード", typeof(string));
            Disp.Columns.Add("工事枝", typeof(string));
            Disp.Columns.Add("工事", typeof(string));
            Disp.Columns.Add("工事種別", typeof(string));
            Disp.Columns.Add("契約項目", typeof(string));
            //6
            Disp.Columns.Add("担当者氏名", typeof(string));
            Disp.Columns.Add("売上額", typeof(decimal));
            Disp.Columns.Add("実施額", typeof(decimal));
            Disp.Columns.Add("人件費", typeof(decimal));

            //7
            Disp.Columns.Add("売上チェック", typeof(string));
            Disp.Columns.Add("売上入力日", typeof(string));
            Disp.Columns.Add("売上対象区分", typeof(string));
            Disp.Columns.Add("作業内容区分", typeof(string));
            Disp.Columns.Add("作業対象区分", typeof(string));

            //支払先
            Disp.Columns.Add("業者コード", typeof(string));
            Disp.Columns.Add("支払業者", typeof(string));
            Disp.Columns.Add("支払金額", typeof(decimal));

            //単価
            
            Disp.Columns.Add("数量", typeof(decimal));
            Disp.Columns.Add("単位", typeof(string));
            Disp.Columns.Add("単価", typeof(decimal));

            //消費税
            Disp.Columns.Add("消費税額", typeof(decimal));
            Disp.Columns.Add("消費税区分", typeof(string));
            Disp.Columns.Add("課税区分", typeof(string));

            //確認書
            Disp.Columns.Add("確認書", typeof(string));
            Disp.Columns.Add("印刷順", typeof(string));
            Disp.Columns.Add("頁No", typeof(string));
            Disp.Columns.Add("まとめNo", typeof(string));

            //処理情報
            Disp.Columns.Add("更新者", typeof(string));
            Disp.Columns.Add("更新日時", typeof(string));
            Disp.Columns.Add("部門区分", typeof(string));

            //経費
            Disp.Columns.Add("細目名", typeof(string));
            Disp.Columns.Add("材料・物品", typeof(decimal));
            Disp.Columns.Add("外注費", typeof(decimal));
            Disp.Columns.Add("その他費目", typeof(decimal));
            Disp.Columns.Add("合計額", typeof(decimal));

            //備考
            Disp.Columns.Add("支払回数", typeof(decimal));
            Disp.Columns.Add("実施回数", typeof(decimal));
            Disp.Columns.Add("備考", typeof(string));

            int uriagegaku = 0;
            int zissigaku = 0;
            int zinkenhi = 0;
            int shiharai = 0;
            int syouhizei = 0;

            foreach (DataRow row in dtrow)
            {
                DataRow nr = Disp.NewRow();

                //
                nr["売上年月"] = row["売上年月"];
                nr["連番"] = row["連番"];

                nr["契約区分"] = row["契約区分"];
                nr["作業区分"] = row["作業区分"];

                nr["部門コード"] = row["部門コード"];
                nr["部門"] = row["部門"];

                nr["オーナーコード"] = row["オーナーコード"];
                nr["オーナー名"] = row["オーナー名"];

                nr["工事コード"] = row["工事コード"];
                nr["工事枝"] = row["工事枝"];
                nr["工事"] = row["工事"];
                nr["工事種別"] = row["工事種別"];
                nr["契約項目"] = row["契約項目"];

                nr["担当者氏名"] = row["担当者氏名"];
                nr["売上額"] = row["売上額"]; uriagegaku += Convert.ToInt32(row["売上額"]);
                nr["実施額"] = row["実施額"]; zissigaku += Convert.ToInt32(row["実施額"]);
                nr["人件費"] = row["人件費"]; zinkenhi += Convert.ToInt32(row["人件費"]);


                //よくわからん区分
                nr["売上入力日"] = row["売上入力日"];
                nr["売上チェック"] = row["売上チェック"];
                nr["作業内容区分"] = row["作業内容区分"];
                nr["売上対象区分"] = row["売上対象区分"];
                nr["作業対象区分"] = row["作業対象区分"];

                //支払先3
                nr["業者コード"] = row["業者コード"];
                nr["支払業者"] = row["支払業者"]; 
                nr["支払金額"] = row["支払金額"]; shiharai += Convert.ToInt32(row["支払金額"]);

                //単価3
                nr["数量"] = row["数量"];
                nr["単位"] = row["単位"];
                nr["単価"] = row["単価"];

                //消費税区分3
                nr["消費税額"] = row["消費税額"]; syouhizei += Convert.ToInt32(row["消費税額"]);
                nr["消費税区分"] = row["消費税区分"];
                nr["課税区分"] = row["課税区分"];

                //確認書4
                nr["確認書"] = row["確認書"];
                nr["印刷順"] = row["印刷順"];
                nr["頁No"] = row["頁No"];
                nr["まとめNo"] = row["まとめNo"];

                //処理情報3
                nr["更新者"] = row["更新者"];
                nr["更新日時"] = row["更新日時"];
                nr["部門区分"] = row["部門区分"];

                //経費
                nr["細目名"] = row["細目名"];
                nr["材料・物品"] = row["材料・物品"];
                nr["外注費"] = row["外注費"];
                nr["その他費目"] = row["その他費目"];
                nr["合計額"] = row["合計額"];

                //備考
                nr["支払回数"] = row["支払回数"];
                nr["実施回数"] = row["実施回数"];
                nr["備考"] = row["備考"];

                Disp.Rows.Add(nr);
            }




            //foreach (DataRow dr in dtrow)
            //{
            //    uriagegaku += Convert.ToInt32(dr["売上額"]);
            //    zissigaku += Convert.ToInt32(dr["実施額"]);
            //    zinkenhi += Convert.ToInt32(dr["人件費"]);
            //    shiharai += Convert.ToInt32(dr["支払金額"]);
            //    syouhizei += Convert.ToInt32(dr["消費税額"]);
            //}

            //合計
            DataRow nrtotal = Disp.NewRow();
            nrtotal["売上年月"] = "合計";
            nrtotal["連番"] = "";
            nrtotal["契約区分"] = "";
            nrtotal["作業区分"] = "";
            nrtotal["部門コード"] = "";
            nrtotal["部門"] = "";
            nrtotal["オーナーコード"] = "";
            nrtotal["オーナー名"] = "";
            nrtotal["工事コード"] = "";
            nrtotal["工事枝"] = "";
            nrtotal["工事"] = "";
            nrtotal["契約項目"] = "";
            nrtotal["担当者氏名"] = "";
            nrtotal["売上額"] = uriagegaku;
            nrtotal["実施額"] = zissigaku;
            nrtotal["人件費"] = zinkenhi;
            nrtotal["業者コード"] = "";
            nrtotal["支払業者"] = "";
            nrtotal["支払金額"] = shiharai;
            nrtotal["工事種別"] = "";
            nrtotal["売上入力日"] = "";
            nrtotal["売上チェック"] = "";
            nrtotal["作業内容区分"] = "";
            nrtotal["売上対象区分"] = "";
            nrtotal["作業対象区分"] = "";
            nrtotal["数量"] = DBNull.Value;
            nrtotal["単位"] = "";
            nrtotal["単価"] = DBNull.Value;
            nrtotal["消費税額"] = syouhizei;
            nrtotal["消費税区分"] = "";
            nrtotal["課税区分"] = "";
            nrtotal["頁No"] = "";
            nrtotal["印刷順"] = "";
            nrtotal["まとめNo"] = "";
            nrtotal["確認書"] = "";
            nrtotal["更新者"] = "";
            nrtotal["更新日時"] = "";
            //nrtotal["細目コード"] = "";
            nrtotal["部門区分"] = "";
            nrtotal["細目名"] = "";
            nrtotal["材料・物品"] = DBNull.Value;
            nrtotal["外注費"] = DBNull.Value;
            nrtotal["その他費目"] = DBNull.Value;
            nrtotal["合計額"] = DBNull.Value;
            nrtotal["支払回数"] = DBNull.Value;
            nrtotal["実施回数"] = DBNull.Value;
            nrtotal["備考"] = "";

            Disp.Rows.Add(nrtotal);


            dataGridView3.DataSource = Disp;

            int ct = Disp.Columns.Count;

            dataGridView3.Columns["売上年月"].Width = 45;
            dataGridView3.Columns["連番"].Width = 40;

            dataGridView3.Columns["契約区分"].Width = 60;
            dataGridView3.Columns["作業区分"].Width = 40;

            dataGridView3.Columns["部門コード"].Width = 40;
            dataGridView3.Columns["部門"].Width = 80;

            dataGridView3.Columns["オーナーコード"].Width = 65;
            dataGridView3.Columns["オーナー名"].Width = 170;

            dataGridView3.Columns["工事コード"].Width = 40;
            dataGridView3.Columns["工事枝"].Width = 25;
            dataGridView3.Columns["工事"].Width = 170;
            dataGridView3.Columns["工事種別"].Width = 70;
            dataGridView3.Columns["契約項目"].Width = 170;

            dataGridView3.Columns["担当者氏名"].Width = 80;
            dataGridView3.Columns["売上額"].Width = 60;
            dataGridView3.Columns["実施額"].Width = 60;
            dataGridView3.Columns["人件費"].Width = 60;

            dataGridView3.Columns["売上チェック"].Width = 30;
            dataGridView3.Columns["売上入力日"].Width = 60;
            dataGridView3.Columns["売上対象区分"].Width = 30;
            dataGridView3.Columns["作業内容区分"].Width = 120;
            dataGridView3.Columns["作業対象区分"].Width = 30;

            //支払先
            dataGridView3.Columns["業者コード"].Width = 65;
            dataGridView3.Columns["支払業者"].Width = 170;
            dataGridView3.Columns["支払金額"].Width = 60;

            //単位3
            dataGridView3.Columns["数量"].Width = 40;
            dataGridView3.Columns["単位"].Width = 30;
            dataGridView3.Columns["単価"].Width = 40;


            //消費税区分3
            dataGridView3.Columns["消費税額"].Width = 60;
            dataGridView3.Columns["消費税区分"].Width = 50;
            dataGridView3.Columns["課税区分"].Width = 30;

            //確認書4
            dataGridView3.Columns["確認書"].Width = 30;
            dataGridView3.Columns["印刷順"].Width = 30;
            dataGridView3.Columns["頁No"].Width = 30;
            dataGridView3.Columns["まとめNo"].Width = 30;

            //処理情報3
            dataGridView3.Columns["更新者"].Width = 80;
            dataGridView3.Columns["更新日時"].Width = 70;
            dataGridView3.Columns["部門区分"].Width = 70;

            //経費
            dataGridView3.Columns["細目名"].Width = 90;
            dataGridView3.Columns["材料・物品"].Width = 60;
            dataGridView3.Columns["外注費"].Width = 60;
            dataGridView3.Columns["その他費目"].Width = 60;
            dataGridView3.Columns["合計額"].Width = 60;

            //備考
            dataGridView3.Columns["支払回数"].Width = 50;
            dataGridView3.Columns["実施回数"].Width = 50;
            dataGridView3.Columns["備考"].Width = 200;


            //表示非表示
            //支払先
            dataGridView3.Columns["業者コード"].Visible = cbshiharai.Checked;
            dataGridView3.Columns["支払業者"].Visible = cbshiharai.Checked;
            dataGridView3.Columns["支払金額"].Visible = cbshiharai.Checked;

            //単位3
            dataGridView3.Columns["数量"].Visible = cbtani.Checked;
            dataGridView3.Columns["単位"].Visible = cbtani.Checked;
            dataGridView3.Columns["単価"].Visible = cbtani.Checked;


            //消費税区分3
            dataGridView3.Columns["消費税額"].Visible = cbsyouhi.Checked;
            dataGridView3.Columns["消費税区分"].Visible = cbsyouhi.Checked;
            dataGridView3.Columns["課税区分"].Visible = cbsyouhi.Checked;

            //確認書4
            dataGridView3.Columns["確認書"].Visible = cbkakunin.Checked;
            dataGridView3.Columns["印刷順"].Visible = cbkakunin.Checked;
            dataGridView3.Columns["頁No"].Visible = cbkakunin.Checked;
            dataGridView3.Columns["まとめNo"].Visible = cbkakunin.Checked;

            //処理情報3
            dataGridView3.Columns["更新者"].Visible = cbsyori.Checked;
            dataGridView3.Columns["更新日時"].Visible = cbsyori.Checked;
            dataGridView3.Columns["部門区分"].Visible = cbsyori.Checked;

            //経費
            dataGridView3.Columns["細目名"].Visible = cbkeihi.Checked;
            dataGridView3.Columns["材料・物品"].Visible = cbkeihi.Checked;
            dataGridView3.Columns["外注費"].Visible = cbkeihi.Checked;
            dataGridView3.Columns["その他費目"].Visible = cbkeihi.Checked;
            dataGridView3.Columns["合計額"].Visible = cbkeihi.Checked;

            //備考
            dataGridView3.Columns["支払回数"].Visible = cbbikou.Checked;
            dataGridView3.Columns["実施回数"].Visible = cbbikou.Checked;
            dataGridView3.Columns["備考"].Visible = cbbikou.Checked;


            //for (int i = 15; i < ct - 1; i++)
            //{
            //    dataGridView3.Columns[i].Width = 100;
            //}

            //ヘッダーの中央表示
            for (int i = 0; i < ct; i++)
            {
                dataGridView3.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            int ii = 0;
            // 表示位置
            foreach (DataColumn col in Disp.Columns)
            {
                
                if (col.DataType.Name == "String")
                {
                    //string
                    dataGridView3.Columns[ii].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                }
                else
                {
                    //decimal
                    dataGridView3.Columns[ii].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[ii].DefaultCellStyle.Format = "#,0";
                }

                ii++;
            }
            //dataGridView3.Columns[ct - 1].Wi  dth = 350;
            //dataGridView3.Columns[ct - 1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

        }

        private void ResultStr()
        {
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            //TODO
            result = "";

            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }


            //契約区分
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        result += " and 契約区分 <> '" + checkedListBox1.Items[i].ToString() + "'";
                    }
                }
            }

            //臨時・外注
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 作業区分 <> '" + checkedListBox2.Items[i].ToString() + "'";
                }
            }

            //部門コード
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                if (!checkedListBox3.GetItemChecked(i))
                {
                    result += " and 部門コード <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
                }
            }

            ////額名
            //for (int i = 0; i < checkedListBox4.Items.Count; i++)
            //{
            //    if (!checkedListBox4.GetItemChecked(i))
            //    {
            //        result += " and 額名 <> '" + checkedListBox3.Items[i].ToString().Substring(0, 5) + "'";
            //    }
            //}

            //先頭が「and」の場合、削除する
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
            }


        }

        private void GetUriageData()
        {
            dataGridView3.DataSource = null;
            
            string sql = "";
            sql = "select * from kpcp01.\"CostomGyoumuKanri_serch\" where 売上年月 = '" + comboBox1.SelectedItem.ToString() + "' ";
            //TODO
            //sql += " order by 工事コード, 連番, 部門コード, 額名 ";

            dt = Com.GetPosDB(sql);

            

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //コンボボックス無効化・カーソル変更
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            if (ym != comboBox1.SelectedItem.ToString())
            {
                GetUriageData();
                ym = comboBox1.SelectedItem.ToString();
            }
            GetDisp();

            //カーソル変更・メッセージキュー処理・コンボボックス有効化
            Cursor.Current = Cursors.Default;
            Application.DoEvents();
            button1.Enabled = true;
        }

        private void label3_Click(object sender, EventArgs e)
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

        private void label1_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        //private void label4_Click(object sender, EventArgs e)
        //{
        //    if (checkedListBox4.CheckedItems.Count == 0)
        //    {
        //        for (int i = 0; i < checkedListBox4.Items.Count; i++)
        //        {
        //            checkedListBox4.SetItemChecked(i, true);
        //        }
        //    }
        //    else
        //    {
        //        for (int i = 0; i < checkedListBox4.Items.Count; i++)
        //        {
        //            checkedListBox4.SetItemChecked(i, false);
        //        }
        //    }
        //}

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //カーソル変更
                Cursor.Current = Cursors.WaitCursor;

                GetDisp();

                //カーソル変更・メッセージキュー処理
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void GyoumuKanri_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {
            if (checkedListBox3.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, false);
                }
            }
        }
    }
}
