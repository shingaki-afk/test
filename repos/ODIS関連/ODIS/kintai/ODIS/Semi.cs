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
    public partial class Semi : Form
    {
        public Semi()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //フォントサイズの変更
            c1FlexGrid1.Font = new Font(c1FlexGrid1.Font.Name, 12);

            GetData();
        }

        private void GetData()
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
                        Cmd.CommandText = "select b.氏名, b.地区名, b.組織名, b.現場名, b.給与支給区分名, b.役職名, a.[id],a.[社員番号],a.[日付],a.[種別],a.[講座等名称],a.[備考],a.[評価],a.[評価者名],a.[更新者],a.[更新時間] from dbo.研修テーブル a left join dbo.社員基本情報 b on a.社員番号 = b.社員番号";
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

            // グリッドのAllowMergingプロパティを設定
            //c1FlexGrid1.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free;
            // 固定行数を設定

            //フィルター設定
            //c1FlexGrid1.AllowFiltering = true;

            //自動グリップボード機能を有効にする
            //c1FlexGrid1.AutoClipboard = true;

            //c1FlexGrid1.Rows.Fixed = 2;
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
        }
    }
}
