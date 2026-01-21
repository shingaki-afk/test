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
    public partial class BillKan_list : Form
    {
        public BillKan_list()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.bビル管登録現場一覧");

            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].Width = 300;
            dataGridView1.Columns[1].Width = 100;
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].Width = 300;

            dataGridView1.Columns[3].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            Com.InHistory("633_ビル管登録現場_一覧", "", ""); 
        }
    }
}
