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
    public partial class ShikakuKigen : Form
    {
        public ShikakuKigen()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            GetData();

            Com.InHistory("24_資格有効期限管理", "", "");
        }

        private void GetData()
        {
            DataTable dt = Com.GetDB("select * from dbo.s資格期限有一覧");
            dataGridView1.DataSource = dt;

        }
    }
}
