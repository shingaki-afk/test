using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS
{
    public partial class OCRList : Form
    {
        private string selectionNum = "";

        public OCRList()
        {
            InitializeComponent();
        }

        public OCRList(DataTable dt, string num)
        {
            InitializeComponent();

            dataGridView1.DataSource = dt;
            selectionNum = num;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // 2列目のセル文字列に「Encoder」が含まれていれば、
            // 文字を赤色にする（セルの選択時にも）
            if (e.ColumnIndex == 1)
            {
                string text = e.Value.ToString();
                if (text.Contains(selectionNum))
                {
                    e.CellStyle.ForeColor = Color.Red;
                    e.CellStyle.SelectionForeColor = Color.Red;
                }
            }
        }

    }
}
