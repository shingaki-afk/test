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
    public partial class Nenchou : Form
    {
        public Nenchou()
        {
            InitializeComponent();
            GetData();
        }

        private void GetData()
        {
            //DataTable dt = Com.GetDB("select * from n年調データ表示('2019','2020','10207','10218')");n年調データ表示_単年
            DataTable dt = Com.GetDB("select * from n年調データ表示_単年('2020','10207','10218') order by 現場CD, 組織CD, カナ名, 年度 desc"); 
            dataGridView1.DataSource = dt;
        }
    }
}
