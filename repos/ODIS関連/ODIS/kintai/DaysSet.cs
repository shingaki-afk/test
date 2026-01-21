using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ODIS
{
    public partial class DaysSet : Form
    {
        public DaysSet()
        {
            InitializeComponent();

            TargetDays td = new TargetDays();
            this.dateTimePicker1.Value = td.StartYMD;
            this.dateTimePicker2.Value = td.EndYMD;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TargetDays td = new TargetDays();
            bool result = td.UpdateTargetDays(this.dateTimePicker1.Value, this.dateTimePicker2.Value);
            if (result)
                MessageBox.Show("登録しました");
            else
                MessageBox.Show("登録失敗しました。電算へ連絡ください");
        }
    }
}
