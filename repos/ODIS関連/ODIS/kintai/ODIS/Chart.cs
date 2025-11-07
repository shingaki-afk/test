using C1.Win.C1Chart;
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
    public partial class Chart : Form
    {
        public Chart()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            comboBox1.Items.Add("全地区");
            comboBox1.Items.Add("那覇");
            comboBox1.Items.Add("八重山");
            comboBox1.Items.Add("北部");
            comboBox1.Items.Add("広域");
            comboBox1.Items.Add("宮古島");
            comboBox1.Items.Add("久米島");
            comboBox1.SelectedIndex = 0;

        }



    }
}
