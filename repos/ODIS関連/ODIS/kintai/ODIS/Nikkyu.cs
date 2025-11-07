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
    public partial class Nikkyu : Form
    {
        //職種
        private DataTable Syoku = new DataTable();

        //年齢
        private DataTable Nen = new DataTable();

        //社外経験
        DataTable Keiken = new DataTable();

        //学歴
        DataTable Gaku = new DataTable();

        public Nikkyu()
        {
            InitializeComponent();

            GetData();

            Com.InHistory("43_給与新基準_日給額", "", "");
        }

        private void GetData()
        {
            DataTable hkdt = new DataTable();
            hkdt = Com.GetDB("select 本給 from dbo.HK_本給 where '" + Convert.ToDateTime(dateTimePicker2.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

            //honkyuu.Value = 142000;
            honyuu.Text = hkdt.Rows[0][0].ToString();

            //honyuu.Text = "142000";
            syokumukyuu.Text = "0";
            nennreikyuu.Text = "0";
            keikenkyuu.Text = "0";
            gakurekikyuu.Text = "0";


            tuukinhouhou.Items.Add("1 車");
            tuukinhouhou.Items.Add("2 バイク");
            //tuukinhouhou.Items.Add("3 徒歩・自転車");
            tuukinhouhou.Items.Add("4 バス・モノレール");
            tuukinhouhou.Items.Add("5 送迎(会社)");
            tuukinhouhou.Items.Add("6 送迎(知人・親族)");
            tuukinhouhou.Items.Add("7 業務車両");
            tuukinhouhou.Items.Add("8 徒歩");
            tuukinhouhou.Items.Add("9 自転車");
            tuukinhouhou.SelectedIndex = 0;

            kyori.Value = 5;
            kinmu.Value = 21;
            //honyuu.Enabled = false;
            //syokumukyuu.Enabled = false;
            //nennreikyuu.Enabled = false;
            //keikenkyuu.Enabled = false;
            //gakurekikyuu.Enabled = false;
            //goukei.Enabled = false;
            //nikkyuu.Enabled = false;

            Syoku = Com.GetDB("select * from dbo.K_職務給_職種");

            foreach (DataRow row in Syoku.Rows)
            {
                comboBoxSyoku.Items.Add(row["備考"]);
            }

            Nen = Com.GetDB("select * from dbo.K_技能給_A年齢");

            Keiken = Com.GetDB("select * from dbo.K_技能給_B社外経験");
            foreach (DataRow row in Keiken.Rows)
            {
                comboBoxKeiken.Items.Add(row["備考"]);
            }


            Gaku = Com.GetDB("select * from dbo.K_技能給_C最終学歴");

            foreach (DataRow row in Gaku.Rows)
            {
                comboBoxGaku.Items.Add(row["備考"]);
            }

            for (int i = 0; i <= 100; i++)
            {
                comboBoxkasan.Items.Add(i*1000);
            }

            //離島手当 0 or 30000
            comboBoxritou.Items.Add(0);
            comboBoxritou.Items.Add(30000);

            Calc();
        }

        private void Calc()
        {

            //職種
            if (comboBoxSyoku.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Syoku.Select("備考 = '" + comboBoxSyoku.SelectedItem.ToString() +"'");
                syokumukyuu.Text = dr[0][1].ToString();
            }

            if (Nen.Rows.Count == 0) return; 

            //年齢
            int old = CalcAge(dateTimePicker1.Value, dateTimePicker2.Value);
            nenrei.Text = old.ToString();

            DataRow[] dro = Nen.Select("年齢 = '" + nenrei.Text + "'");
            

            nennreikyuu.Text = dro[0][1].ToString();

            //学歴
            if (comboBoxGaku.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Gaku.Select("備考 = '" + comboBoxGaku.SelectedItem.ToString() + "'");
                gakurekikyuu.Text = dr[0][1].ToString();
            }

            //社外経験
            if (comboBoxKeiken.SelectedItem == null)
            {

            }
            else
            {
                DataRow[] dr = Keiken.Select("備考 = '" + comboBoxKeiken.SelectedItem.ToString() + "'");
                keikenkyuu.Text = dr[0][1].ToString();
            }

            //本給
            DataTable hkdt = new DataTable();
            hkdt = Com.GetDB("select 本給 from dbo.HK_本給 where '" + Convert.ToDateTime(dateTimePicker2.Value).ToString("yyyy/MM/dd") + "' between 適用開始日 and 適用終了日");

            //honkyuu.Value = 142000;
            honyuu.Text = hkdt.Rows[0][0].ToString();

            //合計額
            //goukei.Text = (Convert.ToDecimal(honyuu.Text) + Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) +Convert.ToDecimal(gakurekikyuu.Text) + Convert.ToDecimal(kasankyuu.Text)).ToString();
            kihongoukei.Text = (Convert.ToDecimal(honyuu.Text) + Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) + Convert.ToDecimal(gakurekikyuu.Text)).ToString("#,0");
            sikyuugoukei.Text = (Convert.ToDecimal(honyuu.Text) + Convert.ToDecimal(syokumukyuu.Text) + Convert.ToDecimal(nennreikyuu.Text) + Convert.ToDecimal(keikenkyuu.Text) + Convert.ToDecimal(gakurekikyuu.Text) + Convert.ToDecimal(comboBoxritou.SelectedItem) + Convert.ToDecimal(comboBoxkasan.SelectedItem)).ToString("#,0"); 


            //日給換算
            nikkyuu.Text = (Math.Round(Convert.ToDecimal(sikyuugoukei.Text) / Convert.ToDecimal(21.5))).ToString("#,0");


        }

        public int CalcAge(DateTime birthday, DateTime NyusyaDay)
        {
            int i = 0;
            i = (int.Parse(NyusyaDay.ToString("yyyyMMdd")) - int.Parse(birthday.ToString("yyyyMMdd"))) / 10000;
            if (i < 0) i = 0;
            return i;
        }


        //

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxSyoku_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void comboBoxGaku_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void comboBoxKeiken_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void kasankyuu_TextChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void comboBoxkasan_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void Nikkyu_Load(object sender, EventArgs e)
        {

        }

        private void tuukinhouhou_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal tankad = 0;
            decimal hid = 0;
            decimal kad = 0;

            Com.CalcTuukin(tuukinhouhou.SelectedItem.ToString(), kyori.Value, kinmu.Value, ref tankad, ref hid, ref kad);

            tanka.Text = tankad.ToString();
            hi.Text = hid.ToString();
            ka.Text = kad.ToString();
        }

        private void kyori_ValueChanged(object sender, EventArgs e)
        {
            decimal tankad = 0;
            decimal hid = 0;
            decimal kad = 0;

            Com.CalcTuukin(tuukinhouhou.SelectedItem.ToString(), kyori.Value, kinmu.Value, ref tankad, ref hid, ref kad);

            tanka.Text = tankad.ToString();
            hi.Text = hid.ToString();
            ka.Text = kad.ToString();
        }

        private void kinmu_ValueChanged(object sender, EventArgs e)
        {
            decimal tankad = 0;
            decimal hid = 0;
            decimal kad = 0;

            Com.CalcTuukin(tuukinhouhou.SelectedItem.ToString(), kyori.Value, kinmu.Value, ref tankad, ref hid, ref kad);

            tanka.Text = tankad.ToString();
            hi.Text = hid.ToString();
            ka.Text = kad.ToString();
        }

        private void comboBoxritou_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calc();
        }
    }
}
