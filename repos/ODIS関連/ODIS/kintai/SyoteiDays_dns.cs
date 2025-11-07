using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ODIS
{
    class SyoteiDays_dns
    {
        /// <summary>
        /// 所定日数
        /// </summary>
        private int syoteiDay;

        /// <summary>
        /// 休日日数
        /// </summary>
        private int kyuuzitsuDay;

        /// <summary>
        /// 法休日数
        /// </summary>
        private int houkyuuDay;

        /// <summary>
        /// 所定労働時間
        /// </summary>
        private int roudouH;

        /// <summary>
        /// 土日祝休日数
        /// </summary>
        private int donichiDay;

        /// <summary>
        /// 月合計日数
        /// </summary>
        private int mtotalDay;

        /// <summary>
        /// パート休日基準日数(週5)
        /// </summary>
        private int kizyun5;

        /// <summary>
        /// パート休日基準日数(週4)
        /// </summary>
        private int kizyun4;

        /// <summary>
        /// パート休日基準日数(週3)
        /// </summary>
        private int kizyun3;

        /// <summary>
        /// パート休日基準日数(週2)
        /// </summary>
        private int kizyun2;

        /// <summary>
        /// パート休日基準日数(週1)
        /// </summary>
        private int kizyun1;

        /// <summary>
        /// 所定日数
        /// </summary>
        public int SyoteiDay
        {
            get { return this.syoteiDay; }
        }

        /// <summary>
        /// 休日日数
        /// </summary>
        public int KyuuzitsuDay
        {
            get { return this.kyuuzitsuDay; }
        }

        /// <summary>
        /// 法定休日日数
        /// </summary>
        public int HoukyuuDay
        {
            get { return this.houkyuuDay; }
        }

        /// <summary>
        /// 所定労働時間
        /// </summary>
        public int RoudouH
        {
            get { return this.roudouH; }
        }

        /// <summary>
        /// 土日祝休日日数
        /// </summary>
        public int DonichiDay
        {
            get { return this.donichiDay; }
        }

        /// <summary>
        /// 月合計日数
        /// </summary>
        public int MtotalDay
        {
            get { return this.mtotalDay; }
        }

        /// <summary>
        /// パート休日基準日数(週5)
        /// </summary>
        public int Kizyun5
        {
            get { return this.kizyun5; }
        }

        /// <summary>
        /// パート休日基準日数(週5)
        /// </summary>
        public int Kizyun4
        {
            get { return this.kizyun4; }
        }

        /// <summary>
        /// パート休日基準日数(週5)
        /// </summary>
        public int Kizyun3
        {
            get { return this.kizyun3; }
        }

        /// <summary>
        /// パート休日基準日数(週5)
        /// </summary>
        public int Kizyun2
        {
            get { return this.kizyun2; }
        }

        /// <summary>
        /// パート休日基準日数(週1)
        /// </summary>
        public int Kizyun1
        {
            get { return this.kizyun1; }
        }



        /// <summary>
        /// 年間所定労働数等を取得
        /// </summary>
        public SyoteiDays_dns()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    //Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "select * from dbo.c所定労働テーブル_土日祝年末年始 where 年月 = (select LEFT(convert(varchar, 対象開始日, 112), 6) from dbo.[001_対象年月マスタ])";
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);

                    //Cmd.CommandText = "select * from dbo.給与立替者";
                    //da = new SqlDataAdapter(Cmd);
                    //da.Fill(dt2);
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                syoteiDay = (int)row["所定"];
                kyuuzitsuDay = (int)row["休日"];
                houkyuuDay = (int)row["法休"];
                roudouH = (int)row["労働時間"];
                mtotalDay = (int)row["所定"] + (int)row["休日"];

                //TODO 
                kizyun5 = (int)row["休日"]; //9
                kizyun4 = (int)row["休日"] + 4; //13
                kizyun3 = (int)row["休日"] + 8;  //18
                kizyun2 = (int)row["休日"] + 12; //22
                kizyun1 = (int)row["休日"] + 16; //26
            }
        }
    }
}
