using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using ODIS.ODIS;

namespace ODIS
{
    class TargetDays
    {
        /// <summary>
        /// 期間開始日
        /// </summary>
        private DateTime startYMD;

        /// <summary>
        /// 期間終了日
        /// </summary>
        private DateTime endYMD;

        /// <summary>
        /// 期間合計日
        /// </summary>
        private int tDays;

        /// <summary>
        /// 有効退職日
        /// </summary>
        private DateTime retireYMD;

        /// <summary>
        /// 期間開始日
        /// </summary>
        public DateTime StartYMD
        {
            get { return this.startYMD; }
        }

        /// <summary>
        /// 期間終了日
        /// </summary>
        public DateTime EndYMD
        {
            get { return this.endYMD; }
        }

        /// <summary>
        /// 期間合計日
        /// </summary>
        public int TDays
        {
            get { return this.tDays; }
        }

        /// <summary>
        /// 有効退職日
        /// </summary>
        public DateTime RetireYMD
        {
            get { return this.retireYMD; }
        }

        /// <summary>
        /// 期間登録年月を取得
        /// </summary>
        public TargetDays()
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[GetDays]";
                    da = new SqlDataAdapter(Cmd);
                    da.Fill(dt);
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                startYMD = (DateTime)row["対象開始日"];
                endYMD = (DateTime)row["対象終了日"];
            }

            //開始日と有効退職日は同じ
            retireYMD = startYMD;


            TimeSpan ts = this.endYMD - this.startYMD;
            //対象期間の日数
            tDays = ts.Days + 1;
        }

        /// <summary>
        /// 期間有効年月を更新
        /// </summary>
        /// <param name="TargetStart">期間開始日</param>
        /// <param name="TargetEnd">期間終了日</param>
        /// <returns>更新結果フラグ 0:成功　1:失敗</returns>
        public bool UpdateTargetDays(DateTime TargetStart, DateTime TargetEnd)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            SqlDataReader dr;


            //勤怠入力状況管理をリセット
            dt2 = Com.GetDB("update dbo.勤怠状況管理 set 処理年月 = '" + TargetStart.AddMonths(1).ToString("yyyyMM") + "', 入力完了日時 = null, チェック完了日時 = null");


            //対象年月の変更
            using (Cn = new SqlConnection(Common.constr))
            {
                Cn.Open();

                using (Cmd = Cn.CreateCommand())
                {
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.CommandText = "[dbo].[UpdateDays]";

                    Cmd.Parameters.Add(new SqlParameter("対象開始日", SqlDbType.Date));
                    Cmd.Parameters["対象開始日"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("対象終了日", SqlDbType.Date));
                    Cmd.Parameters["対象終了日"].Direction = ParameterDirection.Input;

                    Cmd.Parameters.Add(new SqlParameter("ReturnValue", SqlDbType.VarChar));
                    Cmd.Parameters["ReturnValue"].Direction = ParameterDirection.ReturnValue;

                    Cmd.Parameters["対象開始日"].Value = TargetStart;
                    Cmd.Parameters["対象終了日"].Value = TargetEnd;

                    using (dr = Cmd.ExecuteReader())
                    {
                        int returnValue = ((int)Cmd.Parameters["ReturnValue"].Value);
                        if (returnValue == 0)
                            return true;
                        else
                            return false;
                    }
                }
            }


        }
    }
}
