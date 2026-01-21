using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace ODIS
{
    public class Common
    {
        //TODO エラーハンドリング対応
        //TODO ストアド化対応
        //TODO オブジェクト指向リファクタ
        //TODO ヌルぽ対応
        //TODO 勤怠テーブルを過去データとしてコピー&トランケートする処理          

        //TODO サブデータとメインデータの統合を検討する
        //サブ情報データテーブル
        public DataTable nenkyuudt;

        //DB接続文字列
        public static string constr = "Password=Pa$$w0rd;User ID=developer;Initial Catalog=dev;Data Source=192.168.100.10";


        //TODO 20250403コメントアウト
        /// <summary>
        /// 所定労働等インスタンス
        /// </summary>
        //private SyoteiDays sd = new SyoteiDays();

        public Common()
        {
            nenkyuudt = GetKintaiKihon(4, "");
        }

        /// <summary>
        /// 給与区分のコード/表示名のハッシュテーブル
        /// </summary>
        public Dictionary<string, string> GetKubunName = new Dictionary<string, string>()
        {
            //TODO いる？
            {"A1", "役員"},
            {"B1", "兼務役員"},
            {"C1", "月給者"},
            //{"C2", "功労月給者"}, 
            {"D1", "日給者"}, 
            //{"D2", "功労日給者"}, 
            {"E1", "パート"}, 
            {"F1", "アルバイト"}, 
            {"", ""}, 
        };

        /// <summary>
        /// 休暇付与区分コード/表示名のハッシュテーブル
        /// </summary>
        public Dictionary<string, string> GetKyuukaKubun = new Dictionary<string, string>()
        {
            {"0", "5日以上"},
            {"1", "4日"},
            {"2", "3日"},
            {"3", "2日"}, 
            {"4", "1日"}, 
            {"9", "ｱﾙﾊﾞｲﾄ"}, 
        };

        //nullの場合は0.0を代入。
        public decimal ConvDec(object dr)
        {
            return dr.Equals(DBNull.Value)?Convert.ToDecimal("0.0"):Convert.ToDecimal(dr);
        }

        //nullの場合は0を代入
        public Int32 ConvInt(object dr)
        {
            return dr.Equals(DBNull.Value) ? Convert.ToInt32("0") : Convert.ToInt32(dr);
        }

        /// <summary>
        /// エラーチェック
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public string[] ErrorCheck(DataRow row, string flg)
        {
            string syainno = row[0].ToString();
            decimal entyouH = ConvDec(row[1]);
            decimal houteiH = ConvDec(row[2]);
            decimal syoteiH = ConvDec(row[3]);
            decimal souzanH = ConvDec(row[4]);
            decimal tyouzanH = ConvDec(row[5]);
            decimal shinyaH = ConvDec(row[6]);
            int tikoku = ConvInt(row[7]);
            decimal tikokuH = ConvDec(row[8]);
            decimal syotei = ConvDec(row[9]);
            decimal houteikyuu = ConvDec(row[10]);
            decimal syoteikyuu = ConvDec(row[11]);
            decimal yuukyuu = ConvDec(row[12]);
            int tokkyuu = ConvInt(row[13]);
            int mutoku = ConvInt(row[14]);
            int hurikyuu = ConvInt(row[15]);
            decimal koukyuu = ConvDec(row[16]);
            //int tyoukyuu = ConvInt(row[17]);
            decimal todokede = ConvDec(row[17]);
            decimal mutodoke = ConvDec(row[18]);
            int kaisuu1 = ConvInt(row[19]);
            int kaisuu2 = ConvInt(row[20]);
            int tanka1 = ConvInt(row[32]);
            int tanka2 = ConvInt(row[33]);

            int kinmuh = ConvInt(row[49]);
            int syaho = ConvInt(row[51]);
            int jyuumin = ConvInt(row[52]);
            int sonota = ConvInt(row[55]);
            string bikou = row[56].ToString();
            string taisyoku = "";

            //59
            int syoteiday = ConvInt(row[59]); //所定日数
            int kyuuzitsuday = ConvInt(row[60]); //休日日数
            int houkyuuday = ConvInt(row[61]); //法定休日日数
            int roudouh = ConvInt(row[62]); //月労働時間

            //給与計算後チェックでは実施しない
            if (flg != "After")
            {
                //退職マスタ対応
                if (row[31].Equals(DBNull.Value))
                {
                    if (!row["退職日"].Equals(DBNull.Value))
                    {
                        taisyoku = row["退職日"].ToString();
                    }
                }
                else
                {
                    taisyoku = row[31].ToString();
                }
            }
            else
            {
                //TODO これでいいのか不明201606
                taisyoku = row[31].ToString();
            }

            //役職　係長対応で使用
            decimal yakusyoku = ConvDec(row["役職CD"]);

            string kyuuyo = row[29].ToString();

            string roudouD = row["労働日数"].ToString();

            //改行
            string nl = Environment.NewLine;

            #region エラーチェック

            #region ① 日数チェック(今月入社と今月退職の対応)

            TargetDays td = new TargetDays();
            int totalDay = 0;
            TimeSpan ts;
            string lbl = "";

            //入社日が開始日より後　退職日が開始日より後の場合
            //[30]は入社日  [31]は退職日
            if (Convert.ToDateTime(row[30]) >= td.StartYMD)
            {
                if (taisyoku != "")
                {
                    //期間中に入社＆退職
                    ts = Convert.ToDateTime(taisyoku) - Convert.ToDateTime(row[30]);
                    totalDay = ts.Days + 1;
                    lbl = "3";
                }
                else
                {
                    //期間中に入社
                    ts = td.EndYMD - Convert.ToDateTime(row[30]);
                    totalDay = ts.Days + 1;
                    lbl = "1";
                }
            }
            else
            {
                if (taisyoku != "")
                {
                    if (td.EndYMD < Convert.ToDateTime(taisyoku))
                    {
                        //通常
                        totalDay = td.TDays;
                    }
                    else
                    {
                        //期間中に退職
                        ts = Convert.ToDateTime(taisyoku) - td.StartYMD;
                        totalDay = ts.Days + 1;
                        lbl = "2";
                    }

                }
                else
                {
                    //通常
                    totalDay = td.TDays;
                }
            }
            #endregion

            #region ② 日数チェック(入力した日数)
            decimal sumDays_sub1;
            decimal sumDays_sub2;
            decimal sumDays;

            //出勤なしで～に利用
            sumDays_sub1 = syotei + houteikyuu + syoteikyuu;

            //欠格月チェックに利用していた(～2018/09)
            sumDays_sub2 = sumDays_sub1 + yuukyuu + tokkyuu; //+ tyoukyuu;

            //合計日数
            sumDays = sumDays_sub2 + mutoku + hurikyuu + koukyuu + todokede + mutodoke;

            //エラーメッセージ
            string errorMSG = "";

            //警告メッセージ
            string WarningMSG = "";

            #endregion

            #region エラーチェック
            //日数オーバー
            if (syotei > totalDay) errorMSG += "所定オーバーです" + nl;
            if (houteikyuu > totalDay) errorMSG += "法定休出オーバーです" + nl;
            if (syoteikyuu > 10 & syotei == 0) errorMSG += "所定休出オーバーです" + nl;
            if (yuukyuu > totalDay) errorMSG += "有給オーバーです" + nl;

            //対象者のサブデータを取得
            DataRow[] nenkyuudr = nenkyuudt.Select("社員番号 = " + row[0], "");
            foreach (DataRow r in nenkyuudr)
            {
                if (r[1].ToString() != "")
                {
                    if (yuukyuu > Convert.ToDecimal(r[1]))
                    {
                        errorMSG += "有給オーバーです" + nl;
                    }
                    else if (lbl == "2" && yuukyuu != Convert.ToDecimal(r[1]))
                    {
                        //退職月で有給残あり
                        WarningMSG += "退職月ですが有給が残っています。本人了承済みですか？" + nl;
                    }
                }
                else
                {
                    if (yuukyuu > 0)
                    {
                        errorMSG += "有給ないです" + nl;
                    }
                }

                //当月手遅れ
                //10 状況
                //11 必須残
                if (sumDays != 0 && r[10].ToString().Length > 4)
                {
                    if (r[10].ToString().Substring(0, 4) == "【当月中" && Convert.ToDecimal(r[11].ToString()) > yuukyuu)
                    {
                        //年休5日オーバー
                        WarningMSG += "年5日有給取得するには、当月であと" + (Convert.ToDecimal(r[11].ToString()) - yuukyuu).ToString() + "日有給取得する必要がありました。" + nl;
                    }
                }
            }

            if (tokkyuu > totalDay) errorMSG += "特休オーバーです" + nl;
            if (hurikyuu > 10) errorMSG += "振休オーバーです" + nl;

            SyoteiDays sd = new SyoteiDays();

            decimal restall = koukyuu + houteikyuu + syoteikyuu;
            decimal restall2 = restall + mutodoke + todokede;

            //bool syukkintyouka = false;

            if (sumDays != 0 && kyuuyo == "C1" && syotei + tokkyuu + mutoku == 0)
            {
                if (bikou == "")
                { 
                    errorMSG += "「出勤無の理由」を選択お願いします。";
                }
                else
                {
                    WarningMSG += "【情報】出勤無理由: " + bikou;
                }
            }

            if (kyuuyo == "C1" || kyuuyo == "C2" || kyuuyo == "D1" || kyuuyo == "D2")
            {
                if (roudouD == "0")
                {
                    //所定出勤日数が多い。
                    //if (syotei > sd.SyoteiDay)
                    if (syotei > syoteiday)
                    {
                        if (row["出勤超過理由"].ToString() == "")
                        { 
                            errorMSG += "所定日数が多いです。多い理由を選択してください。該当理由がなければ休日出勤を使用してください。";
                        }
                        else if (row["出勤超過理由"].ToString() == "次月振休取得予定")
                        {
                            //WarningMSG += "【情報】次月に" + (syotei + yuukyuu - sd.SyoteiDay) + "日振休予定";
                            WarningMSG += "【情報】次月に" + (syotei + yuukyuu - syoteiday) + "日振休予定";
                        }
                        else if (row["出勤超過理由"].ToString() == "シフトで1日8時間未満勤務がある現場")
                        {
                            //syukkintyouka = true;
                            //WarningMSG += "シフトで1日8時間未満勤務がある現場で間違いないですか？　所定出勤の労働時間はトータルで" + (sd.SyoteiDay * 8).ToString() + "時間で間違いないですか？" + nl;
                            WarningMSG += "シフトで1日8時間未満勤務がある現場で間違いないですか？　所定出勤の労働時間はトータルで" + (syoteiday * 8).ToString() + "時間で間違いないですか？" + nl;
                        }

                    }

                    //休日が多い
                    //if (restall > sd.KyuuzitsuDay)
                    if (restall > kyuuzitsuday)
                    {
                        if (row["休日超過理由"].ToString() == "")
                        {
                            //errorMSG += "休日が" + (restall - sd.KyuuzitsuDay).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                            errorMSG += "休日が" + (restall - kyuuzitsuday).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                        }
                        else if (row["休日超過理由"].ToString() == "土日祝日年始が休日")
                        {
                            errorMSG += "休日区分が誤ってます。システム管理者へ問合せ願います。" + nl;

                            //if (restall > sd.DonichiDay)
                            //{
                            //    errorMSG += "土日祝日年始が休日としても" + (restall - sd.DonichiDay).ToString() + "日多いです。" + nl;
                            //}
                            //else
                            //{
                            //    //WarningMSG += "土日祝日年始が休みで間違いないですか？" + nl;
                            //}
                        }
                        else if (row["休日超過理由"].ToString() == "シフトで1日8時間超勤務がある現場")
                        {
                            //WarningMSG += "シフトで1日8時間超勤務がある現場で間違いないですか？　" + ((restall - sd.KyuuzitsuDay) * 8).ToString() + "時間が所定出勤での合計超過時間で間違いないですか？" + nl;
                            WarningMSG += "シフトで1日8時間超勤務がある現場で間違いないですか？　" + ((restall - kyuuzitsuday) * 8).ToString() + "時間が所定出勤での合計超過時間で間違いないですか？" + nl;
                        }
                    }
                }
                else if (roudouD == "1")
                {
                    if (row["休日超過理由"].ToString() == "")
                    {
                        //if (restall > sd.KyuuzitsuDay + 4) errorMSG += "休日が" + (restall - sd.KyuuzitsuDay - 4).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                        if (restall > kyuuzitsuday + 4) errorMSG += "休日が" + (restall - kyuuzitsuday - 4).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                    }
                }
                else if (roudouD == "2")
                {
                    //if (restall > sd.KyuuzitsuDay + 9)
                    if (restall > kyuuzitsuday + 9)
                    {
                        if (row["休日超過理由"].ToString() == "")
                        {
                            //errorMSG += "休日が" + (restall - sd.KyuuzitsuDay - 9).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                            errorMSG += "休日が" + (restall - kyuuzitsuday - 9).ToString() + "日多いです。多い理由を選択してください。該当理由がなければ欠勤を使用してください。" + nl;
                        }
                    }
                }
            }
            else if (kyuuyo == "E1") //パート
            {
                if (roudouD == "0")
                {
                    if (restall >= sd.Kizyun4 && restall < sd.Kizyun3)
                    {
                        errorMSG += "週労5日以上ですが、週労4日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun3 && restall < sd.Kizyun2)
                    {
                        errorMSG += "週労5日以上ですが、週労3日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun2 && restall < sd.Kizyun1)
                    {
                        errorMSG += "週労5日以上ですが、週労2日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun1)
                    {
                        errorMSG += "週労5日以上ですが、週労1日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                }
                else if (roudouD == "1")
                {
                    if (restall >= sd.Kizyun3 && restall < sd.Kizyun2)
                    {
                        errorMSG += "週労4日ですが、週労3日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun2 && restall < sd.Kizyun1)
                    {
                        errorMSG += "週労4日ですが、週労2日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun1)
                    {
                        errorMSG += "週労4日ですが、週労1日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }

                    //休日が少ない
                    if (sumDays != 0 && totalDay == td.TDays && restall2 <= sd.Kizyun5)
                    {
                        WarningMSG += "週労4日ですが、週労5日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                }
                else if (roudouD == "2")
                {
                    if (restall >= sd.Kizyun2 && restall < sd.Kizyun1)
                    {
                        errorMSG += "週労3日ですが、週労2日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }
                    else if (restall >= sd.Kizyun1)
                    {
                        errorMSG += "週労3日ですが、週労1日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }

                    //休日が少ない
                    if (totalDay == td.TDays && restall2 <= sd.Kizyun5 && restall2 > sd.Kizyun4)
                    {
                        WarningMSG += "週労3日ですが、週労5日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun4 && restall2 > sd.Kizyun3)
                    {
                        WarningMSG += "週労3日ですが、週労4日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                }
                else if (roudouD == "3")
                {
                    if (restall >= sd.Kizyun1)
                    {
                        errorMSG += "週労2日ですが、週労1日の基準休日数以上です。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;
                    }

                    //休日が少ない
                    if (totalDay == td.TDays && restall2 <= sd.Kizyun5 && restall2 > sd.Kizyun4)
                    {
                        WarningMSG += "週労2日ですが、週労5日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun4 && restall2 > sd.Kizyun3)
                    {
                        WarningMSG += "週労2日ですが、週労4日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun3 && restall2 > sd.Kizyun2)
                    {
                        WarningMSG += "週労2日ですが、週労3日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }

                }
                else if (roudouD == "4")
                {
                    if (restall > sd.Kizyun1) errorMSG += "休日が基準休日数より" + (restall - sd.Kizyun1).ToString() + "日多いです。有給付与の対象判断に影響しますので届出欠勤に振替ください。週労の変更もご検討ください。" + nl;


                    //休日が少ない
                    if (totalDay == td.TDays && restall2 <= sd.Kizyun5 && restall2 > sd.Kizyun4)
                    {
                        WarningMSG += "週労1日ですが、週労5日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun4 && restall2 > sd.Kizyun3)
                    {
                        WarningMSG += "週労1日ですが、週労4日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun3 && restall2 > sd.Kizyun2)
                    {
                        WarningMSG += "週労1日ですが、週労3日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                    else if (totalDay == td.TDays && restall2 <= sd.Kizyun2 && restall2 > sd.Kizyun1)
                    {
                        WarningMSG += "週労1日ですが、週労3日の基準休日数以下です。有給日数が少なく付与される可能性があります。今後も同様の勤務形態なら週労の変更をご検討ください。" + nl;
                    }
                }
            }

            if (todokede > totalDay) errorMSG += "届出オーバーです" + nl;
            if (mutodoke > totalDay) errorMSG += "無届オーバーです" + nl;
            if (kaisuu1 > totalDay) errorMSG += "回数1オーバーです" + nl;

            //TODO202012 追加
            if (kaisuu2 > 2) errorMSG += "回数2オーバーです" + nl;
            //TODO202012 コメントアウト
            //if (kaisuu2 > totalDay) errorMSG += "回数2オーバーです" + nl;


            if (sumDays != 0 && sumDays != totalDay) errorMSG += "日数過不足です" + nl;

            //小数点以下エラー　小数点以下がありえない数字です。
            if ((entyouH * 10 % 5) > 0) errorMSG += "延長時間小終点エラー" + nl;
            if ((houteiH * 10 % 5) > 0) errorMSG += "法休時間小終点エラー" + nl;
            if ((syoteiH * 10 % 5) > 0) errorMSG += "所休時間小終点エラー" + nl;
            if ((souzanH * 10 % 5) > 0) errorMSG += "総残時間小終点エラー" + nl;
            if ((tyouzanH * 10 % 5) > 0) errorMSG += "60超時間小終点エラー" + nl;
            if ((shinyaH * 10 % 5) > 0) errorMSG += "深夜時間小終点エラー" + nl;
            if ((tikokuH * 10 % 5) > 0) errorMSG += "遅刻時間小終点エラー" + nl;

            //所休・法休エラー
            if (syoteikyuu == 0 & syoteiH > 0) errorMSG += "所定休の日数未入力" + nl;
            if (houteikyuu == 0 & houteiH > 0) errorMSG += "法定休の日数未入力" + nl;
            if (syoteikyuu > 0 & syoteiH == 0) errorMSG += "所定休の時間未入力" + nl;
            if (houteikyuu > 0 & houteiH == 0) errorMSG += "法定休の時間未入力" + nl;

            if (syoteikyuu * 24 < syoteiH) errorMSG += "所定休　時間or日数誤り" + nl;
            if (houteikyuu * 24 < houteiH) errorMSG += "法定休　時間or日数誤り" + nl;

            //60超
            //if (souzanH == 0 & tyouzanH > 0) errorMSG += "60超ありで残業なし" + nl;
            //if (souzanH > 60 & ((souzanH - 60) != tyouzanH)) errorMSG += "６０超残を入力しでください" + nl;

            //出勤なしで～
            if (sumDays_sub1 == 0 & entyouH > 0) errorMSG += "出勤なしで延長あり" + nl;
            if (sumDays_sub1 == 0 & souzanH > 0) errorMSG += "出勤なしで残業あり" + nl;
            if (sumDays_sub1 == 0 & shinyaH > 0) errorMSG += "出勤なしで深夜あり" + nl;
            if (sumDays_sub1 == 0 & tikoku > 0) errorMSG += "出勤なしで遅刻あり" + nl;

            //その他
            if (kaisuu1 > 0 & tanka1 == 0) errorMSG += "夜勤単価未登録" + nl;

            //if (kaisuu2 > 0 & tanka2 == 0) errorMSG += "宿直単価未登録" + nl;
            
            
            //if (tikoku == 0 & tikokuH > 0) errorMSG += "遅刻回数エラー" + nl;
            //if (tikoku * kinmuh < tikokuH) errorMSG += "遅刻時間エラー" + nl;

            //if (kyuuyo == "F1" & tyoukyuu > 0) errorMSG += "バイトに調休" + nl;
            if (kyuuyo != "E1" & kyuuyo != "F1" & entyouH > 0) errorMSG += "日・月給者に延長" + nl;

            //時間警告 TODO:
            if (entyouH >= 100) WarningMSG += "【警告】延長100時間OK?" + nl;
            if (houteiH >= 50) WarningMSG += "【警告】法休50時間以上OK?" + nl;
            if (syoteiH >= 50) WarningMSG += "【警告】所休50時間以上OK?" + nl;
            if (souzanH >= 100) WarningMSG += "【警告】総残100時間OK?" + nl;
            if (shinyaH >= 200) WarningMSG += "【警告】深夜200時間OK?" + nl;

            //八重山限定
            if (row[31].ToString() == "八重山" && tanka1 > 0 && syotei > 0 && kaisuu1 == 0)
            {
                WarningMSG += "【警告】回数1単価設定有で所定出勤有だけど回数1が0" + nl;
            }

            if (mutoku + todokede + mutodoke > 2 & kyuuyo != "E1" & kyuuyo != "F1" & lbl == "")
            {
                if (syotei + yuukyuu + tokkyuu <= 21)
                { 
                    if (row["無特理由"].ToString() == "産前産後育休")
                    {
                        WarningMSG += "【情報】産休・育休中" + nl;
                    }
                    else
                    { 
                        WarningMSG += "【警告】出勤日数不足のため、手当が日割計算となります" + nl;
                    }
                }
                else
                {
                    WarningMSG += "【例外】想定外の勤怠" + nl;
                }

            }

            //所定日数達してないけど休日出勤発生
            //if (syotei + yuukyuu + mutoku + tokkyuu + mutodoke + todokede + hurikyuu < sd.SyoteiDay && (houteikyuu > 0 || syoteikyuu > 0))
            if (syotei + yuukyuu + mutoku + tokkyuu + mutodoke + todokede + hurikyuu < syoteiday && (houteikyuu > 0 || syoteikyuu > 0))
            { 
                    WarningMSG += "【情報】所定日数達してないけど休日出勤有。" + nl;
            }


            //実労働時間の表示
            decimal dec = entyouH + syotei * kinmuh + houteiH + syoteiH + souzanH - tikokuH;

            //所定労働時間
            decimal syor = entyouH + syotei * kinmuh - tikokuH;

            //roudouH.Text += Convert.ToString(dec);
            //roudouH.Text += "時間";

            //TODO 
            // if (dec > 130 & syaho == 9) WarningMSG += "【警告】社保未で労働130H超 " + nl + "総労働：" + dec + "時間" + nl;

            if (dec > 250) WarningMSG += "【警告】総労働 " + dec + "時間 OK?" + nl;

            //TODO
            if (syotei == td.TDays) WarningMSG += "休日無OK？" + nl;

            if (sumDays != 0 && syotei + yuukyuu + tokkyuu == 0 && jyuumin > 0) WarningMSG += "【警告】支給額ゼロで住民税あり" + nl;

            //係長以上は休出/残業/遅刻は警告を表示
            if (yakusyoku <= 135)
            {
                if (houteiH > 0) WarningMSG += "【警告】係長以上に法休" + nl;
                if (syoteiH > 0) WarningMSG += "【警告】係長以上に所休" + nl;
                if (souzanH > 0) WarningMSG += "【警告】係長以上に残業" + nl;
                if (shinyaH > 0) WarningMSG += "【警告】係長以上に深夜" + nl;
                if (tikokuH > 0) errorMSG += "【警告】係長以上に遅刻" + nl;
            }

            if (yakusyoku > 135)
            {
                //if (syor > sd.RoudouH)
                if (syor > roudouh)
                {
                    //if (syukkintyouka)
                    if (row["出勤超過理由"].ToString() != "")
                    {
                        //シフト勤務
                    }
                    else
                    {
                        //WarningMSG += "【警告】" + "所定労働時間オーバー" + Convert.ToString(syor - sd.RoudouH) + "時間分の割増賃金が必要" + nl;
                        WarningMSG += "【警告】" + "所定労働時間オーバー" + Convert.ToString(syor - roudouh) + "時間分の割増賃金が必要" + nl;
                    }
                }
            }

            if (kyuuyo != "F1" && flg != "Up")
            {
                //休日オーバー警告が設定されていた。
            }

            //特休/無特　使用理由選択
            if (Convert.ToInt16(tokkyuu) > 0 && row["特休理由"].ToString() == "")
            {
                errorMSG += "特休理由を選択してください" + nl; 
            }

            if (Convert.ToInt16(tokkyuu) == 0 && row["特休理由"].ToString() != "")
            {
                errorMSG += "特休理由を選択してるけど入力無" + nl;
            }

            if (Convert.ToInt16(mutoku) > 0 && row["無特理由"].ToString() == "")
            {
                errorMSG += "無特理由を選択してください" + nl; 
            }

            if (Convert.ToInt16(mutoku) == 0 && row["無特理由"].ToString() != "")
            {
                errorMSG += "無特理由を選択してるけど入力無" + nl;
            }


            #endregion

            string[] strRe = new string[5];
            strRe[0] = errorMSG; //textBox23
            strRe[1] = totalDay.ToString(); //label34
            strRe[2] = sumDays.ToString();  //label32
            strRe[3] = lbl;  //label36
            strRe[4] = WarningMSG;  //label36

            return strRe;
            #endregion
        }


        //DBより各データを取得
        public DataTable GetKintaiKihon(int flg, string str)
        {
            SqlConnection Cn;
            SqlCommand Cmd;
            DataTable dt = new DataTable();
            SqlDataAdapter da;

            string strd;

            try
            {
                if (flg == 1)
                {
                    //対象者データ取得
                    strd = "[dbo].[k勤怠社員情報取得]";
                }
                //else if (flg == 2)
                //{
                //    //地区制限フラグ取得
                //    strd = "[dbo].[地区制限フラグ取得]";
                //}
                //else if (flg == 3)
                //{
                //    //ログイン情報取得
                //    strd = "[dbo].[ログイン情報取得]";
                //}
                else if (flg == 4)
                {
                    //年休等のサブ情報を取得
                    strd = "[dbo].[s社員サブ情報取得]";
                }
                else if (flg == 5)
                {
                    //ZeeMの勤怠データ取得
                    //TODO つかってない
                    strd = "[dbo].[勤怠社員情報取得_ZEEM]";
                }
                else if (flg == 9)
                {
                    //担当別一覧データの取得
                    strd = "[dbo].[担当別一覧表示]";
                }
                else
                {
                    MessageBox.Show("エラー。管理者へ問合せ下さい");
                    strd = "";
                }

                using (Cn = new SqlConnection(constr))
                {
                    Cn.Open();

                    using (Cmd = Cn.CreateCommand())
                    {
                        Cmd.CommandType = CommandType.StoredProcedure;
                        Cmd.CommandText = strd;

                        if (flg == 4 || flg == 5)
                        {
                            Cmd.Parameters.Add(new SqlParameter("year", SqlDbType.VarChar));
                            Cmd.Parameters["year"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("month", SqlDbType.VarChar));
                            Cmd.Parameters["month"].Direction = ParameterDirection.Input;

                            TargetDays td = new TargetDays();
                            Cmd.Parameters["year"].Value = td.StartYMD.AddMonths(1).Year.ToString();
                            Cmd.Parameters["month"].Value = td.StartYMD.AddMonths(1).ToString("MM");
                        }
                        else if (flg == 9)
                        {
                            Cmd.Parameters.Add(new SqlParameter("tiku", SqlDbType.VarChar));
                            Cmd.Parameters["tiku"].Direction = ParameterDirection.Input;

                            Cmd.Parameters.Add(new SqlParameter("name", SqlDbType.VarChar));
                            Cmd.Parameters["name"].Direction = ParameterDirection.Input;

                            Cmd.Parameters["tiku"].Value = Program.loginbusyo;
                            //Cmd.Parameters["busyo"].Value = kintai.Program.loginbusyo;
                            Cmd.Parameters["name"].Value = str;
                        }

                        da = new SqlDataAdapter(Cmd);
                        da.Fill(dt);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー" + ex.ToString());
                throw;
            }
        }
    }
}
