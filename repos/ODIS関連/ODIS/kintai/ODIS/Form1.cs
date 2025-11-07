using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic; // Interaction.InputBox
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ODIS
{
    public partial class Form1 : Form
    {
        // === 必要な状態 ===
        string fileyymm = "";  // YYYYMM
        string nendo = "";     // 年度（4月起算）
        string bmntk = "";     // 地区キー（部門頭1桁）
        string bmnskKey = "";  // 職種キー（部門末尾4桁）
        string tiku = "";
        string syoku = "";
        string parcent = "";
        int maxretu = 0;

        // 配列（0/1は使わない、2〜13 を使用：VBA準拠）
        long[] kturg1 = new long[14];
        long[] rgurg1 = new long[14];
        long[] bpurg1 = new long[14];
        long[] keihg1 = new long[14];
        long[] jinkg1 = new long[14];

        long[] kturg2 = new long[14];
        long[] rgurg2 = new long[14];
        long[] bpurg2 = new long[14];
        long[] keihg2 = new long[14];
        long[] jinkg2 = new long[14];

        long[] kturg3 = new long[14];
        long[] rgurg3 = new long[14];
        long[] bpurg3 = new long[14];
        long[] keihg3 = new long[14];
        long[] jinkg3 = new long[14];

        // 取込データ
        List<CsvRow> rows = new List<CsvRow>();

        public Form1()
        {
            InitializeComponent(); // デザイナ使用
        }

        // 取込ボタン
        private void btnImport_Click(object sender, EventArgs e) { DoImport(); }

        // ============ 文字列ヘルパ（.NET Framework互換） ============
        static string LeftStr(string s, int count)
        {
            if (string.IsNullOrEmpty(s) || count <= 0) return "";
            return s.Length <= count ? s : s.Substring(0, count);
        }
        static string RightStr(string s, int count)
        {
            if (string.IsNullOrEmpty(s) || count <= 0) return "";
            return s.Length <= count ? s : s.Substring(s.Length - count);
        }

        // ============ データ取込（Excel 出力のみ） ============
        void DoImport()
        {
            try
            {
                // 1) 入力：処理年月度 YYYYMM
                fileyymm = Interaction.InputBox(
                    "現場計数データ年月度を入力して下さい（西暦年4桁+月2桁）",
                    "処理年月度指定", DateTime.Now.ToString("yyyyMM"));

                if (string.IsNullOrWhiteSpace(fileyymm))
                {
                    MessageBox.Show("取り込み処理をキャンセルします");
                    return;
                }

                int tmpInt;
                if (!int.TryParse(fileyymm, out tmpInt) || fileyymm.Length != 6)
                {
                    MessageBox.Show("YYYYMM 形式で入力してください。");
                    return;
                }

                // 2) 年度（4月起算）
                int y = int.Parse(fileyymm.Substring(0, 4));
                int m = int.Parse(fileyymm.Substring(4, 2));
                nendo = (m < 4 ? (y - 1) : y).ToString();

                // 3) CSV読み込み（cp932優先／BOM自動判別）
                // UNC 直下と「データ」サブフォルダを両方試す
                string basePath = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\計数\Excel計数集計\";
                string fileName = "keisu2g-" + nendo + ".csv";
                string cand1 = Path.Combine(basePath, fileName);
                string cand2 = Path.Combine(basePath, "データ", fileName);
                string csvPath = File.Exists(cand1) ? cand1 : (File.Exists(cand2) ? cand2 : null);

                if (string.IsNullOrEmpty(csvPath))
                {
                    string msg = "計数ファイルが見つかりません。\n\n試した場所：\n" + cand1 + "\n" + cand2 +
                                 "\n\n指定年月度: " + fileyymm + "  → 年度(nendo): " + nendo +
                                 "\n（例：2025/03 を指定すると nendo=2024 を探します）";
                    throw new FileNotFoundException(msg);
                }

                rows = ReadCsv(csvPath);

                // 4) 2017/04〜 米軍施設/プロジェクトの仮コード変換
                if (int.Parse(fileyymm) >= 201704)
                {
                    for (int i = 0; i < rows.Count; i++)
                    {
                        if (rows[i].BumonCd == "22031") rows[i].BumonCd = "2202A";
                        else if (rows[i].BumonCd == "22032") rows[i].BumonCd = "2202B";
                    }
                }

                // 5) ソート（部門CD→現場CD→年月）
                rows = rows
                    .OrderBy(r => r.BumonCd)
                    .ThenBy(r => r.GenbaCd)
                    .ThenBy(r => r.Yyyymm)
                    .ToList();

                // 6) 累計作成（部門×現場で指定月まで）
                Dictionary<Tuple<string, string>, CumRow> cumTable = BuildCumulative(rows, int.Parse(fileyymm));

                // 7) 抽出％入力と絞り込み/削除
                string inPct = Interaction.InputBox("絞り込む計数を指定して下さい（単位：％）", "抽出計数指定", "85");
                decimal pct;
                if (!decimal.TryParse(inPct, out pct) || pct < -100 || pct >= 1000)
                {
                    MessageBox.Show("入力エラー（-100〜999.99）。再実行してください。");
                    return;
                }
                parcent = inPct;
                decimal keisu = pct / 100m;

                // 指定計数未満は除外・指定月以降は削除
                List<CsvRow> filtered = FilterByKeisu(rows, cumTable, keisu, int.Parse(fileyymm));

                // 8) Excel テンプレへ自動出力（高速＆書式維持）
                string templatePath = Path.Combine(basePath, "現場計数.xlsx");
                string savePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "現場計数-" + parcent + "-" + fileyymm + ".xlsx"
                );
                ExportExcelByTemplateAdvanced(filtered, fileyymm, nendo, templatePath, savePath);

                MessageBox.Show("Excel に出力しました。\n" + savePath);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("ファイルが見つかりません: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー: " + ex.Message + "\n処理はキャンセルしました。");
            }
        }

        // ================= CSV 読み込み =================

        List<CsvRow> ReadCsv(string path)
        {
            var list = new List<CsvRow>();
            using (var sr = new StreamReader(path, Encoding.GetEncoding(932), true))
            {
                string line;
                int lineno = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    lineno++;
                    if (lineno == 1) continue; // ヘッダスキップ

                    var cols = SplitCsv(line);
                    if (cols.Length < 44) continue; // CSVは44列（index 0..43）

                    var r = new CsvRow(cols);
                    list.Add(r);
                }
            }
            return list;
        }

        string[] SplitCsv(string line)
        {
            var res = new List<string>();
            var sb = new StringBuilder();
            bool inQ = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQ && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"'); i++;
                    }
                    else inQ = !inQ;
                }
                else if (c == ',' && !inQ)
                {
                    res.Add(sb.ToString()); sb.Clear();
                }
                else sb.Append(c);
            }
            res.Add(sb.ToString());
            return res.ToArray();
        }

        class CsvRow
        {
            public string BumonCd;
            public string GenbaCd;
            public int Yyyymm;
            public string GenbaName;

            public long FixUriage;     // 5 + 7
            public long RinjiUriage;   // 6 + 8 + 10
            public long BuppinUriage;  // 9
            public long Jinken;        // 23〜25
            public long KeihiSum;      // 30〜42 + 44

            public string[] Raw;

            public CsvRow(string[] cols)
            {
                Raw = cols;
                BumonCd = Safe(cols, 0);
                GenbaCd = Safe(cols, 1);
                Yyyymm = ToInt(Safe(cols, 2), 0);
                GenbaName = Safe(cols, 3).Trim();

                long c5 = ToLong(Safe(cols, 4));
                long c6 = ToLong(Safe(cols, 5));
                long c7 = ToLong(Safe(cols, 6));
                long c8 = ToLong(Safe(cols, 7));
                long c9 = ToLong(Safe(cols, 8));
                long c10 = ToLong(Safe(cols, 9));

                FixUriage = c5 + c7;
                RinjiUriage = c6 + c8 + c10;
                BuppinUriage = c9;

                long c23 = ToLong(Safe(cols, 22));
                long c24 = ToLong(Safe(cols, 23));
                long c25 = ToLong(Safe(cols, 24));
                Jinken = c23 + c24 + c25;

                long keihi = 0;
                for (int i = 29; i <= 41; i++) keihi += ToLong(Safe(cols, i)); // 30〜42
                long lease = ToLong(Safe(cols, 43)); // 44
                KeihiSum = keihi + lease;
            }

            static string Safe(string[] a, int i)
            {
                if (i >= 0 && i < a.Length) return a[i];
                return "";
            }
            static int ToInt(string s, int def)
            {
                int v;
                if (int.TryParse(s, out v)) return v;
                return def;
            }
            static long ToLong(string s)
            {
                if (string.IsNullOrEmpty(s)) return 0L;
                s = s.Replace(",", "");
                long v;
                if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
                return 0L;
            }
        }

        // ============ 累計作成 & 絞り込み ============

        class CumRow
        {
            public string BumonCd;
            public string GenbaCd;
            public int LatestYyyymm;
            public string GenbaName;
            public long FixUriageSum;    // 57
            public long RinjiUriageSum;  // 58
            public long BuppinUriageSum; // 59
            public long JinkenSum;       // 60
            public long KeihiSum;        // 61
            public decimal? Keisu;       // 62
        }

        Dictionary<Tuple<string, string>, CumRow> BuildCumulative(List<CsvRow> data, int fileYyyymm)
        {
            var dict = new Dictionary<Tuple<string, string>, CumRow>();
            for (int i = 0; i < data.Count; i++)
            {
                var r = data[i];
                var key = Tuple.Create(r.BumonCd, r.GenbaCd);
                CumRow cr;
                if (!dict.TryGetValue(key, out cr))
                {
                    cr = new CumRow
                    {
                        BumonCd = r.BumonCd,
                        GenbaCd = r.GenbaCd,
                        GenbaName = r.GenbaName,
                        LatestYyyymm = 0
                    };
                    dict[key] = cr;
                }

                if (r.Yyyymm <= fileYyyymm)
                {
                    cr.LatestYyyymm = r.Yyyymm;
                    cr.FixUriageSum += r.FixUriage;
                    cr.RinjiUriageSum += r.RinjiUriage;
                    cr.BuppinUriageSum += r.BuppinUriage;
                    cr.JinkenSum += r.Jinken;
                    cr.KeihiSum += r.KeihiSum;
                }
            }

            foreach (var kv in dict)
            {
                var cr = kv.Value;
                long u = cr.FixUriageSum + cr.RinjiUriageSum + cr.BuppinUriageSum;
                long k = cr.JinkenSum + cr.KeihiSum;
                if (k <= 0 && u == 0) cr.Keisu = 0m;
                else if (k <= 0 && u != 0) cr.Keisu = 0m;
                else if (k != 0 && u == 0) cr.Keisu = 1m;
                else cr.Keisu = u == 0 ? (decimal?)null : (decimal)k / (decimal)u;
            }

            return dict;
        }

        List<CsvRow> FilterByKeisu(List<CsvRow> data, Dictionary<Tuple<string, string>, CumRow> cum, decimal keisu, int fileYyyymm)
        {
            var allow = new HashSet<string>();
            foreach (var kv in cum)
            {
                var cr = kv.Value;
                decimal val = cr.Keisu.HasValue ? cr.Keisu.Value : 0m;
                if (val >= keisu) allow.Add(cr.BumonCd + "|" + cr.GenbaCd);
            }

            var filtered = data.Where(r => allow.Contains(r.BumonCd + "|" + r.GenbaCd)).ToList();
            filtered = filtered.Where(r => r.Yyyymm <= fileYyyymm).ToList();

            return filtered;
        }

        // ================= Excel 出力（テンプレ書式維持 & 高速化） =================

        enum BlockType { Genba, ShokushuTotal, ChikuTotal, SogoTotal }
        class OutBlock
        {
            public BlockType Type;
            public string Tiku;
            public string Syoku;
            public string Genba;
            public long[] Fix = new long[12];
            public long[] Rin = new long[12];
            public long[] Bp = new long[12];
            public long[] Kei = new long[12];
            public long[] Jin = new long[12];
        }

        List<OutBlock> BuildBlocksForExport(List<CsvRow> data)
        {
            var list = new List<OutBlock>();

            string curBmn = "", curGenba = "";
            string curGenbaName = "";
            string curTiku = "", curSyoku = "";
            string curBmnSk = "";
            string curBmnTk = "";

            ResetArrays(kturg1, rgurg1, bpurg1, keihg1, jinkg1);
            ResetArrays(kturg2, rgurg2, bpurg2, keihg2, jinkg2);
            ResetArrays(kturg3, rgurg3, bpurg3, keihg3, jinkg3);

            long[] g_fix = new long[12], g_rin = new long[12], g_bp = new long[12], g_kei = new long[12], g_jin = new long[12];

            var ordered = data
                .OrderBy(r => r.BumonCd).ThenBy(r => r.GenbaCd).ThenBy(r => r.Yyyymm).ToList();

            Action flushGenba = () =>
            {
                if (curGenba == "") return;
                var b = new OutBlock { Type = BlockType.Genba, Tiku = curTiku, Syoku = curSyoku, Genba = curGenbaName };
                Array.Copy(g_fix, b.Fix, 12);
                Array.Copy(g_rin, b.Rin, 12);
                Array.Copy(g_bp, b.Bp, 12);
                Array.Copy(g_kei, b.Kei, 12);
                Array.Copy(g_jin, b.Jin, 12);
                list.Add(b);

                for (int mIdx = 0; mIdx < 12; mIdx++)
                {
                    int ret = (mIdx <= 8) ? (mIdx + 2) : (mIdx - 10 + 12);
                    kturg1[ret] += g_fix[mIdx];
                    rgurg1[ret] += g_rin[mIdx];
                    bpurg1[ret] += g_bp[mIdx];
                    keihg1[ret] += g_kei[mIdx];
                    jinkg1[ret] += g_jin[mIdx];
                }

                Array.Clear(g_fix, 0, 12); Array.Clear(g_rin, 0, 12); Array.Clear(g_bp, 0, 12);
                Array.Clear(g_kei, 0, 12); Array.Clear(g_jin, 0, 12);
            };

            Action flushShokushu = () =>
            {
                if (curBmn == "") return;
                if (BmnSkSuppressTotal(curBmnSk))
                {
                    for (int ret = 2; ret <= 13; ret++)
                        kturg1[ret] = rgurg1[ret] = bpurg1[ret] = keihg1[ret] = jinkg1[ret] = 0;
                    return;
                }
                var b = new OutBlock { Type = BlockType.ShokushuTotal, Tiku = curTiku, Syoku = curSyoku };
                for (int ret = 2; ret <= 13; ret++)
                {
                    int mIdx = (ret <= 11) ? (ret - 2) : (ret - 14 + 12);
                    b.Fix[mIdx] = kturg1[ret]; b.Rin[mIdx] = rgurg1[ret]; b.Bp[mIdx] = bpurg1[ret];
                    b.Kei[mIdx] = keihg1[ret]; b.Jin[mIdx] = jinkg1[ret];
                }
                list.Add(b);
                for (int ret = 2; ret <= 13; ret++)
                {
                    kturg2[ret] += kturg1[ret]; rgurg2[ret] += rgurg1[ret]; bpurg2[ret] += bpurg1[ret];
                    keihg2[ret] += keihg1[ret]; jinkg2[ret] += jinkg1[ret];
                    kturg1[ret] = rgurg1[ret] = bpurg1[ret] = keihg1[ret] = jinkg1[ret] = 0;
                }
            };

            Action flushChiku = () =>
            {
                if (curBmn == "") return;
                var b = new OutBlock { Type = BlockType.ChikuTotal, Tiku = GetChikuLabel(curBmnTk) };
                for (int ret = 2; ret <= 13; ret++)
                {
                    int mIdx = (ret <= 11) ? (ret - 2) : (ret - 14 + 12);
                    b.Fix[mIdx] = kturg2[ret]; b.Rin[mIdx] = rgurg2[ret]; b.Bp[mIdx] = bpurg2[ret];
                    b.Kei[mIdx] = keihg2[ret]; b.Jin[mIdx] = jinkg2[ret];
                }
                list.Add(b);
                for (int ret = 2; ret <= 13; ret++)
                {
                    kturg3[ret] += kturg2[ret]; rgurg3[ret] += rgurg2[ret]; bpurg3[ret] += bpurg2[ret];
                    keihg3[ret] += keihg2[ret]; jinkg3[ret] += jinkg2[ret];
                    kturg2[ret] = rgurg2[ret] = bpurg2[ret] = keihg2[ret] = jinkg2[ret] = 0;
                }
            };

            for (int i = 0; i < ordered.Count; i++)
            {
                var r = ordered[i];
                bool changedGenba = (r.BumonCd != curBmn) || (r.GenbaCd != curGenba);
                if (changedGenba)
                {
                    flushGenba();

                    if (!string.IsNullOrEmpty(curBmn) && r.BumonCd.Length >= 4 && RightStr(r.BumonCd, 4) != curBmnSk)
                        flushShokushu();

                    if (!string.IsNullOrEmpty(curBmn) && LeftStr(r.BumonCd, 1) != curBmnTk)
                    {
                        flushChiku();
                        curBmnTk = LeftStr(r.BumonCd, 1);
                    }

                    curBmn = r.BumonCd;
                    curGenba = r.GenbaCd;

                    // ★ここを追加：現場名を保持（名称が空ならCDでフォールバック）
                    curGenbaName = string.IsNullOrWhiteSpace(r.GenbaName) ? r.GenbaCd : r.GenbaName;

                    curBmnSk = (r.BumonCd.Length >= 4) ? RightStr(r.BumonCd, 4) : "";
                    curBmnTk = LeftStr(r.BumonCd, 1);
                    GetChikuShokushuName(r.BumonCd, out curTiku, out curSyoku);
                }

                int colIndex2 = MonthToIndex(r.Yyyymm);
                int mIdx2 = (colIndex2 <= 11) ? (colIndex2 - 2) : (colIndex2 - 14 + 12);
                g_fix[mIdx2] += r.FixUriage;
                g_rin[mIdx2] += r.RinjiUriage;
                g_bp[mIdx2] += r.BuppinUriage;
                g_kei[mIdx2] += r.KeihiSum;
                g_jin[mIdx2] += r.Jinken;
            }
            flushGenba();
            flushShokushu();
            flushChiku();

            var z = new OutBlock { Type = BlockType.SogoTotal, Tiku = "◆◇◆　総　合　計　◆◇◆" };
            for (int ret = 2; ret <= 13; ret++)
            {
                int mIdx = (ret <= 11) ? (ret - 2) : (ret - 14 + 12);
                z.Fix[mIdx] = kturg3[ret]; z.Rin[mIdx] = rgurg3[ret]; z.Bp[mIdx] = bpurg3[ret];
                z.Kei[mIdx] = keihg3[ret]; z.Jin[mIdx] = jinkg3[ret];
            }
            list.Add(z);

            return list;
        }

        // 画面更新などの高速化（Calculation は触らない）
        void WithExcelFast(Excel.Application app, Action action)
        {
            bool? screenUpdating = null;
            bool? displayAlerts = null;
            bool? enableEvents = null;

            try
            {
                try { screenUpdating = app.ScreenUpdating; app.ScreenUpdating = false; } catch { }
                try { displayAlerts = app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
                try { enableEvents = app.EnableEvents; app.EnableEvents = false; } catch { }

                action();
            }
            finally
            {
                try { if (enableEvents.HasValue) app.EnableEvents = enableEvents.Value; } catch { }
                try { if (displayAlerts.HasValue) app.DisplayAlerts = displayAlerts.Value; } catch { }
                try { if (screenUpdating.HasValue) app.ScreenUpdating = screenUpdating.Value; } catch { }
            }
        }

        int FindFirstHeaderRow(Excel.Worksheet ws)
        {
            for (int r = 1; r <= 300; r++)
            {
                string n = Convert.ToString((ws.Cells[r, 14] as Excel.Range).Value2); // N列=総合計
                string o = Convert.ToString((ws.Cells[r, 15] as Excel.Range).Value2); // O列=第一四半期
                if (!string.IsNullOrEmpty(n) && n.IndexOf("総合計") >= 0 &&
                    !string.IsNullOrEmpty(o) && o.IndexOf("第一四半期") >= 0)
                {
                    return r;
                }
            }
            return -1;
        }

        int FindBlockHeight(Excel.Worksheet ws, int headerRow)
        {
            for (int r = headerRow + 1; r <= headerRow + 50; r++)
            {
                string n = Convert.ToString((ws.Cells[r, 14] as Excel.Range).Value2);
                string o = Convert.ToString((ws.Cells[r, 15] as Excel.Range).Value2);
                if (!string.IsNullOrEmpty(n) && n.IndexOf("総合計") >= 0 &&
                    !string.IsNullOrEmpty(o) && o.IndexOf("第一四半期") >= 0)
                {
                    return r - headerRow;
                }
            }
            return 13; // 既定
        }

        void EnsureBlockCountFast(Excel.Worksheet ws, int headerRow, int blockHeight, int blocksNeeded)
        {
            int count = 1, cursor = headerRow;
            while (true)
            {
                int next = cursor + blockHeight;
                string n = Convert.ToString((ws.Cells[next, 14] as Excel.Range).Value2);
                string o = Convert.ToString((ws.Cells[next, 15] as Excel.Range).Value2);
                if (!string.IsNullOrEmpty(n) && n.IndexOf("総合計") >= 0 &&
                    !string.IsNullOrEmpty(o) && o.IndexOf("第一四半期") >= 0)
                {
                    count++; cursor = next;
                }
                else break;
                if (count > 500) break;
            }

            if (count < blocksNeeded)
            {
                int toAdd = blocksNeeded - count;
                Excel.Range block = ws.Range[ws.Cells[headerRow, 1], ws.Cells[headerRow + blockHeight - 1, 18]];
                int insertAt = headerRow + count * blockHeight;
                ws.Rows[insertAt + ":" + (insertAt + toAdd * blockHeight - 1)].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                block.Copy(ws.Cells[insertAt, 1]);
                ws.Range[
                    ws.Cells[insertAt, 1],
                    ws.Cells[insertAt + toAdd * blockHeight - 1, 18]
                ].FillDown();
            }
            else if (count > blocksNeeded)
            {
                int toDel = count - blocksNeeded;
                int delStart = headerRow + blocksNeeded * blockHeight;
                ws.Rows[delStart + ":" + (delStart + toDel * blockHeight - 1)].Delete();
            }
        }

        // ラベル書き込み（テンプレ準拠・結合セル対応）
        // A=【地　区】, C=【職　種】, E=【現場名】 はテンプレ固定。
        // 表示/非表示を必要に応じて切替える。
        void WriteLabels(Excel.Worksheet ws, int labelRow, int colTiku, int colSyoku, int colGenba, OutBlock b)
        {
            var cellA = (Excel.Range)ws.Cells[labelRow, colTiku - 1]; // A: 【地　区】
            var cellC = (Excel.Range)ws.Cells[labelRow, colSyoku - 1]; // C: 【職　種】
            var cellE = (Excel.Range)ws.Cells[labelRow, colGenba - 1]; // E: 【現場名】

            var cellB = (Excel.Range)ws.Cells[labelRow, colTiku];   // B: 地区名
            var cellD = (Excel.Range)ws.Cells[labelRow, colSyoku];  // D: 職種 or 「◆◇ 職種 合計 ◇◆」
            var cellF = (Excel.Range)ws.Cells[labelRow, colGenba];  // F: 現場名

            void SetFixedHeaders(bool showA, bool showC, bool showE)
            {
                SetMergedValue(cellA, showA ? "【地　区】" : "");
                SetMergedValue(cellC, showC ? "【職　種】" : "");
                SetMergedValue(cellE, showE ? "【現場名】" : "");
            }

            switch (b.Type)
            {
                case BlockType.Genba:
                    // 現場：A/C/E見出しあり、B/D/Fに値
                    SetFixedHeaders(true, true, true);
                    SetMergedValue(cellB, b.Tiku ?? "");
                    SetMergedValue(cellD, b.Syoku ?? "");
                    SetMergedValue(cellF, b.Genba ?? "");
                    break;

                case BlockType.ShokushuTotal:
                    // 職種合計：Aは残す（ご要望）、C/Eは消す。B=地区, D=「◆◇ 職種 合計 ◇◆」, Fは空
                    SetFixedHeaders(true, false, false);
                    SetMergedValue(cellB, b.Tiku ?? "");
                    SetMergedValue(cellD, "◆◇ " + (b.Syoku ?? "") + " 合計 ◇◆");
                    ClearMerged(cellF);
                    break;

                case BlockType.ChikuTotal:
                    // 地区合計：A/C/Eを消す。B=「◆◇◆ 地区 合計 ◆◇◆」, D/Fは空
                    SetFixedHeaders(false, false, false);
                    SetMergedValue(cellB, b.Tiku ?? "");
                    ClearMerged(cellD);
                    ClearMerged(cellF);
                    break;

                case BlockType.SogoTotal:
                    // 総合計：A/C/Eを消す。B=「◆◇◆ 総 合 計 ◆◇◆」, D/Fは空
                    SetFixedHeaders(false, false, false);
                    SetMergedValue(cellB, "◆◇◆　総　合　計　◆◇◆");
                    ClearMerged(cellD);
                    ClearMerged(cellF);
                    break;
            }
        }





        // 結合セルなら結合範囲の先頭セルに値を設定
        void SetMergedValue(Excel.Range cell, string value)
        {
            try
            {
                if (cell != null && cell.MergeCells is bool mc && mc)
                {
                    var area = cell.MergeArea;
                    (area.Cells[1, 1] as Excel.Range).Value2 = value ?? "";
                }
                else
                {
                    cell.Value2 = value ?? "";
                }
            }
            catch
            {
                // 最終手段
                try { cell.Value2 = value ?? ""; } catch { }
            }
        }

        // 結合セルなら結合範囲全体をクリア
        void ClearMerged(Excel.Range cell)
        {
            try
            {
                if (cell != null && cell.MergeCells is bool mc && mc)
                {
                    cell.MergeArea.ClearContents();
                }
                else
                {
                    cell.ClearContents();
                }
            }
            catch
            {
                // 値だけでも空に
                try { cell.Value2 = ""; } catch { }
            }
        }


        // 1ブロック出力（A:R）＋ヘッダ行処理
        void FillBlock(Excel.Worksheet ws, int headerRow, OutBlock b)
        {
            int colTiku = 2; // B
            int colSyoku = 4; // D
            int colGenba = 6; // F

            // ラベルは「年月見出し行」の1つ上の行
            int labelRow = headerRow - 1;

            // ラベルは種類に応じて WriteLabels が担当（Genba は B/D/F のみ）
            WriteLabels(ws, labelRow, colTiku, colSyoku, colGenba, b);

            // 値は見出しの下から 7 行に一括書き込み
            int dataStart = headerRow + 1;
            WriteBlockValues(ws, dataStart, b.Fix, b.Rin, b.Bp, b.Kei, b.Jin);
        }

        void WriteBlockValues(Excel.Worksheet ws, int dataStartRow,
            long[] fix, long[] rin, long[] bp, long[] kei, long[] jin)
        {
            long[] uri = new long[12];
            long[] prof = new long[12];
            for (int i = 0; i < 12; i++)
            {
                uri[i] = fix[i] + rin[i] + bp[i];
                prof[i] = uri[i] - kei[i] - jin[i];
            }

            long[][] src = new[] { fix, rin, bp, uri, kei, jin, prof };
            string[] labels = { "固定売上", "臨時売上", "物品売上", "売上合計", "諸経費", "人件費", "損益" };

            // ✅ A〜R(1〜18列) までしか作らない
            var buf = new object[7, 18];

            for (int r = 0; r < 7; r++)
            {
                buf[r, 0] = labels[r]; // A
                long sum = 0;
                for (int i = 0; i < 12; i++)
                {
                    buf[r, 1 + i] = src[r][i]; // B..M
                    sum += src[r][i];
                }
                buf[r, 13] = sum;                                 // N
                buf[r, 14] = src[r][0] + src[r][1] + src[r][2];   // O
                buf[r, 15] = src[r][3] + src[r][4] + src[r][5];   // P
                buf[r, 16] = src[r][6] + src[r][7] + src[r][8];   // Q
                buf[r, 17] = src[r][9] + src[r][10] + src[r][11]; // R
            }

            // ✅ 書き込むのも A..R だけに限定
            var topLeft = ws.Cells[dataStartRow, 1];
            var bottomRight = ws.Cells[dataStartRow + 6, 18];
            ws.Range[topLeft, bottomRight].Value2 = buf;
        }


        void ExportExcelByTemplateAdvanced(List<CsvRow> filtered, string fileyymm, string nendo,
                                           string templatePath, string savePath)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("テンプレートが見つかりません", templatePath);

            var blocks = BuildBlocksForExport(filtered);

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application { Visible = false };
                WithExcelFast(app, delegate ()
                {
                    wb = app.Workbooks.Open(templatePath, ReadOnly: false);
                    ws = wb.Worksheets["Sheet1"] as Excel.Worksheet;

                    int yyyy = int.Parse(fileyymm.Substring(0, 4));
                    int mm = int.Parse(fileyymm.Substring(4, 2));
                    int kisuu = (mm < 4) ? (yyyy - 1972) : (yyyy - 1971);
                    ws.Range["D3"].Value2 = "◇◇◇　　第" + kisuu + "期（" + (mm < 4 ? (yyyy - 1) : yyyy) + "年度）現場総計数実績表　　◇◇◇";
                    ws.Range["R1"].Value2 = yyyy + "年" + mm.ToString("00") + "月度現在";

                    // ブロック構造を把握
                    int headerRow = FindFirstHeaderRow(ws);
                    int blockHeight = FindBlockHeight(ws, headerRow);
                    if (headerRow <= 0 || blockHeight <= 0) throw new Exception("テンプレートのブロック構造を認識できませんでした。");

                    // 必要数に増減
                    EnsureBlockCountFast(ws, headerRow, blockHeight, blocks.Count);

                    // 各ブロックを出力
                    for (int i = 0; i < blocks.Count; i++)
                    {
                        int top = headerRow + i * blockHeight;
                        FillBlock(ws, top, blocks[i]);
                    }

                    // --- 余った“次ブロック用ラベル行”をクリア（A:R 文字のみ消す）---
                    try
                    {
                        int trailingLabelRow = headerRow + blocks.Count * blockHeight - 1; // 次ブロックのラベル行位置
                        Excel.Range tail = ws.Range[ws.Cells[trailingLabelRow, 1], ws.Cells[trailingLabelRow, 18]];
                        tail.ClearContents();
                    }
                    catch { }

                    // --- Excel を開いたら「総合計」ラベル行を選択表示 ---
                    try
                    {
                        int sogoHeaderRow = headerRow + (blocks.Count - 1) * blockHeight; // 最後のヘッダ（年月）行
                        int sogoLabelRow = sogoHeaderRow - 1;                              // その1行上がラベル行
                        (ws.Cells[sogoLabelRow, 1] as Excel.Range).Select();
                        ws.Application.ActiveWindow.ScrollRow = Math.Max(1, sogoLabelRow - 5);
                    }
                    catch { }

                    // 保存（SaveCopyAs→Move で安定・高速）
                    string tmp = Path.Combine(Path.GetDirectoryName(savePath), "~tmp_" + Path.GetFileName(savePath));
                    wb.SaveCopyAs(tmp);
                    wb.Close(SaveChanges: false);
                    if (File.Exists(savePath)) File.Delete(savePath);
                    File.Move(tmp, savePath);
                });
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
                ws = null; wb = null; app = null;
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        // ============ 共通ユーティリティ ============
        void ResetArrays(params long[][] arrays)
        {
            for (int a = 0; a < arrays.Length; a++)
            {
                var arr = arrays[a];
                for (int i = 0; i < arr.Length; i++) arr[i] = 0;
            }
        }

        bool BmnSkSuppressTotal(string bmnskVal)
        {
            var engi = new HashSet<string>(new[] { "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "202A", "202B" });
            var beigun = new HashSet<string>(new[] { "2031", "2032" });
            var shitei = new HashSet<string>(new[] { "4000", "4010", "4020", "4030", "4040", "4050", "4060", "4070", "4090", "4110", "4120" });

            return engi.Contains(bmnskVal) || beigun.Contains(bmnskVal) || shitei.Contains(bmnskVal);
        }

        string GetChikuLabel(string tk)
        {
            if (tk == "2") return "◆◇◆　那　覇　合　計　◆◇◆";
            if (tk == "3") return "◆◇◆　八重山　合　計　◆◇◆";
            if (tk == "4") return "◆◇◆　北　部　合　計　◆◇◆";
            if (tk == "5") return "◆◇◆　広　域　合　計　◆◇◆";
            if (tk == "6") return "◆◇◆　宮古島　合　計　◆◇◆";
            if (tk == "7") return "◆◇◆　久米島　合　計　◆◇◆";
            return "地区合計";
        }

        void GetChikuShokushuName(string bumonCd, out string chiku, out string shokushu)
        {
            chiku = "";
            string head = LeftStr(bumonCd, 1);
            if (head == "2") chiku = "那　覇";
            else if (head == "3") chiku = "八重山";
            else if (head == "4") chiku = "北　部";
            else if (head == "5") chiku = "広　域";
            else if (head == "6") chiku = "宮古島";
            else if (head == "7") chiku = "久米島";

            string tail = bumonCd.Length >= 4 ? RightStr(bumonCd, 4) : "";
            shokushu = "？？？？？";
            if (tail == "1010") shokushu = (head == "5" ? "多面展開" : "現　業");
            else if (tail == "1020") shokushu = "技術企画";
            else if (tail == "1030") shokushu = "警　備";
            else if (tail == "1040") shokushu = "遠方監視";
            else if (tail == "1050" || tail == "1051") shokushu = "サービス";
            else if (tail == "1053") shokushu = "車　輌";
            else if (tail == "1054") shokushu = "フロント";
            else if (tail == "1060") shokushu = "客　室";
            else if (tail == "1070") shokushu = "総　括";
            else if (tail == "1080") shokushu = "マンション";
            else if (tail == "1090") shokushu = "食　堂";
            else if (tail == "2010") shokushu = "施設管理";
            else if (tail == "2021") shokushu = "エンジ１課";
            else if (tail == "2022") shokushu = "エンジ２課";
            else if (tail == "2023") shokushu = "エンジ３課";
            else if (tail == "2024") shokushu = "技術営業課";
            else if (tail == "2025") shokushu = "宮古１課";
            else if (tail == "2026") shokushu = "宮古２課";
            else if (tail == "2027") shokushu = "久米島１課";
            else if (tail == "2028") shokushu = "久米島２課";
            else if (tail == "2031" || tail == "202A") shokushu = "米軍施設";
            else if (tail == "2032" || tail == "202B") shokushu = "米軍ﾌﾟﾛｼﾞｪｸﾄ";
            else if (tail == "3000") shokushu = "行　雲";
            else if (tail == "4000") shokushu = "指定管理";
            else if (tail == "4010") shokushu = "指）現　業";
            else if (tail == "4020") shokushu = "指）技術企画";
            else if (tail == "4030") shokushu = "指）警　備";
            else if (tail == "4040") shokushu = "指）機械警備";
            else if (tail == "4050") shokushu = "北指）サービス";
            else if (tail == "4070") shokushu = "指）統　括";
            else if (tail == "4090") shokushu = "指）飲　食";
            else if (tail == "4110") shokushu = "指）施　設";
            else if (tail == "4120") shokushu = "指）エンジ";
            else if (tail == "4051") shokushu = "指）サービス";
            else if (tail == "4055") shokushu = "指）植　栽";
            else if (tail == "1055") shokushu = "植　栽";
        }

        int MonthToIndex(int yyyymm)
        {
            int m = yyyymm % 100;
            return (m <= 3) ? (m + 10) : (m - 2);
        }


    }
}
