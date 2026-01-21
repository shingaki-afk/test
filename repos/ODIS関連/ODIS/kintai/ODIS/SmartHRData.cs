using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic; // StrConv を使うために追加


namespace ODIS.ODIS
{
    public partial class SmartHRData : Form
    {
        public SmartHRData()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            // グリッドビュー設定
            // 1. 列名をコピーしない
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            // 2. 行ヘッダの空白列を非表示にする
            dataGridView1.RowHeadersVisible = false;


            //dataGridView1.DataSource = Com.GetDB("select * from dbo.SmartHRデータ移行 where 社員番号 between '22000001' and '23000000' order by 社員番号"); 
            dataGridView1.DataSource = Com.GetDB("select * from dbo.SmartHRデータ移行");

            DataTable dt = (DataTable)dataGridView1.DataSource;

            foreach (DataRow row in dt.Rows)
            {
                // --- 氏名（漢字） ---
                string fullName = row["緊急連絡先の姓"].ToString(); // ビューでは氏名1を丸ごと返す
                string[] nameParts = fullName.Split(new char[] { '　', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (nameParts.Length >= 2)
                {
                    row["緊急連絡先の姓"] = nameParts[0];
                    row["緊急連絡先の名"] = nameParts[1];
                }
                else
                {
                    row["緊急連絡先の姓"] = fullName;
                    row["緊急連絡先の名"] = "";
                }

                // --- 氏名（カナ） ---
                string fullKana = row["緊急連絡先の姓（ヨミガナ）"].ToString();

                // 半角カナ → 全角カナに変換
                fullKana = Strings.StrConv(fullKana, VbStrConv.Wide, 0x411); // 0x411 = 日本語ロケール

                string[] kanaParts = fullKana.Split(new char[] { '　', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (kanaParts.Length >= 2)
                {
                    row["緊急連絡先の姓（ヨミガナ）"] = kanaParts[0];
                    row["緊急連絡先の名（ヨミガナ）"] = kanaParts[1];
                }
                else
                {
                    row["緊急連絡先の姓（ヨミガナ）"] = fullKana;
                    row["緊急連絡先の名（ヨミガナ）"] = "";
                }

                string account = row["給与振込口座 口座番号"].ToString().Trim();

                if (!string.IsNullOrEmpty(account) && account.All(char.IsDigit))
                {
                    // 文字列として 7 桁ゼロ埋め
                    row["給与振込口座 口座番号"] = account.PadLeft(7, '0');
                }
                else
                {
                    row["給与振込口座 口座番号"] = account; // 数字以外が混ざっている場合はそのまま
                }

                string raw = row["雇用保険の被保険者番号"].ToString().Trim();

                // 数字だけを抽出
                string digits = new string(raw.Where(char.IsDigit).ToArray());

                if (digits.Length == 11) // 桁数が正しい場合
                {
                    // 1234-567890-9 の形式に整形
                    string formatted = digits.Substring(0, 4) + "-" +
                                       digits.Substring(4, 6) + "-" +
                                       digits.Substring(10, 1);

                    row["雇用保険の被保険者番号"] = formatted;
                }
                else
                {
                    // 桁数が合わない場合はそのまま
                    row["雇用保険の被保険者番号"] = raw;
                }

            }

            dataGridView1.DataSource = dt;

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "CSVファイル (*.csv)|*.csv",
                FileName = "SmartHR.csv"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ExportCsv(dataGridView1, sfd.FileName);
                MessageBox.Show("CSVファイルを出力しました。");
            }
        }


        // CSV保存メソッド
        private void ExportCsv(DataGridView dgv, string filePath)
        {
            var sb = new StringBuilder();

            // 1. 列名を書き出し
            var columnNames = dgv.Columns
                                 .Cast<DataGridViewColumn>()
                                 .Select(c => c.HeaderText);
            sb.AppendLine(string.Join(",", columnNames));

            // 2. 各行を書き出し
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (!row.IsNewRow)
                {
                    var cells = row.Cells.Cast<DataGridViewCell>().Select(cell =>
                    {
                        string val = cell.Value?.ToString() ?? "";

                        // CSVでダブルクォートがある場合はエスケープ
                        if (val.Contains(",") || val.Contains("\"") || val.Contains("\n"))
                        {
                            val = "\"" + val.Replace("\"", "\"\"") + "\"";
                        }

                        return val;
                    });

                    sb.AppendLine(string.Join(",", cells));
                }
            }

            // UTF-8 BOM付きで保存（Excelで文字化けしないように）
            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(true));
        }
    }
}
