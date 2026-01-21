using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class GinouZ : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private DataSet ds = new DataSet();

        // ▼ s社員基本情報を保持
        private DataTable dtEmployee;

        // ▼ CellFormattingで使う高速参照用インデックス
        private Dictionary<string, DataRow> empIndex;

        // ▼ 表示追加したい列（s社員基本情報 側の列名）
        //   ※ DataGridViewの表示名も同一にしています（必要なら別名にしてもOK）
        private readonly string[] empColumns = new[]
        {
            "氏名","カナ名","地区CD","地区名","組織CD","組織名","現場CD","現場名",
            "役職CD","役職名","契約社員","契約社員名","休日区分","休日区分名","友の会区分",
            "給与支給区分","給与支給区分名","性別区分","生年月日","入社年月日","退職年月日",
            "休暇付与区分","週労働数","国籍","勤務時間","時給","日給",
            "パスポート番号","パスポート有効期限","在籍年月"
        };

        public GinouZ()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            try
            {
                Cn = new SqlConnection(Com.SQLConstr);

                // g技能実習生情報 は従来通り単独SELECT（更新のためにJOINはしない）
                var sql = @"SELECT * FROM dbo.g技能実習生情報;";

                da = new SqlDataAdapter(sql, Cn)
                {
                    MissingSchemaAction = MissingSchemaAction.AddWithKey
                };
                new SqlCommandBuilder(da);

                // ▼ s社員基本情報：必要な列をまとめて取得
                var empSql = @"
SELECT
    社員番号,
    氏名,
    カナ名,
    地区CD,
    地区名,
    組織CD,
    組織名,
    現場CD,
    現場名,
    役職CD,
    役職名,
    契約社員,
    契約社員名,
    休日区分,
    休日区分名,
    友の会区分,
    給与支給区分,
    給与支給区分名,
    性別区分,
    生年月日,
    入社年月日,
    退職年月日,
    休暇付与区分,
    週労働数,
    国籍,
    勤務時間,
    時給,
    日給,
    パスポート番号,
    パスポート有効期限,
    在籍年月
FROM dbo.s社員基本情報;
";
                dtEmployee = new DataTable();
                using (var tempCn = new SqlConnection(Com.SQLConstr))
                using (var tempDa = new SqlDataAdapter(empSql, tempCn))
                {
                    tempDa.Fill(dtEmployee);
                }

                // ▼ g技能実習生情報 取得
                Cn.Open();
                da.Fill(ds, "g技能実習生情報");
            }
            catch (Exception ex)
            {
                MessageBox.Show("初期化中にエラーが発生しました。\n" + ex.Message);
            }
            finally
            {
                if (Cn.State == ConnectionState.Open) Cn.Close();
            }

            // ▼ DataGridViewバインド
            dataGridView1.DataSource = ds.Tables["g技能実習生情報"];

            // ▼ s社員基本情報 からの表示列を追加（存在しない場合だけ）
            AddEmployeeColumnsToGrid();

            // ▼ 参照用インデックス（社員番号 → DataRow）
            BuildEmployeeIndex();

            // ▼ 表示差し込み
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;


            Com.InHistory("27_技能実習生", "", "");
        }

        private void AddEmployeeColumnsToGrid()
        {
            // まず社員番号列がDataGridView側に存在することを確認
            if (!dataGridView1.Columns.Contains("社員番号"))
            {
                // g技能実習生情報に「社員番号」が無い場合は処理できないため通知
                MessageBox.Show("g技能実習生情報に「社員番号」列が見つかりません。社員番号を含めてください。");
                return;
            }

            foreach (var colName in empColumns)
            {
                if (!dataGridView1.Columns.Contains(colName))
                {
                    var dgvCol = new DataGridViewTextBoxColumn
                    {
                        Name = colName,
                        HeaderText = colName,
                        ReadOnly = true
                    };
                    dataGridView1.Columns.Add(dgvCol);
                }
            }

            // ★ ここを追加：更新可能な列（g技能実習生情報由来）を白背景に統一
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (!empColumns.Contains(col.Name))
                {
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.Honeydew;
                }
            }
        }

        private void BuildEmployeeIndex()
        {
            // 社員番号の型が文字/数値いずれでも動くよう ToString() をキーに統一
            empIndex = dtEmployee
                .AsEnumerable()
                .Where(r => r["社員番号"] != DBNull.Value)
                .GroupBy(r => Convert.ToString(r["社員番号"]))
                .ToDictionary(g => g.Key, g => g.First());
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // 社員番号列が無い場合は何もしない
            if (dataGridView1.Columns["社員番号"] == null) return;

            // 行範囲チェック
            if (e.RowIndex < 0 || e.RowIndex >= dataGridView1.Rows.Count) return;

            // 現在のセルが、追加表示対象の列でなければ戻る
            var colName = dataGridView1.Columns[e.ColumnIndex].Name;
            if (!empColumns.Contains(colName)) return;

            var employeeIdCell = dataGridView1.Rows[e.RowIndex].Cells["社員番号"];
            if (employeeIdCell?.Value == null || employeeIdCell.Value is DBNull) return;

            var employeeId = Convert.ToString(employeeIdCell.Value);
            if (string.IsNullOrEmpty(employeeId)) return;

            if (empIndex != null && empIndex.TryGetValue(employeeId, out var row))
            {
                // 値取得
                var raw = row.Table.Columns.Contains(colName) ? row[colName] : null;
                if (raw == null || raw is DBNull) return;

                // 型に応じて体裁を整える
                if (raw is DateTime dt)
                {
                    e.Value = dt.ToShortDateString();      // 例: 2025/09/03
                    e.FormattingApplied = true;
                }
                else if (raw is TimeSpan ts)
                {
                    e.Value = ts.ToString(@"hh\:mm");      // 例: 08:30
                    e.FormattingApplied = true;
                }
                else if (raw is decimal dec)
                {
                    e.Value = dec.ToString("#,0.##");      // 例: 1,234.5
                    e.FormattingApplied = true;
                }
                else if (raw is double dbl)
                {
                    e.Value = dbl.ToString("#,0.##");
                    e.FormattingApplied = true;
                }
                else if (raw is int i)
                {
                    e.Value = i.ToString("#,0");
                    e.FormattingApplied = true;
                }
                else
                {
                    e.Value = Convert.ToString(raw);
                    e.FormattingApplied = true;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.EndEdit();

                Cn.Open();
                da.Update(ds.Tables["g技能実習生情報"]);

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新中にエラーが発生しました。\n" + ex.Message);
            }
            finally
            {
                if (Cn.State == ConnectionState.Open) Cn.Close();
            }
        }
    }
}
