using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class Saichin : Form
    {
        private DataTable dt = new DataTable();
        private DataTable wkdt = new DataTable();

        private SqlConnection Cn;
        private SqlDataAdapter daMain;   // 更新対象：s最賃改定
        private SqlDataAdapter daEmp;    // 表示専用：s社員基本情報
        private DataSet ds = new DataSet();

        private string result;
        private DataView dvMain;         // 表示用ビュー（RowFilterで絞り込み）

        private TextBox editingTextBox;                 // 改定額編集用テキストボックス参照
        private const int MIN_KAITEIGAKU = 1023;        // 1022以下NG → 1023以上OK

        private bool _fillingDown = false;  // フィルダウン中の再入防止フラグ

        public Saichin()
        {
            InitializeComponent();
            dgvyosan.AllowUserToAddRows = false;   // ★ 追加：新規行を出さない

            // フォーム最大化・見た目調整
            this.WindowState = FormWindowState.Maximized;
            dgvyosan.Font = new Font(dgvyosan.Font.Name, 10);
            dgvyosan.RowHeadersVisible = false;
            dgvyosan.ColumnHeadersHeight = 10;

            // ラベル初期化（デザイナで lblError を置いておく）
            lblError.Text = "";
            lblError.Visible = false;
            lblError.ForeColor = Color.Red;  // 必要なら色指定

            dgvyosan.ShowCellErrors = true;   // 既定trueだが明示しておく
            dgvyosan.ShowRowErrors = false;  // 行ヘッダーを使わないならfalseでOK（任意）


            dgvyosan.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            soshiki.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            soshikigenba.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            kobetsu.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            gekkyuu.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;


            // ★ 自動生成の揺れ対策：バインド完了時にも順序補正
            dgvyosan.DataBindingComplete += (s, e) => FixColumnOrder();

            dgvyosan.CellValidated += Dgvyosan_CellValidated;   // フィルダウン発火

            IniSet();

            Com.InHistory("26_最賃", "", "");

            // テキストボックス変更時も即時フィルタ
            this.textBox1.TextChanged += textBox1_TextChanged;

            if (Program.loginname == "喜屋武　大祐")
            { 
                tsuika.Visible = true;
            }
        }

        private void IniSet()
        {
            SetTiku();
            SetBumon();
            SetGenba();
            GetData();
            ApplyFilter();

            GetSyuukeiData();


        }

        // 予算表示（データ取得とバインド）
        private void GetData()
        {
            // 都度作り直し（Relation/式列の二重追加を防ぐ）
            ds = new DataSet();
            Cn = new SqlConnection(Com.SQLConstr);

            // === 更新対象：s最賃改定（列明示） ===
            var sqlMain = @"
SELECT
    a.年度,
    a.社員番号,
    a.改定額,
    a.備考,
    a.更新者,
    a.更新日時
FROM dbo.s最賃改定 AS a;";

            daMain = new SqlDataAdapter(sqlMain, Cn)
            {
                MissingSchemaAction = MissingSchemaAction.Add
            };

            // === 表示専用：s社員基本情報（検索用列も含めて列明示） ===
            var sqlEmp = @"
SELECT
    b.社員番号,
    b.氏名,
    b.組織名,
    b.現場名,
    b.年齢,
    b.給与支給区分名,
    b.役職名,
    b.勤務時間,
    b.週労働数,
    b.在籍年月,
    b.週労時間,
    b.基準支給額,
    /* 追加：フィルタ用の元列 */
    b.担当区分,
    b.担当事務,
    b.職種,
    b.現場CD,
    b.組織CD,
    現行額 = CASE 
               WHEN b.給与支給区分名 LIKE N'日給者%' THEN b.日給
               ELSE b.時給
             END,
    /* 追加：全文検索っぽく使う寄せ集めキー */
    (ISNULL(b.氏名,'') + ' ' + ISNULL(b.組織名,'') + ' ' + ISNULL(b.現場名,'') + ' ' + ISNULL(b.給与支給区分名,'')
        + ' ' + ISNULL(b.担当区分,'') + ' ' + ISNULL(b.担当事務,'') + ' ' + ISNULL(b.職種,'')
    ) AS reskey
FROM dbo.s社員基本情報 AS b;";

            daEmp = new SqlDataAdapter(sqlEmp, Cn)
            {
                MissingSchemaAction = MissingSchemaAction.Add
            };

            // === 取得 ===
            Cn.Open();
            daMain.Fill(ds, "s最賃改定");
            daEmp.Fill(ds, "s社員基本情報");
            Cn.Close();

            var tMain = ds.Tables["s最賃改定"];
            var tEmp = ds.Tables["s社員基本情報"];

            // 複合キー（年度＋社員番号）で更新識別、社員基本は社員番号が一意
            if (tMain.PrimaryKey.Length == 0)
                tMain.PrimaryKey = new[] { tMain.Columns["年度"], tMain.Columns["社員番号"] };
            if (tEmp.PrimaryKey.Length == 0)
                tEmp.PrimaryKey = new[] { tEmp.Columns["社員番号"] };

            // Relation: 親(社員基本.社員番号) → 子(最賃改定.社員番号)
            const string REL = "rel_社員";
            if (!ds.Relations.Contains(REL))
                ds.Relations.Add(REL, tEmp.Columns["社員番号"], tMain.Columns["社員番号"], createConstraints: false);

            // === s最賃改定 に式列（表示専用）を追加 ===
            string[] empCols = new[] {
    "氏名","組織名","現場名","年齢","給与支給区分名","役職名",
    "勤務時間","週労働数","在籍年月","週労時間","基準支給額","現行額", // ★ ここ
    // 追加：フィルタ用列
    "担当区分","担当事務","職種","現場CD","組織CD","reskey"
};

            foreach (var colName in empCols)
            {
                if (!tMain.Columns.Contains(colName))
                {
                    var col = new DataColumn(colName, tEmp.Columns[colName].DataType)
                    {
                        Expression = $"Parent({REL}).[{colName}]",
                        ReadOnly = true
                    };
                    tMain.Columns.Add(col);
                }
            }

            if (!tMain.Columns.Contains("差額"))
            {
                var diffCol = new DataColumn("差額", typeof(decimal))
                {
                    Expression = "Convert(改定額, System.Decimal) - Convert(現行額, System.Decimal)", // ★ 現行額ベース
                    ReadOnly = true
                };
                tMain.Columns.Add(diffCol);
            }


            // === 更新系コマンド（手書き）：CommandBuilder不要 ===
            // UPDATE
            var cmdUpdate = new SqlCommand(@"
UPDATE dbo.s最賃改定
SET 改定額 = @改定額,
    備考   = @備考,
    更新者 = @更新者,
    更新日時 = @更新日時
WHERE 年度 = @年度_key AND 社員番号 = @社員番号_key;", Cn);

            cmdUpdate.Parameters.Add(new SqlParameter("@改定額", SqlDbType.Decimal) { SourceColumn = "改定額", SourceVersion = DataRowVersion.Current });
            cmdUpdate.Parameters.Add(new SqlParameter("@備考", SqlDbType.NVarChar) { SourceColumn = "備考", SourceVersion = DataRowVersion.Current, IsNullable = true });
            //cmdUpdate.Parameters.Add(new SqlParameter("@更新者", SqlDbType.NVarChar) { SourceColumn = "更新者", SourceVersion = DataRowVersion.Current });
            //cmdUpdate.Parameters.Add(new SqlParameter("@更新日時", SqlDbType.DateTime) { SourceColumn = "更新日時", SourceVersion = DataRowVersion.Current });

            cmdUpdate.Parameters.Add(new SqlParameter("@更新者", SqlDbType.NVarChar) { Value = Program.loginname }); // 固定値
            cmdUpdate.Parameters.Add(new SqlParameter("@更新日時", SqlDbType.DateTime) { Value = DateTime.Now }); // 固定値

            cmdUpdate.Parameters.Add(new SqlParameter("@年度_key", SqlDbType.NVarChar) { SourceColumn = "年度", SourceVersion = DataRowVersion.Original });
            cmdUpdate.Parameters.Add(new SqlParameter("@社員番号_key", SqlDbType.NVarChar) { SourceColumn = "社員番号", SourceVersion = DataRowVersion.Original });

            // INSERT
            var cmdInsert = new SqlCommand(@"
INSERT INTO dbo.s最賃改定 (年度, 社員番号, 改定額, 備考, 更新者, 更新日時)
VALUES (@年度, @社員番号, @改定額, @備考, @更新者, @更新日時);", Cn);

            cmdInsert.Parameters.Add(new SqlParameter("@年度", SqlDbType.NVarChar) { SourceColumn = "年度" });
            cmdInsert.Parameters.Add(new SqlParameter("@社員番号", SqlDbType.NVarChar) { SourceColumn = "社員番号" });
            cmdInsert.Parameters.Add(new SqlParameter("@改定額", SqlDbType.Decimal) { SourceColumn = "改定額" });
            cmdInsert.Parameters.Add(new SqlParameter("@備考", SqlDbType.NVarChar) { SourceColumn = "備考", IsNullable = true });
            cmdInsert.Parameters.Add(new SqlParameter("@更新者", SqlDbType.NVarChar) { SourceColumn = "更新者" });
            cmdInsert.Parameters.Add(new SqlParameter("@更新日時", SqlDbType.DateTime) { SourceColumn = "更新日時" });

            // DELETE（必要なら）
            var cmdDelete = new SqlCommand(@"
DELETE FROM dbo.s最賃改定
WHERE 年度 = @年度_key AND 社員番号 = @社員番号_key;", Cn);
            cmdDelete.Parameters.Add(new SqlParameter("@年度_key", SqlDbType.NVarChar) { SourceColumn = "年度", SourceVersion = DataRowVersion.Original });
            cmdDelete.Parameters.Add(new SqlParameter("@社員番号_key", SqlDbType.NVarChar) { SourceColumn = "社員番号", SourceVersion = DataRowVersion.Original });

            daMain.UpdateCommand = cmdUpdate;
            daMain.InsertCommand = cmdInsert;
            daMain.DeleteCommand = cmdDelete;

            // === バインド ===
            dvMain = new DataView(tMain);
            //dvMain.Sort = "[組織CD] ASC, [現場CD] ASC";
            dvMain.Sort = "[組織CD] ASC, [現場CD] ASC, [給与支給区分名] ASC, [社員番号] ASC";
            dgvyosan.AutoGenerateColumns = true;   // 自動生成してから並べ替える
            dgvyosan.DataSource = dvMain;

            //// 表示専用（社員基本情報）をグリッドでもReadOnlyに
            //foreach (var colName in empCols)
            //    if (dgvyosan.Columns[colName] != null) dgvyosan.Columns[colName].ReadOnly = true;


            // まず全列ReadOnly
            foreach (DataGridViewColumn c in dgvyosan.Columns)
                c.ReadOnly = true;

            // 例外：この2列だけ編集可
            void makeEditable(string name)
            {
                if (dgvyosan.Columns[name] != null)
                {
                    var col = dgvyosan.Columns[name];
                    col.ReadOnly = false;
                    // 見た目（薄い黄色系）
                    col.DefaultCellStyle.BackColor = Color.FromArgb(255, 252, 214);
                    col.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 235, 150);
                    col.HeaderCell.Style.BackColor = Color.FromArgb(255, 242, 153);
                    col.HeaderCell.Style.Font = new Font(dgvyosan.Font, FontStyle.Bold);
                }
            }

            dgvyosan.EnableHeadersVisualStyles = false; // ヘッダー色を反映させるため
            makeEditable("改定額");
            makeEditable("備考");

            // 追加の安全策：セル編集開始でガード（何かの拍子にReadOnlyが外れた場合に備える）
            dgvyosan.CellBeginEdit += (s, e) =>
            {
                var colName = dgvyosan.Columns[e.ColumnIndex].Name;
                if (colName != "改定額" && colName != "備考")
                {
                    e.Cancel = true;
                    System.Media.SystemSounds.Beep.Play();
                }
            };

            //多重購読防止のうえでイベント購読
            dgvyosan.EditingControlShowing -= Dgvyosan_EditingControlShowing;
            dgvyosan.EditingControlShowing += Dgvyosan_EditingControlShowing;

            dgvyosan.CellValidating -= Dgvyosan_CellValidating;
            dgvyosan.CellValidating += Dgvyosan_CellValidating;

            dgvyosan.CellParsing -= Dgvyosan_CellParsing;
            dgvyosan.CellParsing += Dgvyosan_CellParsing;

            dgvyosan.CellEndEdit -= Dgvyosan_CellEndEdit;
            dgvyosan.CellEndEdit += Dgvyosan_CellEndEdit;

            dgvyosan.DataError -= Dgvyosan_DataError;
            dgvyosan.DataError += Dgvyosan_DataError;

            // 非表示
            // 追加：フィルタ専用列は非表示にする
            string[] hideCols = { "年度", "基準支給額", "担当区分", "担当事務", "職種", "現場CD", "組織CD", "reskey" };
            foreach (var name in hideCols)
            {
                if (dgvyosan.Columns[name] != null)
                    dgvyosan.Columns[name].Visible = false;
            }

            // ヘッダー中央寄せなど
            for (int i = 0; i < dgvyosan.Columns.Count; i++)
                dgvyosan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvyosan.RowTemplate.Height = 20;

            // 表示幅
            if (dgvyosan.Columns["社員番号"] != null) dgvyosan.Columns["社員番号"].Width = 90;
            if (dgvyosan.Columns["氏名"] != null) dgvyosan.Columns["氏名"].Width = 100;
            if (dgvyosan.Columns["組織名"] != null) dgvyosan.Columns["組織名"].Width = 100;
            if (dgvyosan.Columns["現場名"] != null) dgvyosan.Columns["現場名"].Width = 150;
            if (dgvyosan.Columns["年齢"] != null) dgvyosan.Columns["年齢"].Width = 30;
            if (dgvyosan.Columns["給与支給区分名"] != null) dgvyosan.Columns["給与支給区分名"].Width = 100;
            if (dgvyosan.Columns["役職名"] != null) dgvyosan.Columns["役職名"].Width = 50;
            if (dgvyosan.Columns["勤務時間"] != null) dgvyosan.Columns["勤務時間"].Width = 30;
            if (dgvyosan.Columns["週労働数"] != null) dgvyosan.Columns["週労働数"].Width = 70;
            if (dgvyosan.Columns["在籍年月"] != null) dgvyosan.Columns["在籍年月"].Width = 100;
            if (dgvyosan.Columns["週労時間"] != null) dgvyosan.Columns["週労時間"].Width = 50;
            if (dgvyosan.Columns["基準支給額"] != null) dgvyosan.Columns["基準支給額"].Width = 50;
            if (dgvyosan.Columns["現行額"] != null) dgvyosan.Columns["現行額"].Width = 50;
            if (dgvyosan.Columns["改定額"] != null) dgvyosan.Columns["改定額"].Width = 70;
            if (dgvyosan.Columns["備考"] != null) dgvyosan.Columns["備考"].Width = 150;
            if (dgvyosan.Columns["更新者"] != null) dgvyosan.Columns["更新者"].Width = 100;
            if (dgvyosan.Columns["更新日時"] != null) dgvyosan.Columns["更新日時"].Width = 120;

            if (dgvyosan.Columns["差額"] != null)
            {
                dgvyosan.Columns["差額"].Width = 50;
                dgvyosan.Columns["差額"].DefaultCellStyle.Format = "#,0";
                dgvyosan.Columns["差額"].ReadOnly = true; // 念押し（表示専用）
            }

            // 三桁区切り表示
            if (dgvyosan.Columns["勤務時間"] != null) dgvyosan.Columns["勤務時間"].DefaultCellStyle.Format = "#,0";
            if (dgvyosan.Columns["週労時間"] != null) dgvyosan.Columns["週労時間"].DefaultCellStyle.Format = "#,0";
            if (dgvyosan.Columns["現行額"] != null) dgvyosan.Columns["現行額"].DefaultCellStyle.Format = "#,0";
            if (dgvyosan.Columns["改定額"] != null) dgvyosan.Columns["改定額"].DefaultCellStyle.Format = "#,0";

            // 数値系の列を右寄せ
            string[] rightCols = { "年齢", "勤務時間", "週労時間", "現行額", "改定額", "差額" };
            foreach (var name in rightCols)
            {
                if (dgvyosan.Columns[name] != null)
                {
                    dgvyosan.Columns[name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            if (dgvyosan.Columns["備考"] != null)
                dgvyosan.Columns["備考"].DefaultCellStyle.NullValue = "";


            // ヘッダー表示名変更
            if (dgvyosan.Columns["給与支給区分名"] != null)
                dgvyosan.Columns["給与支給区分名"].HeaderText = "区分";

            // ★ 最後に一発だけ順序補正
            FixColumnOrder();

            
        }



        // 表示順にしたい列（この順で左→右に並ぶ）
        private static readonly string[] DisplayOrder = new[] {
  "社員番号","氏名","組織名","現場名","給与支給区分名","役職名",
  "年齢","在籍年月","勤務時間","週労働数","週労時間","現行額", // ★
  "改定額","差額",                                             // ★
  "備考","更新者","更新日時"
};


        // DataSourceバインド後に呼ぶと、確実に並び替える
        private void FixColumnOrder()
        {
            if (dgvyosan.Columns.Count == 0) return;

            dgvyosan.SuspendLayout();

            // 1) 指定列を順番通りに前へ詰める
            int idx = 0;
            foreach (var name in DisplayOrder)
            {
                if (dgvyosan.Columns.Contains(name))
                    dgvyosan.Columns[name].DisplayIndex = idx++;
            }

            // 2) 指定外の列は、現在の相対順を維持したまま後方へ
            foreach (var col in dgvyosan.Columns
                                        .Cast<DataGridViewColumn>()
                                        .Where(c => !DisplayOrder.Contains(c.Name))
                                        .OrderBy(c => c.DisplayIndex))
            {
                col.DisplayIndex = idx++;
            }

            dgvyosan.ResumeLayout();
        }



        // 変更内容をDB反映
        private void button1_Click(object sender, EventArgs e)
        {
            // 1) いま編集中のセルの検証＆確定を強制
            this.Validate(); // フォーム全体の Validating を走らせる
            dgvyosan.EndEdit();
            dgvyosan.CommitEdit(DataGridViewDataErrorContexts.Commit); // 念のため
            var cm = (CurrencyManager)BindingContext[dgvyosan.DataSource];
            cm.EndCurrentEdit();

            // （任意）新規行を使わないなら混乱防止
            dgvyosan.AllowUserToAddRows = false;

            try
            {
                Cn.Open();
                daMain.Update(ds.Tables["s最賃改定"]);
                Cn.Close();
                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                try { if (Cn.State == ConnectionState.Open) Cn.Close(); } catch { }
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
            }

            GetSyuukeiData();
        }

        private void GetSyuukeiData()
        {
            //TODO 年月、最賃、埋め込み

            DataTable dt = new DataTable();
            dt = Com.GetDB("select * from dbo.s最賃改定集計組織別('2025','10') order by 組織CD");
            soshiki.DataSource = dt;
            FormatSoshiki(soshiki);

            DataTable dt2 = new DataTable();
            dt2 = Com.GetDB("select * from dbo.s最賃改定集計組織別現場別('2025','10') order by 組織CD, 現場CD");
            soshikigenba.DataSource = dt2;
            FormatSoshikiGenba(soshikigenba);

            DataTable dt3 = new DataTable();
            dt3 = Com.GetDB("select * from dbo.s最賃改定個別一覧('2025','10') order by 組織CD, 現場CD");
            kobetsu.DataSource = dt3;

            DataTable dt4 = new DataTable();
            dt4 = Com.GetDB("select * from dbo.s最賃改定月給一覧('1023') order by 組織CD, 現場CD");
            gekkyuu.DataSource = dt4;
        }

        private void FormatSoshiki(DataGridView dgv)
        {
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToResizeColumns = true;
            dgv.AllowUserToOrderColumns = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgv.Font = new Font(dgv.Font.Name, 10);

            // === 幅まとめ ===
            if (dgv.Columns["組織CD"] != null) dgv.Columns["組織CD"].Width = 80;
            if (dgv.Columns["組織名"] != null) dgv.Columns["組織名"].Width = 150;
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].Width = 100;
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].Width = 90;
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].Width = 90;
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].Width = 90;

            // === 寄せまとめ ===
            if (dgv.Columns["組織CD"] != null) dgv.Columns["組織CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dgv.Columns["組織名"] != null) dgv.Columns["組織名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // === 数値フォーマットまとめ ===
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].DefaultCellStyle.Format = "#,0";
        }

        private void FormatSoshikiGenba(DataGridView dgv)
        {
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToResizeColumns = true;
            dgv.AllowUserToOrderColumns = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgv.Font = new Font(dgv.Font.Name, 10);

            // === 幅まとめ ===
            if (dgv.Columns["組織CD"] != null) dgv.Columns["組織CD"].Width = 80;
            if (dgv.Columns["組織名"] != null) dgv.Columns["組織名"].Width = 150;
            if (dgv.Columns["現場CD"] != null) dgv.Columns["現場CD"].Width = 80;
            if (dgv.Columns["現場名"] != null) dgv.Columns["現場名"].Width = 150;
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].Width = 100;
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].Width = 90;
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].Width = 90;
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].Width = 90;

            // === 寄せまとめ ===
            if (dgv.Columns["組織CD"] != null) dgv.Columns["組織CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dgv.Columns["組織名"] != null) dgv.Columns["組織名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            if (dgv.Columns["現場CD"] != null) dgv.Columns["現場CD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dgv.Columns["現場名"] != null) dgv.Columns["現場名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // === 数値フォーマットまとめ ===
            if (dgv.Columns["増額"] != null) dgv.Columns["増額"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["パート人数"] != null) dgv.Columns["パート人数"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["日給者人数"] != null) dgv.Columns["日給者人数"].DefaultCellStyle.Format = "#,0";
            if (dgv.Columns["月給者人数"] != null) dgv.Columns["月給者人数"].DefaultCellStyle.Format = "#,0";
        }


        // === DataView.RowFilter を適用 ===
        private void ApplyFilter()
        {
            // 既存の検索条件生成
            ResultStr();

            // 先頭が「 and」の場合、削除する
            if (!string.IsNullOrEmpty(result) && result.StartsWith(" and"))
                result = result.Remove(0, 4);

            // RowFilter の IsNull 置換
            string filter = (result ?? string.Empty).Replace("isnull(", "IsNull(");

            // ★ 追加：パート/アルバイト/日給者 固定条件
            const string MUST =
                "(IsNull([給与支給区分名], '') LIKE 'パート%' " +
                " OR IsNull([給与支給区分名], '') LIKE 'アルバイト%' " +
                " OR IsNull([給与支給区分名], '') LIKE '日給者%')";

            // 空なら固定条件だけ、あるなら AND で結合
            string finalFilter = string.IsNullOrWhiteSpace(filter)
                ? MUST
                : $"{MUST} AND ({filter})";

            try
            {
                if (dvMain != null) dvMain.RowFilter = finalFilter;
            }
            catch (EvaluateException ex)
            {
                MessageBox.Show("検索条件に誤りがあります。入力内容をご確認ください。\n\n" + ex.Message);
                if (dvMain != null) dvMain.RowFilter = MUST; // 固定条件は維持
            }
        }

        private static string Esc(string s) => s?.Replace("'", "''") ?? string.Empty;

        // 検索条件文字列を生成（RowFilter式）
        private void ResultStr()
        {
            string res = textBox1.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            result = "";

            // キーワード（かな/カナ/半角含む）
            if (ar.Length > 0 && !string.IsNullOrEmpty(ar[0]))
            {
                foreach (string s0 in ar)
                {
                    if (string.IsNullOrWhiteSpace(s0)) continue;

                    var s = Esc(s0);
                    var s1 = Esc(Com.isOneByteChar(s0));
                    var sKat = Esc(Strings.StrConv(s0, VbStrConv.Katakana));
                    var sKat1 = Esc(Com.isOneByteChar(Strings.StrConv(s0, VbStrConv.Katakana)));
                    var sHir = Esc(Strings.StrConv(s0, VbStrConv.Hiragana));
                    var sHir1 = Esc(Strings.StrConv(Com.isOneByteChar(s0), VbStrConv.Hiragana));

                    result +=
                        " and (reskey LIKE '%" + s + "%'" +
                        " or reskey LIKE '%" + s1 + "%'" +
                        " or reskey LIKE '%" + sKat + "%'" +
                        " or reskey LIKE '%" + sKat1 + "%'" +
                        " or reskey LIKE '%" + sHir + "%'" +
                        " or reskey LIKE '%" + sHir1 + "%')";
                }
            }

            // 部門（担当区分）…未チェックは除外
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i))
                {
                    result += " and 担当区分 <> '" + Esc(checkedListBox1.Items[i].ToString()) + "'";
                }
            }

            // 職種 …未チェックは除外
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i))
                {
                    result += " and 担当事務 <> '" + Esc(checkedListBox2.Items[i].ToString()) + "'";
                }
            }

            // 現場 …多数/少数選択でIN/NOT IN 相当を生成
            int itemcount = checkedListBox3.Items.Count; // 項目数合計
            int ckcount = checkedListBox3.CheckedItems.Count; // チェック項目数

            if (ckcount == 0)
            {
                // 何も選択されていないときはヒットさせない
                result += " and 現場CD = '99999' ";
                return;
            }

            if (itemcount / 2 > ckcount)
            {
                // 選択が少ない → 選択したものだけ許可（OR）
                string sql3 = "";
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (checkedListBox3.GetItemChecked(i))
                    {
                        var s = checkedListBox3.Items[i].ToString();
                        if (s.Length < 5)
                            sql3 += " or IsNull(現場CD,'') = '' ";
                        else
                            sql3 += " or IsNull(現場CD,'') = '" + Esc(s.Substring(0, 5)) + "'";
                    }
                }
                if (sql3.Length > 0) result += " and ( " + sql3.Substring(4) + " ) ";
            }
            else
            {
                // 未選択が少ない → 未選択だけ除外（ANDで積む）
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    if (!checkedListBox3.GetItemChecked(i))
                    {
                        var s = checkedListBox3.Items[i].ToString();
                        if (s.Length < 5)
                            result += " and IsNull(現場CD,'') <> '' ";
                        else
                            result += " and IsNull(現場CD,'') <> '" + Esc(s.Substring(0, 5)) + "'";
                    }
                }
            }
        }

        // ====== マスタ側（絞り込みUI） ======
        private void SetTiku()
        {
            checkedListBox1.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当区分 from dbo.s社員基本情報 where 在籍区分 <> '9' and 給与支給区分名 in (N'パート', N'アルバイト', N'日給者') order by 担当区分";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox1.Items.Add(row["担当区分"]);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void SetBumon()
        {
            checkedListBox2.Items.Clear();

            DataTable dt = new DataTable();
            string sql = "select distinct 担当事務 from dbo.s社員基本情報 where 在籍区分 <> '9' and 給与支給区分名 in (N'パート', N'アルバイト', N'日給者')";


            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            sql += " order by 担当事務 ";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox2.Items.Add(row["担当事務"]);
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void SetGenba()
        {
            //リストボックスの項目(Item)を消去
            checkedListBox3.Items.Clear();

            DataTable dt = new DataTable();

            string sql = "select distinct 現場CD, 現場名 from dbo.s社員基本情報 where 在籍区分 <> '9' and 給与支給区分名 in (N'パート', N'アルバイト', N'日給者')";


            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!checkedListBox1.GetItemChecked(i)) sql += " and 担当区分 <> '" + checkedListBox1.Items[i].ToString() + "' ";
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (!checkedListBox2.GetItemChecked(i)) sql += " and 担当事務 <> '" + checkedListBox2.Items[i].ToString() + "' ";
            }

            sql += " order by 現場CD,現場名 ";

            dt = Com.GetDB(sql);

            foreach (DataRow row in dt.Rows)
            {
                checkedListBox3.Items.Add(row["現場CD"].ToString() + ' ' + row["現場名"].ToString());
            }

            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, true);
            }
        }

        // ====== ラベルの全選択/全解除 ======
        private void label3_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }

            SetBumon();
            SetGenba();
            GetData();
            ApplyFilter();
        }

        private void label23_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
            }

            SetGenba();
            GetData();
            ApplyFilter();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            if (checkedListBox3.CheckedItems.Count == 0)
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    checkedListBox3.SetItemChecked(i, false);
                }
            }
            GetData();
            ApplyFilter();
        }

        // ====== チェック変更で即時反映 ======
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetBumon();
            SetGenba();
            GetData();
            ApplyFilter();
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetGenba();
            GetData();
            ApplyFilter();
        }

        private void checkedListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetData();
            ApplyFilter();
        }

        // テキスト変更でもフィルタ
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (dvMain != null) ApplyFilter();
        }

        // ★★★ ここから追加：入力正規化＆検証ロジック

        // 全角→半角／数字以外除去／前後空白除去
        private static string NormalizeDigits(string s)
        {
            var narrow = Strings.StrConv(s ?? "", VbStrConv.Narrow).Trim();
            var onlyDigits = new string(narrow.Where(char.IsDigit).ToArray()); // カンマ等も除去
            return onlyDigits;
        }

        // 編集コントロール出現時：改定額列のみ数値入力に制限
        private void Dgvyosan_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ClearErrorMessage();

            if (dgvyosan.CurrentCell == null) return;

            var colName = dgvyosan.Columns[dgvyosan.CurrentCell.ColumnIndex].Name;

            // 以前のハンドラをデタッチ
            if (editingTextBox != null)
            {
                editingTextBox.KeyPress -= EditingTextBox_KeyPress_NumericOnly;
                editingTextBox.TextChanged -= EditingTextBox_TextChanged_Normalize;
                editingTextBox = null;
            }

            if (colName == "改定額")
            {
                editingTextBox = e.Control as TextBox;
                if (editingTextBox != null)
                {
                    // IME無効（全角入力になりにくくする）
                    editingTextBox.ImeMode = ImeMode.Disable;

                    // 既存表示（#,0 などのカンマ付き）をその場で正規化
                    var normalized = NormalizeDigits(editingTextBox.Text);
                    if (editingTextBox.Text != normalized)
                    {
                        editingTextBox.Text = normalized;
                        editingTextBox.SelectionStart = editingTextBox.Text.Length;
                    }

                    // 数字/制御キーのみ許可
                    editingTextBox.KeyPress += EditingTextBox_KeyPress_NumericOnly;
                    // 貼り付け等で混入した全角・非数字を即時除去
                    editingTextBox.TextChanged += EditingTextBox_TextChanged_Normalize;
                }
            }
        }

        // 数字と制御キー（Backspace等）以外はブロック
        private void EditingTextBox_KeyPress_NumericOnly(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;
            if (char.IsDigit(e.KeyChar)) return; // 全角数字もtrueだが、TextChangedで半角化＆非数字除去する
            e.Handled = true;
        }

        // 入力中に全角→半角＆非数字除去（貼り付け対策）
        private void EditingTextBox_TextChanged_Normalize(object sender, EventArgs e)
        {
            var tb = (TextBox)sender;
            var selStart = tb.SelectionStart;
            var selLen = tb.SelectionLength;

            var normalized = NormalizeDigits(tb.Text);
            if (tb.Text != normalized)
            {
                tb.Text = normalized;
                // キャレット位置復元（可能な範囲で）
                tb.SelectionStart = Math.Min(selStart, tb.Text.Length);
                tb.SelectionLength = selLen;
            }
        }

        // 値の確定前検証：空/非数/しきい値(1023以下)を拒否
        private void Dgvyosan_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dgvyosan.Columns[e.ColumnIndex].Name != "改定額") return;

            string input = NormalizeDigits(e.FormattedValue?.ToString() ?? "");

            if (string.IsNullOrEmpty(input))
            {
                ShowErrorMessage("改定額を入力してください。");
                e.Cancel = true;
                return;
            }

            if (!decimal.TryParse(input, out var val))
            {
                ShowErrorMessage("数値として認識できません。");
                e.Cancel = true;
                return;
            }

            // ここ重要：1023 は OK。1023 未満はエラー
            if (val < MIN_KAITEIGAKU) // MIN_KAITEIGAKU = 1023
            {
                ShowErrorMessage("最賃割れです。");
                e.Cancel = true;
                return;
            }

            // 正常時：表示も正規化して、エラーラベルを消す
            var tb = dgvyosan.EditingControl as TextBox;
            if (tb != null) tb.Text = input;
            ClearErrorMessage();
        }

        // 実値への反映時に、正規化した数値を decimal としてセット
        private void Dgvyosan_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (dgvyosan.Columns[e.ColumnIndex].Name != "改定額") return;

            var s = NormalizeDigits(e.Value?.ToString() ?? "");
            if (decimal.TryParse(s, out var val))
            {
                e.Value = val;          // DataTableのdecimal列にそのまま入る
                e.ParsingApplied = true;
            }
        }
        private void ShowErrorMessage(string message)
        {
            lblError.Text = message;
            lblError.Visible = true;
        }

        private void ClearErrorMessage()
        {
            lblError.Text = "";
            lblError.Visible = false;
        }

        private void Dgvyosan_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // 確定できた＝妥当な入力なので消す
            ClearErrorMessage();
        }

        private void Dgvyosan_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // 想定外の型変換エラーなどもラベルに出す
            ShowErrorMessage("入力エラー: " + (e.Exception?.Message ?? ""));
            e.ThrowException = false;
        }

        // ★ フィルダウン：チェックボックスON時、改定額の入力確定で下方向に同値コピー
        private void Dgvyosan_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            // 改定額列以外は対象外／再入防止
            if (_fillingDown) return;
            if (dgvyosan.Columns[e.ColumnIndex].Name != "改定額") return;

            // フォーム上に CheckBox「chkFillDown」（例：Text=「下へコピー」）を置いておいてください
            var cb = this.Controls.Find("chkFillDown", true).FirstOrDefault() as CheckBox;
            if (cb == null || !cb.Checked) return;

            // 入力確定した値を取得（検証は既存の CellValidating が実施済み）
            var valObj = dgvyosan.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (valObj == null || valObj == DBNull.Value) return;
            if (!decimal.TryParse(valObj.ToString(), out var val)) return;

            try
            {
                _fillingDown = true;

                // 画面上の順序で「下方向」へコピー（フィルタ/ソート後の見えている順序）
                for (int r = e.RowIndex + 1; r < dgvyosan.Rows.Count; r++)
                {
                    var gridRow = dgvyosan.Rows[r];
                    if (gridRow.IsNewRow) continue;

                    var drv = gridRow.DataBoundItem as DataRowView;
                    if (drv == null) continue;
                    var dr = drv.Row;

                    // 既存値を取得
                    decimal? old = dr.IsNull("改定額") ? (decimal?)null : (decimal)dr["改定額"];

                    if (old != val)
                    {
                        // 値が違うなら普通に代入 → Modified / Added は自然に付く
                        dr["改定額"] = val;
                    }
                    else
                    {
                        // ★ 同額でも更新者・更新日時を更新したい場合のみ、
                        // RowState が Unchanged のときだけ SetModified を呼ぶ
                        if (dr.RowState == DataRowState.Unchanged)
                        {
                            dr.SetModified();   // ← ここなら例外にならない
                        }
                        // RowState が Added / Modified のときは何もしない
                    }
                }

            }
            finally
            {
                _fillingDown = false;
            }
        }

        private void splitContainer1_Paint(object sender, PaintEventArgs e)
        {
            // Splitter の座標を取得
            int splitterWidth = splitContainer1.SplitterWidth;
            int splitterDistance = splitContainer1.SplitterDistance;

            // 描画する矩形を決定（Orientationで分岐）
            Rectangle rect;
            if (splitContainer1.Orientation == Orientation.Vertical)
            {
                // 縦分割（左右に仕切り）
                rect = new Rectangle(splitContainer1.SplitterDistance, 0, splitterWidth, splitContainer1.Height);
            }
            else
            {
                // 横分割（上下に仕切り）
                rect = new Rectangle(0, splitContainer1.SplitterDistance, splitContainer1.Width, splitterWidth);
            }

            // 好きな色で塗る
            using (Brush brush = new SolidBrush(Color.LightSkyBlue))
            {
                e.Graphics.FillRectangle(brush, rect);
            }
        }

        private void tsuika_Click(object sender, EventArgs e)
        {
            //TODO 年度埋め込み
            if (Program.loginname == "喜屋武　大祐")
            {
                DataTable dt = new DataTable();
                dt = Com.GetDB("exec usp_最賃改定_対象者一括登録 '2025'");
                MessageBox.Show(dt.Rows[0][0].ToString());
            }
        }
    }
}
