using ODIS.ODIS;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class AccountSet : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter daMain;   // 更新対象：アカウント管理
        private SqlDataAdapter daEmp;    // 表示用：s社員基本情報（社員番号は一意）
        private SqlCommandBuilder cb;
        private DataSet ds = new DataSet();

        public AccountSet()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            Cn = new SqlConnection(Com.SQLConstr);

            //--- 更新対象は単一テーブルのみ（*は可能なら必要列に絞る）---
            var sqlMain = @"
SELECT
    a.*   -- 可能なら: a.ID, a.ログイン名, a.部署コード, ... と明示列に
FROM dbo.アカウント管理 AS a;";

            daMain = new SqlDataAdapter(sqlMain, Cn)
            {
                // 主キー情報を取る（UPDATE/DELETE 生成に必要）
                MissingSchemaAction = MissingSchemaAction.AddWithKey
            };

            cb = new SqlCommandBuilder(daMain)
            {
                QuotePrefix = "[",
                QuoteSuffix = "]"
            };

            //--- 表示専用の退職年月日（社員番号は一意なので集約不要）---
            var sqlEmp = @"
SELECT
    b.社員番号,
    b.退職年月日
FROM dbo.s社員基本情報 AS b;";

            daEmp = new SqlDataAdapter(sqlEmp, Cn)
            {
                MissingSchemaAction = MissingSchemaAction.AddWithKey
            };

            //--- 取得 ---
            Cn.Open();
            daMain.Fill(ds, "アカウント管理");
            daEmp.Fill(ds, "社員");
            Cn.Close();

            var acc = ds.Tables["アカウント管理"];
            var emp = ds.Tables["社員"];

            // 社員（親）側の主キーを社員番号に
            if (emp.PrimaryKey == null || emp.PrimaryKey.Length == 0)
            {
                emp.PrimaryKey = new[] { emp.Columns["社員番号"] };
            }

            // Relation: 社員(社員番号, 親) → アカウント管理(ID, 子)
            if (ds.Relations["rel_社員"] == null)
            {
                ds.Relations.Add("rel_社員",
                    emp.Columns["社員番号"],
                    acc.Columns["ID"],
                    createConstraints: false);
            }

            // アカウント管理に式列として退職年月日を追加（自動ReadOnly＝表示専用）
            if (!acc.Columns.Contains("退職年月日"))
            {
                var col = new DataColumn("退職年月日", typeof(DateTime))
                {
                    Expression = "Parent(rel_社員).退職年月日",
                    ReadOnly = true
                };
                acc.Columns.Add(col);
            }

            dataGridView1.DataSource = acc;

            // 見た目上も編集不可を明示
            if (dataGridView1.Columns["退職年月日"] != null)
            {
                dataGridView1.Columns["退職年月日"].ReadOnly = true;
                // 任意の装飾:
                // dataGridView1.Columns["退職年月日"].DefaultCellStyle.BackColor = Color.Gainsboro;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // 退職年月日は式列なので更新対象外。基表のみ差分更新。
                Cn.Open();
                daMain.Update(ds.Tables["アカウント管理"]);
                Cn.Close();

                MessageBox.Show("更新しました。");
            }
            catch (Exception ex)
            {
                try { if (Cn.State == ConnectionState.Open) Cn.Close(); } catch { }
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
            }
        }
    }
}
