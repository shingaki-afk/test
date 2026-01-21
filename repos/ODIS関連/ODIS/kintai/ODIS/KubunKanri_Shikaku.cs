using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class KubunKanri_Shikaku : Form
    {
        private SqlConnection Cn;
        private SqlDataAdapter da;
        private SqlCommandBuilder cb;
        private DataTable dt = new DataTable();

        public KubunKanri_Shikaku()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            //グリッドビューのコピー
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            GetData();

            Com.InHistory("93_区分管理_資格登録情報", "", "");
        }

        private void GetData()
        {
            dataGridView1.DataSource = "";
            dt.Clear();

            Cn = new SqlConnection(Com.SQLConstr);
            Cn.Open();
            string sql = "select 会社コード, 情報キー, 管理コード, 適用開始日, 適用終了日, ソート順, 摘要, 数値５ as 金額, メモ１ as ID,  数値４ as 連番, (select count(*) from QUATRO.dbo.SJMTSHIKAK b where b.会社コード = 'E0' and a.管理コード = b.資格コード) as 数, フラグ１ as 表示フラグ, フラグ２ as 有効期限有無, フラグ３ as 支援有無,ユーザ任意フィールド１ as エンジ他,  ユーザ任意フィールド２ as 施設,  ユーザ任意フィールド３ as 警備, ユーザ任意フィールド４ as 現業,  ユーザ任意フィールド５ as 客室,  ユーザ任意フィールド６ as サービス,  ユーザ任意フィールド７ as 食堂,  ユーザ任意フィールド８ as 飲食,  ユーザ任意フィールド９ as 植栽,  ユーザ任意フィールド１０ as 車両,  ユーザ任意フィールド１１ as フロント,  ユーザ任意フィールド１２ as 管理事務・シス管,  ユーザ任意フィールド１３ as 営業, メモ２ as 備考 from QUATRO.dbo.QCMTCODED a where 情報キー = 'SJMT095' order by メモ１";
            //string sql = "select 会社コード, 情報キー, 管理コード, 適用開始日, 適用終了日, ソート順, 摘要, 数値５ as 金額, メモ１ as ID,  数値４ as 連番, (select count(*) from QUATRO.dbo.SJMTSHIKAK b where b.会社コード = 'E0' and a.管理コード = b.資格コード) as 数, フラグ１ as 表示フラグ, フラグ２ as 期限フラグ, ユーザ任意フィールド１ as エンジ,  フラグ３ as 施設,  フラグ４ as 警備, フラグ５ as 現業,  フラグ６ as 客室,  フラグ７ as サービス,  フラグ８ as 食堂,  フラグ９ as 飲食,  フラグ１０ as 植栽,  メモ２ as 車両,  メモ３ as フロント,  メモ４ as 事務,  メモ５ as 営業 from QUATRO.dbo.QCMTCODED a where 情報キー = 'SJMT095' order by メモ１"; //and メモ１ is not null
            //string sql = "select 会社コード, 情報キー, 管理コード, 摘要, メモ１ as ID, メモ２ as 備考, 数値４ as 連番, 数値５ as 金額, (select count(*) from QUATRO.dbo.SJMTSHIKAK b where a.管理コード = b.資格コード) as 数, フラグ１ as 表示フラグ, フラグ２ as 有効期限有無,  フラグ３ as 支援有無,  ユーザ任意フィールド１ as エンジ, ユーザ任意フィールド２ as 施設, ユーザ任意フィールド３ as 警備, ユーザ任意フィールド４ as 現業,  ユーザ任意フィールド５ as 客室,  ユーザ任意フィールド６ as サービス,  ユーザ任意フィールド７ as 食堂,  ユーザ任意フィールド８ as 飲食,  ユーザ任意フィールド９ as 植栽,  ユーザ任意フィールド１０ as 車両,  ユーザ任意フィールド１１ as フロント,  ユーザ任意フィールド１２ as 事務,  ユーザ任意フィールド１３ as 営業 from QUATRO.dbo.QCMTCODED a where 情報キー = 'SJMT095' order by メモ１";
            da = new SqlDataAdapter(sql, Cn);
            cb = new SqlCommandBuilder(da);
            da.Fill(dt);

            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //データ更新
                da.Update(dt);

                //データ更新終了をDataTableに伝える
                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー。管理者までご連絡ください。" + ex.ToString());
                throw;
            }

            GetData();
        }
    }
}
