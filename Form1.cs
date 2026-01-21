using C1.C1Excel;
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

namespace KanriKeisuSyukei
{
    public partial class KanriKeisuSyukei : Form
    {
        public KanriKeisuSyukei()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                btnExport.Enabled = false;
                lblStatus.Text = "処理中...";


                //マウスカーソルを砂時計にする
                Cursor.Current = Cursors.WaitCursor;
                btnExport.Enabled = false;

                string fileName = @"\\daikensrv03\21_全体共通\40_総務発信_管理\ODIS\doc\計数\現場計数.xlsx";

                //手順1：新しいワークブックを作成します。
                C1XLBook c1XLBook1 = new C1XLBook();

                c1XLBook1.Load(fileName);

                string localPass = @"C:\ODIS\KEISUU\";
                string exlName = localPass + "計数" + DateTime.Now.ToString("yyyy年MM月dd日_HH時mm分ss秒_");
                string fullPath = Path.Combine(localPass, exlName);

                //フォルダがなければ作成する
                if (!System.IO.File.Exists(localPass))
                {
                    System.IO.Directory.CreateDirectory(localPass);
                }

                // 手順3：ファイルを保存します。
                c1XLBook1.Save(exlName + ".xlsx");

                //マウスカーソルをデフォルトにする
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
                btnExport.Enabled = true;

                //excel出力
                System.Diagnostics.Process.Start(exlName + ".xlsx");

                Com.InHistory("Excel計数を開いた。。", "", "");

            MessageBox.Show($"Excelの出力が完了しました。\n保存先: {fullPath}", "完了");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー");
            }
            finally
            {
                btnExport.Enabled = true;
                lblStatus.Text = "待機中";
            }
        }
    }
}