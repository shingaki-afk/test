using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Net.Mail;
using System.Net;

namespace ODIS.ODIS
{
    public partial class Ringi : Form
    {
        private DataTable dt = new DataTable();

        public Ringi()
        {
            InitializeComponent();

            //フォームを最大化
            this.WindowState = FormWindowState.Maximized;

            GetData();

            Com.InHistory("61_過去稟議(～2020年8月末)", "", "");

            //フィルター設定
            c1FlexGrid1.AllowFiltering = true;

            //自動グリップボード機能を有効にする
            c1FlexGrid1.AutoClipboard = true;

            // グリッドのAllowMergingプロパティを設定
            c1FlexGrid1.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free;

            
        }

        private void GetData()
        {

            string res = textBox3.Text.Trim().Replace("　", " ");
            string[] ar = res.Split(' ');

            string result = "";
            if (ar[0] != "")
            {
                foreach (string s in ar)
                {
                    result += " and (reskey like '%" + s + "%' or reskey like '%" + Com.isOneByteChar(s) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Katakana) + "%' or reskey like '%" + Com.isOneByteChar(Strings.StrConv(s, VbStrConv.Katakana)) + "%' or reskey like '%" + Strings.StrConv(s, VbStrConv.Hiragana) + "%' or reskey like '%" + Strings.StrConv(Com.isOneByteChar(s), VbStrConv.Hiragana) + "%')";
                }
            }

            //先頭が「and」の場合、「where」にする
            if (result.StartsWith(" and"))
            {
                result = result.Remove(0, 4);
                result = " where " + result;
            }

            dt = Com.GetDB("select 稟議番号, 地区名, 組織名, 氏名, 登録日, 申請額, 区分, 目的, 結果, コメント, 社員番号, reskey from dbo.r稟議データ検索" + result + " order by 登録日 desc, 稟議番号 desc");

            //列幅自動調整　バインド前でないといけない
            //c1FlexGrid1.AutoResize = true;

            c1FlexGrid1.DataSource = dt;

            c1FlexGrid1.Cols["社員番号"].Visible = false;
            c1FlexGrid1.Cols["reskey"].Visible = false;
            //マージ設定
            c1FlexGrid1.Cols["地区名"].AllowMerging = true;
            c1FlexGrid1.Cols["組織名"].AllowMerging = true;
            c1FlexGrid1.Cols["氏名"].AllowMerging = true;
            c1FlexGrid1.Cols["登録日"].AllowMerging = true;

            //折り返し
            //c1FlexGrid1.Cols["目的"].Style.WordWrap = true;


            //列幅
            c1FlexGrid1.Cols["稟議番号"].Width = 75;
            c1FlexGrid1.Cols["地区名"].Width = 75;
            c1FlexGrid1.Cols["組織名"].Width = 120;
            c1FlexGrid1.Cols["氏名"].Width = 100;
            c1FlexGrid1.Cols["登録日"].Width = 100;
            c1FlexGrid1.Cols["申請額"].Width = 80;
            c1FlexGrid1.Cols["区分"].Width = 100;
            c1FlexGrid1.Cols["目的"].Width = 300;
            c1FlexGrid1.Cols["結果"].Width = 50;
            c1FlexGrid1.Cols["コメント"].Width = 150;

            //テキスト配置
            //c1FlexGrid1.Cols["稟議番号"].TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            //c1FlexGrid1.Cols["地区名"].TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            //c1FlexGrid1.Cols["組織名"].TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            //c1FlexGrid1.Cols["氏名"].TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;

            //3ケタカンマ
            c1FlexGrid1.Cols["申請額"].Format = "#,0";


            // 選択モードを行単位での選択のみにする
            c1FlexGrid1.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;

            c1FlexGrid1.Cols["結果"].ComboList = " |承認|否認|保留|取消|";

            //c1FlexGrid1.Cols["結果"].Caption = "Multi";
            //C1.Win.C1FlexGrid.Column col;

            //// カスタムスタイル"Combo1"を作成
            //c1FlexGrid1.Styles.Add("Combo1");
            //// ComboListプロパティを設定
            //c1FlexGrid1.Styles["Combo1"].ComboList = "AAA|BBB|CCC";
            //// 3行1列目のセルを選択
            //col = c1FlexGrid1.Cols["目的"];
            //// カスタムスタイルを割り当てます
            //col.Style = col.Styles["Combo1"];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GetData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //係長以上かチェック
            if (Convert.ToInt16(Program.yakusyokucd) > 135)
            {
                MessageBox.Show("稟議起案者は係長以上となります。"); return;
            }

            //dt.Clear();
            //GetData();

            //Form2に送るテキスト
            string[] sendText = new string[] { "I", "" };

            //Form2から送られてきたテキストを受け取る。
            string[] receiveText = RingiMini.ShowMiniForm(sendText);　//Form2を開く

            //画面バッテン押した場合の対処
            if (receiveText == null) return;

            //Form2から受け取ったテキストをForm1で表示させてあげる。
            //ReceiveTextBox.Text = receiveText;
            //MessageBox.Show(receiveText);
            //_number.Text = receiveText[0];
            //_name.Text = receiveText[1];

            string en = Environment.NewLine;
            //MessageBox.Show("登録しました。" + en + receiveText[0] + en + receiveText[1] + en + receiveText[2] + en + receiveText[3]);


            GetData();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            //押されたキーがエンターキーかどうかの条件分岐
            //if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            //{
            //    GetData();
            //}
        }

        private void c1FlexGrid1_ComboCloseUp(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            MessageBox.Show("これでは変更できないす。。該当選択して右クリック⇒「修正」を選択してください。");
        }

        private void Syuusei_Click(object sender, EventArgs e)
        {
            //承認後の修正はできない
            if (c1FlexGrid1[c1FlexGrid1.RowSel, 9].ToString() != "")
            {
                MessageBox.Show("決裁後の修正はできません。");
                return;
            }

            //修正は起案者しかできない
            if (c1FlexGrid1[c1FlexGrid1.RowSel, 4].ToString() != Program.loginname)
            {
                if (Program.loginID == "21151800")
                {
                    MessageBox.Show("修正は起案者しかできません。が、システム管理者例外");
                } else
                {
                    MessageBox.Show("修正は起案者しかできません。");
                    return;
                }
            }

            string[] str = new string[] { "S", c1FlexGrid1[c1FlexGrid1.RowSel, 1].ToString() };
            string[] receiveText = RingiMini.ShowMiniForm(str);
            //画面バッテン押した場合の対処
            if (receiveText == null) return;
            GetData();
        }

        private void Kessai_Click(object sender, EventArgs e)
        {
            //総務部長(佐久川常務金城部長)しかできません
            if (Program.loginID != "18923000" && Program.loginID != "19300309")
            {
                if (Program.loginID == "21151800")
                {
                    MessageBox.Show("決裁は総務部長しかできません。が、システム管理者例外");
                }
                else
                {
                    MessageBox.Show("決裁は総務部長しかできません。");
                    return;
                }

            }

            string[] str = new string[] { "K", c1FlexGrid1[c1FlexGrid1.RowSel, 1].ToString() };

            //string[] receiveText = RingiMini.ShowMiniForm(c1FlexGrid1[c1FlexGrid1.RowSel, 1].ToString());
            string[] receiveText = RingiMini.ShowMiniForm(str);

            //画面バッテン押した場合の対処
            if (receiveText == null) return;

            //0社員番号
            //1稟議番号
            //2氏名
            //3目的
            //4申請額
            //5コメント
            //6結果
            //7起案日

            //社員番号からメールアドレスを取得
            DataTable dt = new DataTable();
            dt = Com.GetDB("select メール from dbo.社員基本情報 where 社員番号 = '" + receiveText[0] + "';");
            string ml = dt.Rows[0][0].ToString();
            //TODO
            SendMail(ml, receiveText[1], receiveText[2], receiveText[3], receiveText[4], receiveText[5], receiveText[6], receiveText[7]);

            GetData();
         }

        private void SendMail(string mail, string no, string name, string moku, string gaku, string com, string kekka, string kiday)
        {
            string nl = Environment.NewLine;

            string sub = "";
            //sub = "【" + kekka + "】【" + no + "】【" + moku + "】";
            sub = "【" + kekka + "】【" + moku + "】";
            string body = "";
            body = "【決裁結果】" + kekka + nl;
            body += "【コメント】" + com + nl + nl;
            body += "【起案日】" + kiday + nl;
            body += "【起案者】" + name + nl;
            body += "【稟議番号】" + no + nl;
            body += "【内容】" + moku + nl;
            body += "【金額】" + gaku + nl;


            using (var client = new SmtpClient("smtp.gmail.com", 587))
            {
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential("admin@oki-daiken.co.jp", "admin0110");
                client.EnableSsl = true;

                // MailMessageクラスを使って送信するメールを作成する
                var message = new MailMessage();

                // 差出人アドレス
                message.From = new MailAddress("webmaster@oki-daiken.co.jp", "稟議決裁通知");

                if (Program.loginml == "kyan@oki-daiken.co.jp")
                {
                    message.To.Add(new MailAddress("kyan@oki-daiken.co.jp"));
                }
                else
                {
                    message.To.Add(new MailAddress(mail));
                    message.Bcc.Add(new MailAddress("sakugawa@oki-daiken.co.jp"));
                    //message.Bcc.Add(new MailAddress("iskw@oki-daiken.co.jp"));
                    message.Bcc.Add(new MailAddress("kyan@oki-daiken.co.jp"));
                }

                //添付ファイル
                //message.Attachments.Add(new Attachment(@"\\daikensrv03\23_労働安全衛生\特性要因図（例2）.xls"));

                // メールの優先度を設定する
                //message.Priority = MailPriority.High;

                // メールの送信日時(Dateヘッダ)を設定する
                //message.Headers["Date"] = (new DateTime(2001, 2, 3, 4, 56, 7)).ToString("r");

                // メールの件名
                message.Subject = sub;

                // メールの本文
                message.Body = body;

                try
                {
                    // 作成したメールを送信する
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("例外が発生しました" + ex);
                }
                //}
            }
            MessageBox.Show("決裁メール送信しました。" + nl + body);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            GetData();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }
    }
}
