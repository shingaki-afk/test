using System;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace ODIS
{
    class SendMail
    {
        public SendMail(string subject, string body, string pass)
        {
            //if (Program.loginml == "kyan@oki-daiken.co.jp")
            //{
            //    MailMessage msg = new MailMessage();
            //    msg.From = new MailAddress("webmaster@oki-daiken.co.jp", "沖縄ダイケン 事故速報メール");
            //    msg.To.Add(new MailAddress("kyan@oki-daiken.co.jp"));
            //    msg.Subject = "本日はどうも";
            //    msg.Body = "すみませんでした";

            //    SmtpClient sc = new SmtpClient();

            //    sc.Host = "smtp.gmail.com";
            //    sc.Port = 587;
            //    sc.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    //sc.Credentials = new System.Net.NetworkCredential("admin@oki-daiken.co.jp", "admin0110");
            //    sc.Credentials = new System.Net.NetworkCredential("admin@oki-daiken.co.jp", "admin0110");
            //    sc.EnableSsl = true;

            //    sc.Send(msg);
            //    msg.Dispose();
            //}
            //else
            //{

                using (var client = new SmtpClient("smtp.gmail.com", 587))
                {
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Credentials = new NetworkCredential("admin@oki-daiken.co.jp", "admin0110");
                    client.EnableSsl = true;
                    // MailMessageクラスを使って送信するメールを作成する
                    var message = new MailMessage();

                    // 差出人アドレス
                    message.From = new MailAddress("webmaster@oki-daiken.co.jp", "沖縄ダイケン 事故速報");

                if (Program.loginml == "kyan@oki-daiken.co.jp" || Program.loginml == "saya@oki-daiken.co.jp")
                {
                    message.To.Add(new MailAddress("kyan@oki-daiken.co.jp"));
                }
                else
                {
                    //役員
                    message.To.Add(new MailAddress("yakuin@oki-daiken.co.jp"));

                        //部門長
                        message.To.Add(new MailAddress("bumonchou@oki-daiken.co.jp"));

                        //KM
                        message.To.Add(new MailAddress("groupkm@oki-daiken.co.jp"));
                    
                        //管理職　那覇安全衛生メンバー
                        message.To.Add(new MailAddress("kawamitsu-h@oki-daiken.co.jp"));
                        message.To.Add(new MailAddress("touma-t@oki-daiken.co.jp"));
  
                        //八重山
                        //message.To.Add(new MailAddress("oshiro-h@oki-daiken.co.jp"));
                        message.To.Add(new MailAddress("k-arakaki@oki-daiken.co.jp"));
                        message.To.Add(new MailAddress("masakatsu@oki-daiken.co.jp"));

                        message.CC.Add(new MailAddress(Program.loginml)); //登録者のメールアドレス
                        message.CC.Add(new MailAddress("kyan@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("saya@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("hanako@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("miyu-ki@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("ruriko@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("yuzuki@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("ayaka@oki-daiken.co.jp"));
                        message.CC.Add(new MailAddress("sayoko@oki-daiken.co.jp"));
                }

                //添付ファイル
                //message.Attachments.Add(new Attachment(@"\\daikensrv03\23_労働安全衛生\特性要因図（例2）.xls"));

                // メールの優先度を設定する
                //message.Priority = MailPriority.High;

                // メールの送信日時(Dateヘッダ)を設定する
                //message.Headers["Date"] = (new DateTime(2001, 2, 3, 4, 56, 7)).ToString("r");

                // メールの件名
                message.Subject = subject;

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
            }
        }
    }
}