namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Shownが発生しました。");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.BackgroundImage = Image.FromFile(@"C:\Users\22503263\Desktop\01.png");
        }
        private void btnRed_Click(object sender, EventArgs e)
        {
            pictureBox1.BackColor = Color.Red;
            textBox1.Clear();
            textBox1.AppendText("風ニモマケズ");
        }

        private void btnBlue_Click(object sender, EventArgs e)
        {
            pictureBox1.BackColor = Color.Blue;
            this.BackgroundImage = Image.FromFile(@"C:\Windows\Web\Wallpaper\Windows\img0.jpg");
            textBox1.Focus();
            textBox1.SelectAll();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            lblTest.BackColor = Color.MediumBlue;
            lblTest.BorderStyle = BorderStyle.Fixed3D;
            lblTest.Font = new Font("MS 明朝", 24, FontStyle.Bold);
            lblTest.Text = "プロパティは簡単";
        }

        private void btnRed_MouseHover(object sender, EventArgs e)
        {
            btnRed.BackColor = Color.Yellow;
        }

        private void btnRed_MouseLeave(object sender, EventArgs e)
        {
            btnRed.BackColor = Color.Red;

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("フォームが閉じますよ。", "FormClosingです。", MessageBoxButtons.OK, MessageBoxIcon.Information);
            MessageBox.Show("本当に閉じますよ！", "FormClosingです。", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            MessageBox.Show("フォームが閉じました。", "FormClosedです。", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

    }
}
