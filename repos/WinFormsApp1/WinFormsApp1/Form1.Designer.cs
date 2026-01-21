namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnRed = new Button();
            textBox1 = new TextBox();
            lbl1 = new Label();
            comboBox1 = new ComboBox();
            checkBox1 = new CheckBox();
            btnBlue = new Button();
            pictureBox1 = new PictureBox();
            lblTest = new Label();
            btnTest = new Button();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // btnRed
            // 
            btnRed.BackColor = Color.Red;
            btnRed.Location = new Point(628, 347);
            btnRed.Name = "btnRed";
            btnRed.Size = new Size(119, 57);
            btnRed.TabIndex = 0;
            btnRed.Text = "赤";
            btnRed.UseVisualStyleBackColor = false;
            btnRed.Click += btnRed_Click;
            btnRed.MouseLeave += btnRed_MouseLeave;
            btnRed.MouseHover += btnRed_MouseHover;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(403, 141);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(100, 23);
            textBox1.TabIndex = 1;
            // 
            // lbl1
            // 
            lbl1.AutoSize = true;
            lbl1.BackColor = Color.DarkBlue;
            lbl1.BorderStyle = BorderStyle.Fixed3D;
            lbl1.Font = new Font("ＭＳ 明朝", 24F, FontStyle.Bold, GraphicsUnit.Point, 128);
            lbl1.ForeColor = Color.Coral;
            lbl1.Location = new Point(136, 103);
            lbl1.Name = "lbl1";
            lbl1.Size = new Size(119, 35);
            lbl1.TabIndex = 2;
            lbl1.Text = "label1";
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(134, 141);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(121, 23);
            comboBox1.TabIndex = 3;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(288, 145);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(83, 19);
            checkBox1.TabIndex = 4;
            checkBox1.Text = "checkBox1";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // btnBlue
            // 
            btnBlue.BackColor = Color.Blue;
            btnBlue.Location = new Point(494, 347);
            btnBlue.Name = "btnBlue";
            btnBlue.Size = new Size(105, 57);
            btnBlue.TabIndex = 5;
            btnBlue.Text = "青";
            btnBlue.UseVisualStyleBackColor = false;
            btnBlue.Click += btnBlue_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Location = new Point(117, 189);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(354, 215);
            pictureBox1.TabIndex = 6;
            pictureBox1.TabStop = false;
            // 
            // lblTest
            // 
            lblTest.AutoSize = true;
            lblTest.Location = new Point(422, 103);
            lblTest.Name = "lblTest";
            lblTest.Size = new Size(38, 15);
            lblTest.TabIndex = 7;
            lblTest.Text = "label2";
            // 
            // btnTest
            // 
            btnTest.Location = new Point(540, 228);
            btnTest.Name = "btnTest";
            btnTest.Size = new Size(142, 74);
            btnTest.TabIndex = 8;
            btnTest.Text = "button1";
            btnTest.UseVisualStyleBackColor = true;
            btnTest.Click += btnTest_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnTest);
            Controls.Add(lblTest);
            Controls.Add(pictureBox1);
            Controls.Add(btnBlue);
            Controls.Add(checkBox1);
            Controls.Add(comboBox1);
            Controls.Add(lbl1);
            Controls.Add(textBox1);
            Controls.Add(btnRed);
            Name = "Form1";
            Text = "Form1";
            FormClosing += Form1_FormClosing;
            FormClosed += Form1_FormClosed;
            Load += Form1_Load;
            Shown += Form1_Shown;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnRed;
        private TextBox textBox1;
        private Label lbl1;
        private ComboBox comboBox1;
        private CheckBox checkBox1;
        private Button btnBlue;
        private PictureBox pictureBox1;
        private Label lblTest;
        private Button btnTest;
    }
}
