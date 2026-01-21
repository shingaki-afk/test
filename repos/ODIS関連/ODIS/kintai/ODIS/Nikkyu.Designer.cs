namespace ODIS.ODIS
{
    partial class Nikkyu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.comboBoxGaku = new System.Windows.Forms.ComboBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxKeiken = new System.Windows.Forms.ComboBox();
            this.nikkyuu = new System.Windows.Forms.TextBox();
            this.comboBoxSyoku = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.honyuu = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.syokumukyuu = new System.Windows.Forms.TextBox();
            this.nennreikyuu = new System.Windows.Forms.TextBox();
            this.gakurekikyuu = new System.Windows.Forms.TextBox();
            this.tanka = new System.Windows.Forms.TextBox();
            this.kihongoukei = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.nenrei = new System.Windows.Forms.TextBox();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.comboBoxkasan = new System.Windows.Forms.ComboBox();
            this.tuukinhouhou = new System.Windows.Forms.ComboBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.kyori = new System.Windows.Forms.NumericUpDown();
            this.label17 = new System.Windows.Forms.Label();
            this.kinmu = new System.Windows.Forms.NumericUpDown();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.hi = new System.Windows.Forms.TextBox();
            this.ka = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.keikenkyuu = new System.Windows.Forms.TextBox();
            this.comboBoxritou = new System.Windows.Forms.ComboBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.sikyuugoukei = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.kyori)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kinmu)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBoxGaku
            // 
            this.comboBoxGaku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxGaku.FormattingEnabled = true;
            this.comboBoxGaku.Location = new System.Drawing.Point(152, 128);
            this.comboBoxGaku.Name = "comboBoxGaku";
            this.comboBoxGaku.Size = new System.Drawing.Size(121, 20);
            this.comboBoxGaku.TabIndex = 0;
            this.comboBoxGaku.SelectedIndexChanged += new System.EventHandler(this.comboBoxGaku_SelectedIndexChanged);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(152, 72);
            this.dateTimePicker1.MinDate = new System.DateTime(1945, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(121, 19);
            this.dateTimePicker1.TabIndex = 1;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(83, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "生年月日";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(83, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "最終学歴";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(59, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "社外経験年数";
            // 
            // comboBoxKeiken
            // 
            this.comboBoxKeiken.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxKeiken.FormattingEnabled = true;
            this.comboBoxKeiken.Location = new System.Drawing.Point(152, 157);
            this.comboBoxKeiken.Name = "comboBoxKeiken";
            this.comboBoxKeiken.Size = new System.Drawing.Size(121, 20);
            this.comboBoxKeiken.TabIndex = 0;
            this.comboBoxKeiken.SelectedIndexChanged += new System.EventHandler(this.comboBoxKeiken_SelectedIndexChanged);
            // 
            // nikkyuu
            // 
            this.nikkyuu.Font = new System.Drawing.Font("MS UI Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.nikkyuu.Location = new System.Drawing.Point(482, 221);
            this.nikkyuu.Name = "nikkyuu";
            this.nikkyuu.ReadOnly = true;
            this.nikkyuu.Size = new System.Drawing.Size(100, 26);
            this.nikkyuu.TabIndex = 3;
            this.nikkyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // comboBoxSyoku
            // 
            this.comboBoxSyoku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxSyoku.FormattingEnabled = true;
            this.comboBoxSyoku.Location = new System.Drawing.Point(152, 15);
            this.comboBoxSyoku.Name = "comboBoxSyoku";
            this.comboBoxSyoku.Size = new System.Drawing.Size(121, 20);
            this.comboBoxSyoku.TabIndex = 0;
            this.comboBoxSyoku.SelectedIndexChanged += new System.EventHandler(this.comboBoxSyoku_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(447, 14);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 12);
            this.label6.TabIndex = 2;
            this.label6.Text = "本給";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(107, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 2;
            this.label7.Text = "職種";
            // 
            // honyuu
            // 
            this.honyuu.Location = new System.Drawing.Point(482, 11);
            this.honyuu.Name = "honyuu";
            this.honyuu.ReadOnly = true;
            this.honyuu.Size = new System.Drawing.Size(100, 19);
            this.honyuu.TabIndex = 3;
            this.honyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.Location = new System.Drawing.Point(382, 156);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(88, 16);
            this.label8.TabIndex = 2;
            this.label8.Text = "基本給合計";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(435, 42);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "職務給";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(407, 70);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(69, 12);
            this.label5.TabIndex = 2;
            this.label5.Text = "技能給_年齢";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(407, 98);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(69, 12);
            this.label9.TabIndex = 2;
            this.label9.Text = "技能給_学歴";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(383, 128);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(93, 12);
            this.label10.TabIndex = 2;
            this.label10.Text = "技能給_社外経験";
            // 
            // syokumukyuu
            // 
            this.syokumukyuu.Location = new System.Drawing.Point(482, 39);
            this.syokumukyuu.Name = "syokumukyuu";
            this.syokumukyuu.ReadOnly = true;
            this.syokumukyuu.Size = new System.Drawing.Size(100, 19);
            this.syokumukyuu.TabIndex = 3;
            this.syokumukyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // nennreikyuu
            // 
            this.nennreikyuu.Location = new System.Drawing.Point(482, 67);
            this.nennreikyuu.Name = "nennreikyuu";
            this.nennreikyuu.ReadOnly = true;
            this.nennreikyuu.Size = new System.Drawing.Size(100, 19);
            this.nennreikyuu.TabIndex = 3;
            this.nennreikyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // gakurekikyuu
            // 
            this.gakurekikyuu.Location = new System.Drawing.Point(482, 95);
            this.gakurekikyuu.Name = "gakurekikyuu";
            this.gakurekikyuu.ReadOnly = true;
            this.gakurekikyuu.Size = new System.Drawing.Size(100, 19);
            this.gakurekikyuu.TabIndex = 3;
            this.gakurekikyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // tanka
            // 
            this.tanka.Location = new System.Drawing.Point(465, 38);
            this.tanka.Name = "tanka";
            this.tanka.ReadOnly = true;
            this.tanka.Size = new System.Drawing.Size(100, 19);
            this.tanka.TabIndex = 3;
            this.tanka.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // kihongoukei
            // 
            this.kihongoukei.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.kihongoukei.Location = new System.Drawing.Point(482, 151);
            this.kihongoukei.Name = "kihongoukei";
            this.kihongoukei.ReadOnly = true;
            this.kihongoukei.Size = new System.Drawing.Size(100, 23);
            this.kihongoukei.TabIndex = 3;
            this.kihongoukei.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.Location = new System.Drawing.Point(348, 228);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(128, 16);
            this.label11.TabIndex = 2;
            this.label11.Text = "日給　(合計/21.5)";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(51, 230);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(85, 12);
            this.label12.TabIndex = 2;
            this.label12.Text = "他手当　加算額";
            // 
            // nenrei
            // 
            this.nenrei.Location = new System.Drawing.Point(152, 100);
            this.nenrei.Name = "nenrei";
            this.nenrei.ReadOnly = true;
            this.nenrei.Size = new System.Drawing.Size(40, 19);
            this.nenrei.TabIndex = 3;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(152, 44);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(121, 19);
            this.dateTimePicker2.TabIndex = 1;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(71, 48);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(65, 12);
            this.label13.TabIndex = 2;
            this.label13.Text = "入社年月日";
            this.label13.Click += new System.EventHandler(this.label1_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(107, 108);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(29, 12);
            this.label14.TabIndex = 2;
            this.label14.Text = "年齢";
            // 
            // comboBoxkasan
            // 
            this.comboBoxkasan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxkasan.FormattingEnabled = true;
            this.comboBoxkasan.Location = new System.Drawing.Point(152, 227);
            this.comboBoxkasan.Name = "comboBoxkasan";
            this.comboBoxkasan.Size = new System.Drawing.Size(121, 20);
            this.comboBoxkasan.TabIndex = 4;
            this.comboBoxkasan.SelectedIndexChanged += new System.EventHandler(this.comboBoxkasan_SelectedIndexChanged);
            // 
            // tuukinhouhou
            // 
            this.tuukinhouhou.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tuukinhouhou.FormattingEnabled = true;
            this.tuukinhouhou.Location = new System.Drawing.Point(135, 33);
            this.tuukinhouhou.Name = "tuukinhouhou";
            this.tuukinhouhou.Size = new System.Drawing.Size(121, 20);
            this.tuukinhouhou.TabIndex = 4;
            this.tuukinhouhou.SelectedIndexChanged += new System.EventHandler(this.tuukinhouhou_SelectedIndexChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(66, 38);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(53, 12);
            this.label15.TabIndex = 2;
            this.label15.Text = "通勤方法";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(66, 67);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(53, 12);
            this.label16.TabIndex = 2;
            this.label16.Text = "通勤距離";
            // 
            // kyori
            // 
            this.kyori.DecimalPlaces = 1;
            this.kyori.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.kyori.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.kyori.Location = new System.Drawing.Point(135, 65);
            this.kyori.Margin = new System.Windows.Forms.Padding(0);
            this.kyori.Name = "kyori";
            this.kyori.Size = new System.Drawing.Size(121, 19);
            this.kyori.TabIndex = 17;
            this.kyori.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.kyori.ThousandsSeparator = true;
            this.kyori.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            this.kyori.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.kyori.ValueChanged += new System.EventHandler(this.kyori_ValueChanged);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(66, 105);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(53, 12);
            this.label17.TabIndex = 2;
            this.label17.Text = "出勤日数";
            // 
            // kinmu
            // 
            this.kinmu.DecimalPlaces = 1;
            this.kinmu.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.kinmu.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.kinmu.Location = new System.Drawing.Point(135, 97);
            this.kinmu.Margin = new System.Windows.Forms.Padding(0);
            this.kinmu.Maximum = new decimal(new int[] {
            31,
            0,
            0,
            0});
            this.kinmu.Name = "kinmu";
            this.kinmu.Size = new System.Drawing.Size(121, 19);
            this.kinmu.TabIndex = 17;
            this.kinmu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.kinmu.ThousandsSeparator = true;
            this.kinmu.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            this.kinmu.Value = new decimal(new int[] {
            21,
            0,
            0,
            0});
            this.kinmu.ValueChanged += new System.EventHandler(this.kinmu_ValueChanged);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(388, 41);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(71, 12);
            this.label18.TabIndex = 2;
            this.label18.Text = "通勤1日単価";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(362, 70);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(97, 12);
            this.label19.TabIndex = 2;
            this.label19.Text = "通勤手当　非課税";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(374, 104);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(85, 12);
            this.label20.TabIndex = 2;
            this.label20.Text = "通勤手当　課税";
            this.label20.Click += new System.EventHandler(this.label20_Click);
            // 
            // hi
            // 
            this.hi.Location = new System.Drawing.Point(465, 67);
            this.hi.Name = "hi";
            this.hi.ReadOnly = true;
            this.hi.Size = new System.Drawing.Size(100, 19);
            this.hi.TabIndex = 3;
            this.hi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // ka
            // 
            this.ka.Location = new System.Drawing.Point(465, 101);
            this.ka.Name = "ka";
            this.ka.ReadOnly = true;
            this.ka.Size = new System.Drawing.Size(100, 19);
            this.ka.TabIndex = 3;
            this.ka.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tuukinhouhou);
            this.groupBox1.Controls.Add(this.kinmu);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.ka);
            this.groupBox1.Controls.Add(this.kyori);
            this.groupBox1.Controls.Add(this.hi);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.tanka);
            this.groupBox1.Controls.Add(this.label17);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.label19);
            this.groupBox1.Controls.Add(this.label20);
            this.groupBox1.Location = new System.Drawing.Point(17, 268);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(574, 149);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "通勤手当";
            // 
            // keikenkyuu
            // 
            this.keikenkyuu.Location = new System.Drawing.Point(482, 125);
            this.keikenkyuu.Name = "keikenkyuu";
            this.keikenkyuu.ReadOnly = true;
            this.keikenkyuu.Size = new System.Drawing.Size(100, 19);
            this.keikenkyuu.TabIndex = 3;
            this.keikenkyuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // comboBoxritou
            // 
            this.comboBoxritou.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxritou.FormattingEnabled = true;
            this.comboBoxritou.Location = new System.Drawing.Point(152, 198);
            this.comboBoxritou.Name = "comboBoxritou";
            this.comboBoxritou.Size = new System.Drawing.Size(121, 20);
            this.comboBoxritou.TabIndex = 4;
            this.comboBoxritou.SelectedIndexChanged += new System.EventHandler(this.comboBoxritou_SelectedIndexChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(39, 201);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(97, 12);
            this.label21.TabIndex = 2;
            this.label21.Text = "離島手当　加算額";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label22.Location = new System.Drawing.Point(382, 197);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(88, 16);
            this.label22.TabIndex = 2;
            this.label22.Text = "支給額合計";
            // 
            // sikyuugoukei
            // 
            this.sikyuugoukei.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.sikyuugoukei.Location = new System.Drawing.Point(482, 192);
            this.sikyuugoukei.Name = "sikyuugoukei";
            this.sikyuugoukei.ReadOnly = true;
            this.sikyuugoukei.Size = new System.Drawing.Size(100, 23);
            this.sikyuugoukei.TabIndex = 3;
            this.sikyuugoukei.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // Nikkyu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 440);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.comboBoxritou);
            this.Controls.Add(this.comboBoxkasan);
            this.Controls.Add(this.sikyuugoukei);
            this.Controls.Add(this.kihongoukei);
            this.Controls.Add(this.nenrei);
            this.Controls.Add(this.gakurekikyuu);
            this.Controls.Add(this.nennreikyuu);
            this.Controls.Add(this.syokumukyuu);
            this.Controls.Add(this.keikenkyuu);
            this.Controls.Add(this.honyuu);
            this.Controls.Add(this.nikkyuu);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.comboBoxKeiken);
            this.Controls.Add(this.comboBoxSyoku);
            this.Controls.Add(this.comboBoxGaku);
            this.Name = "Nikkyu";
            this.Text = "給与月額";
            this.Load += new System.EventHandler(this.Nikkyu_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kyori)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kinmu)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxGaku;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxKeiken;
        private System.Windows.Forms.TextBox nikkyuu;
        private System.Windows.Forms.ComboBox comboBoxSyoku;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox honyuu;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox syokumukyuu;
        private System.Windows.Forms.TextBox nennreikyuu;
        private System.Windows.Forms.TextBox gakurekikyuu;
        private System.Windows.Forms.TextBox tanka;
        private System.Windows.Forms.TextBox kihongoukei;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox nenrei;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.ComboBox comboBoxkasan;
        private System.Windows.Forms.ComboBox tuukinhouhou;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.NumericUpDown kyori;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.NumericUpDown kinmu;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox hi;
        private System.Windows.Forms.TextBox ka;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox keikenkyuu;
        private System.Windows.Forms.ComboBox comboBoxritou;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox sikyuugoukei;
    }
}