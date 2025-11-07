namespace ODIS.ODIS
{
    partial class YosanUp
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
            this.dgvyosan = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dgvlist = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.genbaname = new System.Windows.Forms.Label();
            this.bumonname = new System.Windows.Forms.Label();
            this.bumoncd = new System.Windows.Forms.Label();
            this.genbacd = new System.Windows.Forms.Label();
            this.dgvex = new System.Windows.Forms.DataGridView();
            this.dgvzisseki = new System.Windows.Forms.DataGridView();
            this.label5 = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label23 = new System.Windows.Forms.Label();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.cbzi = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.dgvyosansum = new System.Windows.Forms.DataGridView();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.genba = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.cbki = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvlist)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvex)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvzisseki)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosansum)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvyosan
            // 
            this.dgvyosan.AllowUserToAddRows = false;
            this.dgvyosan.AllowUserToDeleteRows = false;
            this.dgvyosan.AllowUserToResizeRows = false;
            this.dgvyosan.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvyosan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvyosan.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvyosan.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dgvyosan.Location = new System.Drawing.Point(378, 46);
            this.dgvyosan.Name = "dgvyosan";
            this.dgvyosan.RowTemplate.Height = 21;
            this.dgvyosan.Size = new System.Drawing.Size(971, 306);
            this.dgvyosan.TabIndex = 0;
            this.dgvyosan.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvyosan_CellFormatting);
            this.dgvyosan.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
            this.dgvyosan.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(1269, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "更新";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(52, 248);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(320, 19);
            this.textBox1.TabIndex = 3;
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            // 
            // dgvlist
            // 
            this.dgvlist.AllowUserToAddRows = false;
            this.dgvlist.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgvlist.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvlist.Location = new System.Drawing.Point(3, 272);
            this.dgvlist.Name = "dgvlist";
            this.dgvlist.RowTemplate.Height = 21;
            this.dgvlist.Size = new System.Drawing.Size(369, 452);
            this.dgvlist.TabIndex = 4;
            this.dgvlist.SelectionChanged += new System.EventHandler(this.dataGridView2_SelectionChanged);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Outset;
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 141F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 88F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 441F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.genbaname, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.bumonname, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.bumoncd, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.genbacd, 2, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(374, 10);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(732, 28);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // genbaname
            // 
            this.genbaname.AutoSize = true;
            this.genbaname.Dock = System.Windows.Forms.DockStyle.Fill;
            this.genbaname.Location = new System.Drawing.Point(292, 2);
            this.genbaname.Name = "genbaname";
            this.genbaname.Size = new System.Drawing.Size(435, 24);
            this.genbaname.TabIndex = 0;
            this.genbaname.Text = "label1";
            this.genbaname.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // bumonname
            // 
            this.bumonname.AutoSize = true;
            this.bumonname.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bumonname.Location = new System.Drawing.Point(59, 2);
            this.bumonname.Name = "bumonname";
            this.bumonname.Size = new System.Drawing.Size(135, 24);
            this.bumonname.TabIndex = 0;
            this.bumonname.Text = "label1";
            this.bumonname.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // bumoncd
            // 
            this.bumoncd.AutoSize = true;
            this.bumoncd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bumoncd.Location = new System.Drawing.Point(5, 2);
            this.bumoncd.Name = "bumoncd";
            this.bumoncd.Size = new System.Drawing.Size(46, 24);
            this.bumoncd.TabIndex = 0;
            this.bumoncd.Text = "label1";
            this.bumoncd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // genbacd
            // 
            this.genbacd.AutoSize = true;
            this.genbacd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.genbacd.Location = new System.Drawing.Point(202, 2);
            this.genbacd.Name = "genbacd";
            this.genbacd.Size = new System.Drawing.Size(82, 24);
            this.genbacd.TabIndex = 0;
            this.genbacd.Text = "label1";
            this.genbacd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgvex
            // 
            this.dgvex.AllowUserToAddRows = false;
            this.dgvex.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvex.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvex.Location = new System.Drawing.Point(3, 3);
            this.dgvex.MultiSelect = false;
            this.dgvex.Name = "dgvex";
            this.dgvex.RowTemplate.Height = 21;
            this.dgvex.Size = new System.Drawing.Size(965, 295);
            this.dgvex.TabIndex = 6;
            this.dgvex.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellDoubleClick);
            this.dgvex.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView3_CellFormatting);
            // 
            // dgvzisseki
            // 
            this.dgvzisseki.AllowUserToAddRows = false;
            this.dgvzisseki.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvzisseki.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvzisseki.Location = new System.Drawing.Point(3, 3);
            this.dgvzisseki.Name = "dgvzisseki";
            this.dgvzisseki.RowTemplate.Height = 21;
            this.dgvzisseki.Size = new System.Drawing.Size(965, 295);
            this.dgvzisseki.TabIndex = 6;
            this.dgvzisseki.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvzisseki_CellDoubleClick);
            this.dgvzisseki.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView4_CellFormatting);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(2, 72);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(125, 15);
            this.label5.TabIndex = 56;
            this.label5.Text = "□部門_全外/選択";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(3, 91);
            this.checkedListBox1.MultiColumn = true;
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(245, 144);
            this.checkedListBox1.TabIndex = 55;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label23.Location = new System.Drawing.Point(248, 73);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(125, 15);
            this.label23.TabIndex = 58;
            this.label23.Text = "□職種_全外/選択";
            this.label23.Click += new System.EventHandler(this.label23_Click);
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.CheckOnClick = true;
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(250, 91);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(122, 144);
            this.checkedListBox2.TabIndex = 57;
            this.checkedListBox2.SelectedIndexChanged += new System.EventHandler(this.checkedListBox2_SelectedIndexChanged);
            // 
            // cbzi
            // 
            this.cbzi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbzi.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbzi.FormattingEnabled = true;
            this.cbzi.Location = new System.Drawing.Point(171, 7);
            this.cbzi.Name = "cbzi";
            this.cbzi.Size = new System.Drawing.Size(77, 24);
            this.cbzi.TabIndex = 60;
            this.cbzi.SelectedIndexChanged += new System.EventHandler(this.cbzi_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("MS UI Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(1108, -1);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(19, 19);
            this.label7.TabIndex = 7;
            this.label7.Text = "-";
            this.label7.Visible = false;
            // 
            // dgvyosansum
            // 
            this.dgvyosansum.AllowUserToAddRows = false;
            this.dgvyosansum.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvyosansum.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvyosansum.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvyosansum.Location = new System.Drawing.Point(378, 353);
            this.dgvyosansum.Name = "dgvyosansum";
            this.dgvyosansum.ReadOnly = true;
            this.dgvyosansum.RowTemplate.Height = 21;
            this.dgvyosansum.Size = new System.Drawing.Size(971, 41);
            this.dgvyosansum.TabIndex = 0;
            this.dgvyosansum.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView1_CellFormatting);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(1121, 26);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(145, 16);
            this.checkBox1.TabIndex = 61;
            this.checkBox1.Text = "入力月以降を同額にする";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tabControl1.Location = new System.Drawing.Point(374, 402);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(979, 330);
            this.tabControl1.TabIndex = 62;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgvex);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(971, 301);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "前年度 53期(2024年4月～2025年3月)";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgvzisseki);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(971, 301);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "実績 54期(2025年4月～2026年3月)";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // genba
            // 
            this.genba.AutoSize = true;
            this.genba.Location = new System.Drawing.Point(5, 252);
            this.genba.Name = "genba";
            this.genba.Size = new System.Drawing.Size(41, 12);
            this.genba.TabIndex = 63;
            this.genba.Text = "現場名";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(1210, 5);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(53, 12);
            this.linkLabel1.TabIndex = 64;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "入力方法";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(205, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 15);
            this.label1.TabIndex = 65;
            this.label1.Text = "label1";
            // 
            // cbki
            // 
            this.cbki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbki.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbki.FormattingEnabled = true;
            this.cbki.Location = new System.Drawing.Point(5, 7);
            this.cbki.Name = "cbki";
            this.cbki.Size = new System.Drawing.Size(160, 24);
            this.cbki.TabIndex = 60;
            this.cbki.SelectedIndexChanged += new System.EventHandler(this.cbki_SelectedIndexChanged);
            // 
            // YosanUp54
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1350, 729);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.genba);
            this.Controls.Add(this.dgvyosan);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.cbki);
            this.Controls.Add(this.cbzi);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.dgvlist);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dgvyosansum);
            this.Name = "YosanUp54";
            this.Text = "予算更新画面";
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvlist)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvex)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvzisseki)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosansum)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvyosan;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dgvlist;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label bumonname;
        private System.Windows.Forms.Label genbaname;
        private System.Windows.Forms.DataGridView dgvex;
        private System.Windows.Forms.DataGridView dgvzisseki;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.ComboBox cbzi;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label bumoncd;
        private System.Windows.Forms.Label genbacd;
        private System.Windows.Forms.DataGridView dgvyosansum;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label genba;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbki;
    }
}