namespace ODIS
{
    partial class Main
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.monthout = new System.Windows.Forms.Label();
            this.month1 = new System.Windows.Forms.Label();
            this.month2 = new System.Windows.Forms.Label();
            this.month3 = new System.Windows.Forms.Label();
            this.month4 = new System.Windows.Forms.Label();
            this.month5 = new System.Windows.Forms.Label();
            this.month6 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.treeView1.ImageIndex = 0;
            this.treeView1.ImageList = this.imageList1;
            this.treeView1.Location = new System.Drawing.Point(11, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.SelectedImageIndex = 0;
            this.treeView1.ShowNodeToolTips = true;
            this.treeView1.Size = new System.Drawing.Size(286, 746);
            this.treeView1.TabIndex = 9;
            this.treeView1.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseDoubleClick);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "kaikei.jpeg");
            this.imageList1.Images.SetKeyName(1, "uriage.png");
            this.imageList1.Images.SetKeyName(2, "money");
            this.imageList1.Images.SetKeyName(3, "admin");
            this.imageList1.Images.SetKeyName(4, "keisuu");
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(303, 12);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(408, 28);
            this.comboBox1.TabIndex = 16;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // comboBox2
            // 
            this.comboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox2.Enabled = false;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(1068, 10);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(163, 28);
            this.comboBox2.TabIndex = 18;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(303, 80);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(929, 678);
            this.dataGridView1.TabIndex = 19;
            this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);
            // 
            // comboBox3
            // 
            this.comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Location = new System.Drawing.Point(303, 46);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(173, 28);
            this.comboBox3.TabIndex = 20;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 7;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 148F));
            this.tableLayoutPanel1.Controls.Add(this.monthout, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.month1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.month2, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.month3, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.month4, 4, 0);
            this.tableLayoutPanel1.Controls.Add(this.month5, 5, 0);
            this.tableLayoutPanel1.Controls.Add(this.month6, 6, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(482, 46);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(750, 28);
            this.tableLayoutPanel1.TabIndex = 21;
            // 
            // monthout
            // 
            this.monthout.AutoSize = true;
            this.monthout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.monthout.Location = new System.Drawing.Point(1, 1);
            this.monthout.Margin = new System.Windows.Forms.Padding(0);
            this.monthout.Name = "monthout";
            this.monthout.Size = new System.Drawing.Size(75, 26);
            this.monthout.TabIndex = 0;
            this.monthout.Text = "手遅れ";
            this.monthout.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month1
            // 
            this.month1.AutoSize = true;
            this.month1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month1.Location = new System.Drawing.Point(77, 1);
            this.month1.Margin = new System.Windows.Forms.Padding(0);
            this.month1.Name = "month1";
            this.month1.Size = new System.Drawing.Size(120, 26);
            this.month1.TabIndex = 0;
            this.month1.Text = "今月中に取得";
            this.month1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month2
            // 
            this.month2.AutoSize = true;
            this.month2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month2.Location = new System.Drawing.Point(198, 1);
            this.month2.Margin = new System.Windows.Forms.Padding(0);
            this.month2.Name = "month2";
            this.month2.Size = new System.Drawing.Size(100, 26);
            this.month2.TabIndex = 0;
            this.month2.Text = "残り2ヶ月";
            this.month2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month3
            // 
            this.month3.AutoSize = true;
            this.month3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month3.Location = new System.Drawing.Point(299, 1);
            this.month3.Margin = new System.Windows.Forms.Padding(0);
            this.month3.Name = "month3";
            this.month3.Size = new System.Drawing.Size(100, 26);
            this.month3.TabIndex = 0;
            this.month3.Text = "残り3ヶ月";
            this.month3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month4
            // 
            this.month4.AutoSize = true;
            this.month4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month4.Location = new System.Drawing.Point(400, 1);
            this.month4.Margin = new System.Windows.Forms.Padding(0);
            this.month4.Name = "month4";
            this.month4.Size = new System.Drawing.Size(100, 26);
            this.month4.TabIndex = 23;
            this.month4.Text = "残り3ヶ月";
            this.month4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month5
            // 
            this.month5.AutoSize = true;
            this.month5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month5.Location = new System.Drawing.Point(501, 1);
            this.month5.Margin = new System.Windows.Forms.Padding(0);
            this.month5.Name = "month5";
            this.month5.Size = new System.Drawing.Size(100, 26);
            this.month5.TabIndex = 23;
            this.month5.Text = "残り3ヶ月";
            this.month5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // month6
            // 
            this.month6.AutoSize = true;
            this.month6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.month6.Location = new System.Drawing.Point(602, 1);
            this.month6.Margin = new System.Windows.Forms.Padding(0);
            this.month6.Name = "month6";
            this.month6.Size = new System.Drawing.Size(148, 26);
            this.month6.TabIndex = 23;
            this.month6.Text = "残り3ヶ月";
            this.month6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(737, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(233, 26);
            this.textBox1.TabIndex = 22;
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(976, 10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 30);
            this.button1.TabIndex = 23;
            this.button1.Text = "絞込";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1237, 761);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.treeView1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "ODIS";
            this.Shown += new System.EventHandler(this.Main_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label monthout;
        private System.Windows.Forms.Label month1;
        private System.Windows.Forms.Label month2;
        private System.Windows.Forms.Label month3;
        private System.Windows.Forms.Label month4;
        private System.Windows.Forms.Label month5;
        private System.Windows.Forms.Label month6;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
    }
}