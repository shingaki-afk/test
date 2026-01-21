
namespace ODIS.ODIS
{
    partial class YosanTotalSyokusyu_Zenkihi
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
            this.cbs = new System.Windows.Forms.ComboBox();
            this.cbe = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbzi = new System.Windows.Forms.ComboBox();
            this.cbhikiate = new System.Windows.Forms.CheckBox();
            this.cbhikaku = new System.Windows.Forms.ComboBox();
            this.cbziku = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cbtani = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbki = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvyosan
            // 
            this.dgvyosan.AllowUserToAddRows = false;
            this.dgvyosan.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvyosan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvyosan.Location = new System.Drawing.Point(1, 50);
            this.dgvyosan.Name = "dgvyosan";
            this.dgvyosan.RowTemplate.Height = 21;
            this.dgvyosan.Size = new System.Drawing.Size(1349, 677);
            this.dgvyosan.TabIndex = 0;
            this.dgvyosan.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvyosan_CellFormatting);
            this.dgvyosan.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvyosan_CellPainting);
            // 
            // cbs
            // 
            this.cbs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbs.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbs.FormattingEnabled = true;
            this.cbs.Location = new System.Drawing.Point(262, 13);
            this.cbs.Name = "cbs";
            this.cbs.Size = new System.Drawing.Size(90, 24);
            this.cbs.TabIndex = 1;
            this.cbs.SelectedIndexChanged += new System.EventHandler(this.cbs_SelectedIndexChanged);
            // 
            // cbe
            // 
            this.cbe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbe.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbe.FormattingEnabled = true;
            this.cbe.Location = new System.Drawing.Point(381, 13);
            this.cbe.Name = "cbe";
            this.cbe.Size = new System.Drawing.Size(90, 24);
            this.cbe.TabIndex = 1;
            this.cbe.SelectedIndexChanged += new System.EventHandler(this.cbe_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(358, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "～";
            // 
            // cbzi
            // 
            this.cbzi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbzi.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbzi.FormattingEnabled = true;
            this.cbzi.Location = new System.Drawing.Point(160, 13);
            this.cbzi.Name = "cbzi";
            this.cbzi.Size = new System.Drawing.Size(77, 24);
            this.cbzi.TabIndex = 62;
            this.cbzi.SelectedIndexChanged += new System.EventHandler(this.cbzi_SelectedIndexChanged);
            // 
            // cbhikiate
            // 
            this.cbhikiate.AutoSize = true;
            this.cbhikiate.Location = new System.Drawing.Point(522, 19);
            this.cbhikiate.Name = "cbhikiate";
            this.cbhikiate.Size = new System.Drawing.Size(87, 16);
            this.cbhikiate.TabIndex = 64;
            this.cbhikiate.Text = "引当金を除く";
            this.cbhikiate.UseVisualStyleBackColor = true;
            this.cbhikiate.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // cbhikaku
            // 
            this.cbhikaku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbhikaku.FormattingEnabled = true;
            this.cbhikaku.Location = new System.Drawing.Point(661, 28);
            this.cbhikaku.Name = "cbhikaku";
            this.cbhikaku.Size = new System.Drawing.Size(121, 20);
            this.cbhikaku.TabIndex = 67;
            this.cbhikaku.SelectedIndexChanged += new System.EventHandler(this.cbhikaku_SelectedIndexChanged);
            // 
            // cbziku
            // 
            this.cbziku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbziku.FormattingEnabled = true;
            this.cbziku.Location = new System.Drawing.Point(661, 4);
            this.cbziku.Name = "cbziku";
            this.cbziku.Size = new System.Drawing.Size(121, 20);
            this.cbziku.TabIndex = 68;
            this.cbziku.SelectedIndexChanged += new System.EventHandler(this.cbziku_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(626, 31);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 65;
            this.label4.Text = "比較";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(638, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 66;
            this.label3.Text = "軸";
            // 
            // cbtani
            // 
            this.cbtani.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbtani.FormattingEnabled = true;
            this.cbtani.Location = new System.Drawing.Point(1260, 16);
            this.cbtani.Name = "cbtani";
            this.cbtani.Size = new System.Drawing.Size(78, 20);
            this.cbtani.TabIndex = 70;
            this.cbtani.SelectedIndexChanged += new System.EventHandler(this.cbtani_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1219, 19);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 12);
            this.label5.TabIndex = 69;
            this.label5.Text = "単位：";
            // 
            // cbki
            // 
            this.cbki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbki.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbki.FormattingEnabled = true;
            this.cbki.Location = new System.Drawing.Point(12, 13);
            this.cbki.Name = "cbki";
            this.cbki.Size = new System.Drawing.Size(132, 24);
            this.cbki.TabIndex = 62;
            this.cbki.SelectedIndexChanged += new System.EventHandler(this.cbki_SelectedIndexChanged);
            // 
            // YosanTotalSyokusyu_Zenkihi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1350, 729);
            this.Controls.Add(this.dgvyosan);
            this.Controls.Add(this.cbtani);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbhikaku);
            this.Controls.Add(this.cbziku);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbhikiate);
            this.Controls.Add(this.cbki);
            this.Controls.Add(this.cbzi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbe);
            this.Controls.Add(this.cbs);
            this.Name = "YosanTotalSyokusyu_Zenkihi";
            this.Text = "542_集計職種別_差額";
            this.Load += new System.EventHandler(this.YosanTotalSyokusyu_Zenkihi_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvyosan;
        private System.Windows.Forms.ComboBox cbs;
        private System.Windows.Forms.ComboBox cbe;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbzi;
        private System.Windows.Forms.CheckBox cbhikiate;
        private System.Windows.Forms.ComboBox cbhikaku;
        private System.Windows.Forms.ComboBox cbziku;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbtani;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbki;
    }
}