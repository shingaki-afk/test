
namespace ODIS.ODIS
{
    partial class YosanTotal_Zenkihi
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
            this.label2 = new System.Windows.Forms.Label();
            this.cbziku = new System.Windows.Forms.ComboBox();
            this.cbhikaku = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbtani = new System.Windows.Forms.ComboBox();
            this.cbhikiate = new System.Windows.Forms.CheckBox();
            this.cbki = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvyosan
            // 
            this.dgvyosan.AllowUserToAddRows = false;
            this.dgvyosan.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvyosan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvyosan.Location = new System.Drawing.Point(0, 50);
            this.dgvyosan.Name = "dgvyosan";
            this.dgvyosan.RowTemplate.Height = 21;
            this.dgvyosan.Size = new System.Drawing.Size(1370, 899);
            this.dgvyosan.TabIndex = 0;
            this.dgvyosan.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvyosan_CellFormatting);
            this.dgvyosan.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvyosan_CellPainting);
            // 
            // cbs
            // 
            this.cbs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbs.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbs.FormattingEnabled = true;
            this.cbs.Location = new System.Drawing.Point(257, 16);
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
            this.cbe.Location = new System.Drawing.Point(376, 16);
            this.cbe.Name = "cbe";
            this.cbe.Size = new System.Drawing.Size(90, 24);
            this.cbe.TabIndex = 1;
            this.cbe.SelectedIndexChanged += new System.EventHandler(this.cbe_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(353, 22);
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
            this.cbzi.Location = new System.Drawing.Point(163, 16);
            this.cbzi.Name = "cbzi";
            this.cbzi.Size = new System.Drawing.Size(77, 24);
            this.cbzi.TabIndex = 61;
            this.cbzi.SelectedIndexChanged += new System.EventHandler(this.cbzi_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1327, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 12);
            this.label2.TabIndex = 62;
            this.label2.Text = "単位：";
            // 
            // cbziku
            // 
            this.cbziku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbziku.FormattingEnabled = true;
            this.cbziku.Location = new System.Drawing.Point(654, 3);
            this.cbziku.Name = "cbziku";
            this.cbziku.Size = new System.Drawing.Size(121, 20);
            this.cbziku.TabIndex = 63;
            this.cbziku.SelectedIndexChanged += new System.EventHandler(this.cbziku_SelectedIndexChanged);
            // 
            // cbhikaku
            // 
            this.cbhikaku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbhikaku.FormattingEnabled = true;
            this.cbhikaku.Location = new System.Drawing.Point(654, 26);
            this.cbhikaku.Name = "cbhikaku";
            this.cbhikaku.Size = new System.Drawing.Size(121, 20);
            this.cbhikaku.TabIndex = 63;
            this.cbhikaku.SelectedIndexChanged += new System.EventHandler(this.cbhikaku_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(619, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 62;
            this.label3.Text = "軸";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(619, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 62;
            this.label4.Text = "比較";
            // 
            // cbtani
            // 
            this.cbtani.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbtani.FormattingEnabled = true;
            this.cbtani.Location = new System.Drawing.Point(1368, 16);
            this.cbtani.Name = "cbtani";
            this.cbtani.Size = new System.Drawing.Size(78, 20);
            this.cbtani.TabIndex = 63;
            this.cbtani.SelectedIndexChanged += new System.EventHandler(this.cbtani_SelectedIndexChanged);
            // 
            // cbhikiate
            // 
            this.cbhikiate.AutoSize = true;
            this.cbhikiate.Location = new System.Drawing.Point(498, 21);
            this.cbhikiate.Name = "cbhikiate";
            this.cbhikiate.Size = new System.Drawing.Size(87, 16);
            this.cbhikiate.TabIndex = 64;
            this.cbhikiate.Text = "引当金を除く";
            this.cbhikiate.UseVisualStyleBackColor = true;
            this.cbhikiate.CheckedChanged += new System.EventHandler(this.cbhikiate_CheckedChanged);
            // 
            // cbki
            // 
            this.cbki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbki.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbki.FormattingEnabled = true;
            this.cbki.Location = new System.Drawing.Point(12, 15);
            this.cbki.Name = "cbki";
            this.cbki.Size = new System.Drawing.Size(131, 24);
            this.cbki.TabIndex = 61;
            this.cbki.SelectedIndexChanged += new System.EventHandler(this.cbki_SelectedIndexChanged);
            // 
            // YosanTotal_Zenkihi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1370, 749);
            this.Controls.Add(this.cbhikiate);
            this.Controls.Add(this.cbtani);
            this.Controls.Add(this.cbhikaku);
            this.Controls.Add(this.cbziku);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbki);
            this.Controls.Add(this.cbzi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbe);
            this.Controls.Add(this.cbs);
            this.Controls.Add(this.dgvyosan);
            this.Name = "YosanTotal_Zenkihi";
            this.Text = "532_集計差額";
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbziku;
        private System.Windows.Forms.ComboBox cbhikaku;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbtani;
        private System.Windows.Forms.CheckBox cbhikiate;
        private System.Windows.Forms.ComboBox cbki;
    }
}