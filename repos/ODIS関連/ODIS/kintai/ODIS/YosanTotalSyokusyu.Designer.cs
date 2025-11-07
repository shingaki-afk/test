
namespace ODIS.ODIS
{
    partial class YosanTotalSyokusyu
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
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
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
            this.dgvyosan.Size = new System.Drawing.Size(1350, 667);
            this.dgvyosan.TabIndex = 0;
            // 
            // cbs
            // 
            this.cbs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbs.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbs.FormattingEnabled = true;
            this.cbs.Location = new System.Drawing.Point(238, 13);
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
            this.cbe.Location = new System.Drawing.Point(357, 13);
            this.cbe.Name = "cbe";
            this.cbe.Size = new System.Drawing.Size(90, 24);
            this.cbe.TabIndex = 1;
            this.cbe.SelectedIndexChanged += new System.EventHandler(this.cbe_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(334, 19);
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
            this.cbzi.Location = new System.Drawing.Point(134, 12);
            this.cbzi.Name = "cbzi";
            this.cbzi.Size = new System.Drawing.Size(98, 24);
            this.cbzi.TabIndex = 62;
            this.cbzi.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(489, 19);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(87, 16);
            this.checkBox1.TabIndex = 63;
            this.checkBox1.Text = "引当金を除く";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(606, 18);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(114, 16);
            this.checkBox2.TabIndex = 65;
            this.checkBox2.Text = "各経費率とか表示";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // cbki
            // 
            this.cbki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbki.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbki.FormattingEnabled = true;
            this.cbki.Location = new System.Drawing.Point(12, 12);
            this.cbki.Name = "cbki";
            this.cbki.Size = new System.Drawing.Size(116, 24);
            this.cbki.TabIndex = 62;
            this.cbki.SelectedIndexChanged += new System.EventHandler(this.cbki_SelectedIndexChanged);
            // 
            // YosanTotalSyokusyu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1350, 729);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.cbki);
            this.Controls.Add(this.cbzi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbe);
            this.Controls.Add(this.cbs);
            this.Controls.Add(this.dgvyosan);
            this.Name = "YosanTotalSyokusyu";
            this.Text = "541_集計職種別";
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
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.ComboBox cbki;
    }
}