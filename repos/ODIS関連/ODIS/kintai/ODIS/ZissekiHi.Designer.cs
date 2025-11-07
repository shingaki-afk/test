
namespace ODIS.ODIS
{
    partial class ZissekiHi
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
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
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
            this.dgvyosan.Size = new System.Drawing.Size(1468, 899);
            this.dgvyosan.TabIndex = 0;
            this.dgvyosan.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvyosan_CellFormatting);
            this.dgvyosan.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvyosan_CellPainting);
            // 
            // cbs
            // 
            this.cbs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbs.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbs.FormattingEnabled = true;
            this.cbs.Location = new System.Drawing.Point(123, 16);
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
            this.cbe.Location = new System.Drawing.Point(242, 16);
            this.cbe.Name = "cbe";
            this.cbe.Size = new System.Drawing.Size(90, 24);
            this.cbe.TabIndex = 1;
            this.cbe.SelectedIndexChanged += new System.EventHandler(this.cbe_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(219, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "～";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(0, 15);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(77, 24);
            this.comboBox1.TabIndex = 61;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1397, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 62;
            this.label2.Text = "単位：千円";
            // 
            // YosanTotal_Zenkihi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1468, 961);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbe);
            this.Controls.Add(this.cbs);
            this.Controls.Add(this.dgvyosan);
            this.Name = "YosanTotal_Zenkihi";
            this.Text = "予算集計_前期比";
            ((System.ComponentModel.ISupportInitialize)(this.dgvyosan)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvyosan;
        private System.Windows.Forms.ComboBox cbs;
        private System.Windows.Forms.ComboBox cbe;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label2;
    }
}