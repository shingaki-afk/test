using System.ComponentModel;
using System.Windows.Forms;

namespace ODIS
{
    partial class Form1
    {
        private IContainer components = null;
        private DataGridView reportGrid;
        private DataGridView detailGrid;
        private Button btnImport;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.reportGrid = new System.Windows.Forms.DataGridView();
            this.detailGrid = new System.Windows.Forms.DataGridView();
            this.btnImport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.reportGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.detailGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // reportGrid
            // 
            this.reportGrid.AllowUserToAddRows = false;
            this.reportGrid.AllowUserToDeleteRows = false;
            this.reportGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.reportGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.reportGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.reportGrid.Location = new System.Drawing.Point(12, 56);
            this.reportGrid.MultiSelect = false;
            this.reportGrid.Name = "reportGrid";
            this.reportGrid.ReadOnly = true;
            this.reportGrid.RowHeadersVisible = false;
            this.reportGrid.RowTemplate.Height = 21;
            this.reportGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.reportGrid.Size = new System.Drawing.Size(960, 320);
            this.reportGrid.TabIndex = 0;
            // 
            // detailGrid
            // 
            this.detailGrid.AllowUserToAddRows = false;
            this.detailGrid.AllowUserToDeleteRows = false;
            this.detailGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.detailGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.detailGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.detailGrid.Location = new System.Drawing.Point(12, 392);
            this.detailGrid.MultiSelect = false;
            this.detailGrid.Name = "detailGrid";
            this.detailGrid.ReadOnly = true;
            this.detailGrid.RowHeadersVisible = false;
            this.detailGrid.RowTemplate.Height = 21;
            this.detailGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.detailGrid.Size = new System.Drawing.Size(960, 320);
            this.detailGrid.TabIndex = 1;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(12, 12);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(140, 32);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "データ取込";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 731);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.detailGrid);
            this.Controls.Add(this.reportGrid);
            this.MinimumSize = new System.Drawing.Size(840, 600);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "現場総計数実績（.NET Framework WinForms）";
            ((System.ComponentModel.ISupportInitialize)(this.reportGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.detailGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
    }
}
