namespace ODIS.ODIS
{
    partial class SelectIdouDay
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
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.idoudaynew = new C1.Win.C1Input.C1DateEdit();
            ((System.ComponentModel.ISupportInitialize)(this.idoudaynew)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(328, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "決定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(11, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 11;
            this.label3.Text = "異動日";
            // 
            // idoudaynew
            // 
            this.idoudaynew.AllowSpinLoop = false;
            // 
            // 
            // 
            this.idoudaynew.Calendar.DayNameLength = 1;
            this.idoudaynew.Calendar.MaxDate = new System.DateTime(2047, 12, 31, 0, 0, 0, 0);
            this.idoudaynew.Calendar.MinDate = new System.DateTime(2008, 1, 1, 0, 0, 0, 0);
            this.idoudaynew.EmptyAsNull = true;
            this.idoudaynew.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.idoudaynew.FormatType = C1.Win.C1Input.FormatTypeEnum.LongDate;
            this.idoudaynew.ImagePadding = new System.Windows.Forms.Padding(0);
            this.idoudaynew.Location = new System.Drawing.Point(73, 12);
            this.idoudaynew.Name = "idoudaynew";
            this.idoudaynew.Size = new System.Drawing.Size(221, 26);
            this.idoudaynew.TabIndex = 27;
            this.idoudaynew.Tag = null;
            // 
            // SelectIdouDay
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 54);
            this.ControlBox = false;
            this.Controls.Add(this.idoudaynew);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SelectIdouDay";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "異動日選択";
            this.Load += new System.EventHandler(this.SelectEmp_Load);
            ((System.ComponentModel.ISupportInitialize)(this.idoudaynew)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private C1.Win.C1Input.C1DateEdit idoudaynew;
    }
}