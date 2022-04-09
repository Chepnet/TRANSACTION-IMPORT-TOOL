namespace ReadExcel
{
    partial class frmImportLoanTypes
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
            this.btnReadAndImportData = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtFilenName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnReadAndImportData
            // 
            this.btnReadAndImportData.Location = new System.Drawing.Point(140, 80);
            this.btnReadAndImportData.Name = "btnReadAndImportData";
            this.btnReadAndImportData.Size = new System.Drawing.Size(255, 40);
            this.btnReadAndImportData.TabIndex = 7;
            this.btnReadAndImportData.Text = "Read And Import Loan Types";
            this.btnReadAndImportData.UseVisualStyleBackColor = true;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(463, 28);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(82, 33);
            this.btnBrowse.TabIndex = 6;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtFilenName
            // 
            this.txtFilenName.Location = new System.Drawing.Point(117, 26);
            this.txtFilenName.Name = "txtFilenName";
            this.txtFilenName.Size = new System.Drawing.Size(323, 26);
            this.txtFilenName.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "File Name:";
            // 
            // frmImportLoanTypes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(620, 167);
            this.Controls.Add(this.btnReadAndImportData);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFilenName);
            this.Controls.Add(this.label1);
            this.Name = "frmImportLoanTypes";
            this.ShowIcon = false;
            this.Text = "Import Loan Types";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnReadAndImportData;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtFilenName;
        private System.Windows.Forms.Label label1;
    }
}