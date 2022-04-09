namespace ReadExcel
{
    partial class frmImportLoans
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
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnReadAndImportData
            // 
            this.btnReadAndImportData.Location = new System.Drawing.Point(139, 77);
            this.btnReadAndImportData.Name = "btnReadAndImportData";
            this.btnReadAndImportData.Size = new System.Drawing.Size(255, 40);
            this.btnReadAndImportData.TabIndex = 7;
            this.btnReadAndImportData.Text = "Read And Import Loans";
            this.btnReadAndImportData.UseVisualStyleBackColor = true;
            this.btnReadAndImportData.Click += new System.EventHandler(this.btnReadAndImportData_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(462, 25);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(82, 33);
            this.btnBrowse.TabIndex = 6;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtFilenName
            // 
            this.txtFilenName.Location = new System.Drawing.Point(116, 23);
            this.txtFilenName.Name = "txtFilenName";
            this.txtFilenName.Size = new System.Drawing.Size(323, 26);
            this.txtFilenName.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "File Name:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(139, 123);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(255, 36);
            this.button1.TabIndex = 8;
            this.button1.Text = "MIGRATE LOANS ";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmImportLoans
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(675, 194);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnReadAndImportData);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFilenName);
            this.Controls.Add(this.label1);
            this.Name = "frmImportLoans";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Loans";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnReadAndImportData;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtFilenName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
    }
}