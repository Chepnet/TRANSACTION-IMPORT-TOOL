
namespace ReadExcel
{
    partial class frmMigratemembersKRB
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
            this.btnMigrateShares = new System.Windows.Forms.Button();
            this.btnloans = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnReadAndImportData
            // 
            this.btnReadAndImportData.Location = new System.Drawing.Point(145, 64);
            this.btnReadAndImportData.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnReadAndImportData.Name = "btnReadAndImportData";
            this.btnReadAndImportData.Size = new System.Drawing.Size(227, 32);
            this.btnReadAndImportData.TabIndex = 11;
            this.btnReadAndImportData.Text = "Read And Import Members";
            this.btnReadAndImportData.UseVisualStyleBackColor = true;
            this.btnReadAndImportData.Click += new System.EventHandler(this.btnReadAndImportData_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(432, 22);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(73, 26);
            this.btnBrowse.TabIndex = 10;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtFilenName
            // 
            this.txtFilenName.Location = new System.Drawing.Point(125, 20);
            this.txtFilenName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtFilenName.Name = "txtFilenName";
            this.txtFilenName.Size = new System.Drawing.Size(288, 22);
            this.txtFilenName.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 17);
            this.label1.TabIndex = 8;
            this.label1.Text = "File Name:";
            // 
            // btnMigrateShares
            // 
            this.btnMigrateShares.Location = new System.Drawing.Point(145, 115);
            this.btnMigrateShares.Name = "btnMigrateShares";
            this.btnMigrateShares.Size = new System.Drawing.Size(227, 28);
            this.btnMigrateShares.TabIndex = 12;
            this.btnMigrateShares.Text = "Migrate Shares";
            this.btnMigrateShares.UseVisualStyleBackColor = true;
            this.btnMigrateShares.Click += new System.EventHandler(this.btnMigrateShares_Click);
            // 
            // btnloans
            // 
            this.btnloans.Location = new System.Drawing.Point(145, 163);
            this.btnloans.Name = "btnloans";
            this.btnloans.Size = new System.Drawing.Size(227, 37);
            this.btnloans.TabIndex = 13;
            this.btnloans.Text = "Migrate Loans";
            this.btnloans.UseVisualStyleBackColor = true;
            this.btnloans.Click += new System.EventHandler(this.btn_Click);
            // 
            // frmMigratemembersKRB
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(809, 380);
            this.Controls.Add(this.btnloans);
            this.Controls.Add(this.btnMigrateShares);
            this.Controls.Add(this.btnReadAndImportData);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFilenName);
            this.Controls.Add(this.label1);
            this.Name = "frmMigratemembersKRB";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Migrate members KRB";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnReadAndImportData;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtFilenName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnMigrateShares;
        private System.Windows.Forms.Button btnloans;
    }
}