﻿namespace ReadExcel
{
    partial class frmSunData
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
            this.txtfile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.lblread = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.btnMigrateOfficials = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(585, 28);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 35);
            this.button1.TabIndex = 5;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtfile
            // 
            this.txtfile.Location = new System.Drawing.Point(182, 28);
            this.txtfile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtfile.Name = "txtfile";
            this.txtfile.Size = new System.Drawing.Size(382, 26);
            this.txtfile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(63, 32);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "File Name:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(182, 114);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(132, 35);
            this.button2.TabIndex = 6;
            this.button2.Text = "Migrate Members";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // lblread
            // 
            this.lblread.AutoSize = true;
            this.lblread.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblread.ForeColor = System.Drawing.Color.Red;
            this.lblread.Location = new System.Drawing.Point(400, 121);
            this.lblread.Name = "lblread";
            this.lblread.Size = new System.Drawing.Size(19, 20);
            this.lblread.TabIndex = 7;
            this.lblread.Text = "0";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(182, 196);
            this.button3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(132, 35);
            this.button3.TabIndex = 8;
            this.button3.Text = "Benefiaries";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnMigrateOfficials
            // 
            this.btnMigrateOfficials.Location = new System.Drawing.Point(182, 256);
            this.btnMigrateOfficials.Name = "btnMigrateOfficials";
            this.btnMigrateOfficials.Size = new System.Drawing.Size(132, 34);
            this.btnMigrateOfficials.TabIndex = 9;
            this.btnMigrateOfficials.Text = "Migrate Official";
            this.btnMigrateOfficials.UseVisualStyleBackColor = true;
            this.btnMigrateOfficials.Click += new System.EventHandler(this.btnMigrateOfficials_Click);
            // 
            // frmSunData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(916, 424);
            this.Controls.Add(this.btnMigrateOfficials);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.lblread);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtfile);
            this.Controls.Add(this.label1);
            this.Name = "frmSunData";
            this.Text = "frmSunData";
            this.Load += new System.EventHandler(this.frmSunData_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtfile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label lblread;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btnMigrateOfficials;
    }
}