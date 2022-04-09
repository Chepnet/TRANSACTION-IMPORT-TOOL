
namespace ReadExcel
{
    partial class frmSearchImportationOrder
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSearchImportationOrder));
            this.objListImportationOrder = new BrightIdeasSoftware.ObjectListView();
            this.olvColumn1 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            this.olvColumn2 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.saveToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.chkCheckAll = new System.Windows.Forms.CheckBox();
            this.lblImportname = new System.Windows.Forms.Label();
            this.cmbFormat = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.objListImportationOrder)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // objListImportationOrder
            // 
            this.objListImportationOrder.AllColumns.Add(this.olvColumn1);
            this.objListImportationOrder.AllColumns.Add(this.olvColumn2);
            this.objListImportationOrder.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.objListImportationOrder.CellEditActivation = BrightIdeasSoftware.ObjectListView.CellEditActivateMode.DoubleClick;
            this.objListImportationOrder.CellEditUseWholeCell = false;
            this.objListImportationOrder.CheckBoxes = true;
            this.objListImportationOrder.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.olvColumn1,
            this.olvColumn2});
            this.objListImportationOrder.Cursor = System.Windows.Forms.Cursors.Default;
            this.objListImportationOrder.FullRowSelect = true;
            this.objListImportationOrder.GridLines = true;
            this.objListImportationOrder.HideSelection = false;
            this.objListImportationOrder.Location = new System.Drawing.Point(0, 104);
            this.objListImportationOrder.Name = "objListImportationOrder";
            this.objListImportationOrder.ShowGroups = false;
            this.objListImportationOrder.Size = new System.Drawing.Size(522, 383);
            this.objListImportationOrder.TabIndex = 2;
            this.objListImportationOrder.UseCompatibleStateImageBehavior = false;
            this.objListImportationOrder.View = System.Windows.Forms.View.Details;
            this.objListImportationOrder.SelectedIndexChanged += new System.EventHandler(this.objSharetypes_SelectedIndexChanged);
            // 
            // olvColumn1
            // 
            this.olvColumn1.AspectName = "ProductName";
            this.olvColumn1.Text = "Product Name";
            this.olvColumn1.Width = 133;
            // 
            // olvColumn2
            // 
            this.olvColumn2.AspectName = "PositionId";
            this.olvColumn2.Text = "Position";
            this.olvColumn2.Width = 177;
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(522, 31);
            this.toolStrip1.TabIndex = 3;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // saveToolStripButton
            // 
            this.saveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.saveToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripButton.Image")));
            this.saveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolStripButton.Name = "saveToolStripButton";
            this.saveToolStripButton.Size = new System.Drawing.Size(29, 28);
            this.saveToolStripButton.Text = "&Save";
            this.saveToolStripButton.Click += new System.EventHandler(this.saveToolStripButton_Click);
            // 
            // chkCheckAll
            // 
            this.chkCheckAll.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chkCheckAll.AutoSize = true;
            this.chkCheckAll.Location = new System.Drawing.Point(3, 6);
            this.chkCheckAll.Name = "chkCheckAll";
            this.chkCheckAll.Size = new System.Drawing.Size(88, 21);
            this.chkCheckAll.TabIndex = 4;
            this.chkCheckAll.Text = "Check All";
            this.chkCheckAll.UseVisualStyleBackColor = true;
            this.chkCheckAll.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // lblImportname
            // 
            this.lblImportname.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblImportname.AutoSize = true;
            this.lblImportname.Location = new System.Drawing.Point(12, 49);
            this.lblImportname.Name = "lblImportname";
            this.lblImportname.Size = new System.Drawing.Size(92, 17);
            this.lblImportname.TabIndex = 5;
            this.lblImportname.Text = "Import Name:";
            // 
            // cmbFormat
            // 
            this.cmbFormat.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFormat.FormattingEnabled = true;
            this.cmbFormat.Location = new System.Drawing.Point(101, 49);
            this.cmbFormat.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cmbFormat.Name = "cmbFormat";
            this.cmbFormat.Size = new System.Drawing.Size(389, 24);
            this.cmbFormat.TabIndex = 6;
            this.cmbFormat.SelectedIndexChanged += new System.EventHandler(this.cmbFormat_SelectedIndexChanged);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.chkCheckAll);
            this.panel1.Location = new System.Drawing.Point(0, 493);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(522, 34);
            this.panel1.TabIndex = 7;
            // 
            // frmSearchImportationOrder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(522, 532);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.cmbFormat);
            this.Controls.Add(this.lblImportname);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.objListImportationOrder);
            this.Name = "frmSearchImportationOrder";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Search Importation Order";
            this.Load += new System.EventHandler(this.frmSearchMigrationOrder_Load);
            ((System.ComponentModel.ISupportInitialize)(this.objListImportationOrder)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private BrightIdeasSoftware.ObjectListView objListImportationOrder;
        private BrightIdeasSoftware.OLVColumn olvColumn1;
        private BrightIdeasSoftware.OLVColumn olvColumn2;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton saveToolStripButton;
        private System.Windows.Forms.CheckBox chkCheckAll;
        private System.Windows.Forms.Label lblImportname;
        private System.Windows.Forms.ComboBox cmbFormat;
        private System.Windows.Forms.Panel panel1;
    }
}