
namespace ReadExcel
{
    partial class frmSearchLoanTypes
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
            this.objLoantypes = new BrightIdeasSoftware.ObjectListView();
            this.olvColumn1 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            this.olvColumn2 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            ((System.ComponentModel.ISupportInitialize)(this.objLoantypes)).BeginInit();
            this.SuspendLayout();
            // 
            // objLoantypes
            // 
            this.objLoantypes.AllColumns.Add(this.olvColumn1);
            this.objLoantypes.AllColumns.Add(this.olvColumn2);
            this.objLoantypes.CellEditUseWholeCell = false;
            this.objLoantypes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.olvColumn1,
            this.olvColumn2});
            this.objLoantypes.Cursor = System.Windows.Forms.Cursors.Default;
            this.objLoantypes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objLoantypes.FullRowSelect = true;
            this.objLoantypes.GridLines = true;
            this.objLoantypes.HideSelection = false;
            this.objLoantypes.Location = new System.Drawing.Point(0, 0);
            this.objLoantypes.Name = "objLoantypes";
            this.objLoantypes.ShowGroups = false;
            this.objLoantypes.Size = new System.Drawing.Size(359, 474);
            this.objLoantypes.TabIndex = 2;
            this.objLoantypes.UseCompatibleStateImageBehavior = false;
            this.objLoantypes.View = System.Windows.Forms.View.Details;
            this.objLoantypes.SelectedIndexChanged += new System.EventHandler(this.objSharetypes_SelectedIndexChanged);
            this.objLoantypes.DoubleClick += new System.EventHandler(this.objLoantypes_DoubleClick);
            // 
            // olvColumn1
            // 
            this.olvColumn1.AspectName = "LoanTypecode";
            this.olvColumn1.Text = "Loan Type Code";
            this.olvColumn1.Width = 133;
            // 
            // olvColumn2
            // 
            this.olvColumn2.AspectName = "LoanTypename";
            this.olvColumn2.Text = "Loan Type Name";
            this.olvColumn2.Width = 177;
            // 
            // frmSearchLoanTypes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(359, 474);
            this.Controls.Add(this.objLoantypes);
            this.Name = "frmSearchLoanTypes";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Search Loan Type";
            this.Load += new System.EventHandler(this.frmSearchLoanTypes_Load);
            ((System.ComponentModel.ISupportInitialize)(this.objLoantypes)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private BrightIdeasSoftware.ObjectListView objLoantypes;
        private BrightIdeasSoftware.OLVColumn olvColumn1;
        private BrightIdeasSoftware.OLVColumn olvColumn2;
    }
}