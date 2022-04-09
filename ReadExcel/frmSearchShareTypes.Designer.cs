
namespace ReadExcel
{
    partial class frmSearchShareTypes
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
            this.objSharetypes = new BrightIdeasSoftware.ObjectListView();
            this.olvColumn1 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            this.olvColumn2 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            ((System.ComponentModel.ISupportInitialize)(this.objSharetypes)).BeginInit();
            this.SuspendLayout();
            // 
            // objSharetypes
            // 
            this.objSharetypes.AllColumns.Add(this.olvColumn1);
            this.objSharetypes.AllColumns.Add(this.olvColumn2);
            this.objSharetypes.CellEditUseWholeCell = false;
            this.objSharetypes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.olvColumn1,
            this.olvColumn2});
            this.objSharetypes.Cursor = System.Windows.Forms.Cursors.Default;
            this.objSharetypes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objSharetypes.FullRowSelect = true;
            this.objSharetypes.GridLines = true;
            this.objSharetypes.HideSelection = false;
            this.objSharetypes.Location = new System.Drawing.Point(0, 0);
            this.objSharetypes.Name = "objSharetypes";
            this.objSharetypes.ShowGroups = false;
            this.objSharetypes.Size = new System.Drawing.Size(478, 394);
            this.objSharetypes.TabIndex = 1;
            this.objSharetypes.UseCompatibleStateImageBehavior = false;
            this.objSharetypes.View = System.Windows.Forms.View.Details;
            this.objSharetypes.SelectedIndexChanged += new System.EventHandler(this.objSharetypes_SelectedIndexChanged);
            this.objSharetypes.DoubleClick += new System.EventHandler(this.objSharetypes_DoubleClick);
            // 
            // olvColumn1
            // 
            this.olvColumn1.AspectName = "Sharecode";
            this.olvColumn1.Text = "Share Code";
            this.olvColumn1.Width = 133;
            // 
            // olvColumn2
            // 
            this.olvColumn2.AspectName = "Sharename";
            this.olvColumn2.Text = "Share Name";
            this.olvColumn2.Width = 177;
            // 
            // frmSearchShareTypes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(478, 394);
            this.Controls.Add(this.objSharetypes);
            this.Name = "frmSearchShareTypes";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Search ShareTypes";
            this.Load += new System.EventHandler(this.frmSearchShareTypes_Load);
            ((System.ComponentModel.ISupportInitialize)(this.objSharetypes)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private BrightIdeasSoftware.ObjectListView objSharetypes;
        private BrightIdeasSoftware.OLVColumn olvColumn1;
        private BrightIdeasSoftware.OLVColumn olvColumn2;
    }
}