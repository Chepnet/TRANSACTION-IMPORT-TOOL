
namespace ReadExcel
{
    partial class frmSearchPayModes
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
            this.objListPayMode = new BrightIdeasSoftware.ObjectListView();
            this.olvColumn1 = ((BrightIdeasSoftware.OLVColumn)(new BrightIdeasSoftware.OLVColumn()));
            ((System.ComponentModel.ISupportInitialize)(this.objListPayMode)).BeginInit();
            this.SuspendLayout();
            // 
            // objListPayMode
            // 
            this.objListPayMode.AllColumns.Add(this.olvColumn1);
            this.objListPayMode.CellEditUseWholeCell = false;
            this.objListPayMode.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.olvColumn1});
            this.objListPayMode.Cursor = System.Windows.Forms.Cursors.Default;
            this.objListPayMode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objListPayMode.FullRowSelect = true;
            this.objListPayMode.GridLines = true;
            this.objListPayMode.HideSelection = false;
            this.objListPayMode.Location = new System.Drawing.Point(0, 0);
            this.objListPayMode.Name = "objListPayMode";
            this.objListPayMode.ShowGroups = false;
            this.objListPayMode.Size = new System.Drawing.Size(330, 436);
            this.objListPayMode.TabIndex = 3;
            this.objListPayMode.UseCompatibleStateImageBehavior = false;
            this.objListPayMode.View = System.Windows.Forms.View.Details;
            this.objListPayMode.SelectedIndexChanged += new System.EventHandler(this.objListPayMode_SelectedIndexChanged);
            this.objListPayMode.DoubleClick += new System.EventHandler(this.objListPayMode_DoubleClick);
            // 
            // olvColumn1
            // 
            this.olvColumn1.AspectName = "PaymentModeName";
            this.olvColumn1.Text = "Payment Mode";
            this.olvColumn1.Width = 282;
            // 
            // frmSearchPayModes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(330, 436);
            this.Controls.Add(this.objListPayMode);
            this.Name = "frmSearchPayModes";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Search PayModes";
            this.Load += new System.EventHandler(this.frmSearchPayModes_Load);
            ((System.ComponentModel.ISupportInitialize)(this.objListPayMode)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private BrightIdeasSoftware.ObjectListView objListPayMode;
        private BrightIdeasSoftware.OLVColumn olvColumn1;
    }
}