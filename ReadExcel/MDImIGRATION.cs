using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcel
{
    public partial class MDImIGRATION : Form
    {
        private int childFormNumber = 0;

        public MDImIGRATION()
        {
            InitializeComponent();
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            Form childForm = new Form();
            childForm.MdiParent = this;
            childForm.Text = "Window " + childFormNumber++;
            childForm.Show();
        }

        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = openFileDialog.FileName;
            }
        }

        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
            }
        }

        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //statusStrip.Visible = statusBarToolStripMenuItem.Checked;
        }

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void mowascoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmMowaso frm = new ReadExcel.frmMowaso();
            frm.Show();
        }

        private void mowascoSheet2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void sundataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //frmSunData frm = new ReadExcel.frmSunData();
            //frm.Show();
        }

        private void transactionsMigrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmEquityTransactions2018 frm = new ReadExcel.frmEquityTransactions2018();
            frm.ShowDialog();
        }

        private void barabaraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmImportLoans frm = new ReadExcel.frmImportLoans();
            frm.ShowDialog();
        }

        private void migrateMembersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmBarabaramemberMigration frm = new frmBarabaramemberMigration();
            frm.ShowDialog();
        }

        private void membersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmNascaMigration frm = new frmNascaMigration();
            frm.ShowDialog();
        }

        private void membersToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmMigratemembersKRB frm = new frmMigratemembersKRB();
            frm.ShowDialog();
        }

        private void menuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void addProductToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmKRBImportationTool frm = new frmKRBImportationTool();
            frm.ShowDialog();
        }

        private void migrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSearchImportationOrder frm = new frmSearchImportationOrder();
            frm.ShowDialog();
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmImportFileFormat frm = new frmImportFileFormat();
            frm.ShowDialog();
        }

        private void createFileNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmTransationImportFileName frm = new frmTransationImportFileName();
            frm.ShowDialog();
        }
    }
}
