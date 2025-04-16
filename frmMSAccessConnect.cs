using System;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    public partial class FrmMsAccessConnect : Form
    {
        public FrmMsAccessConnect()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        public bool IsNew => rdoNew.Checked;

        private void btnOK_Click(object sender, EventArgs e)
        {
            
            DialogResult = DialogResult.OK;
            
        }

    }
}
