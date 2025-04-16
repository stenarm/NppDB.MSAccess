using System;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    public partial class FrmPassword : Form
    {
        public FrmPassword()
        {
            InitializeComponent();
        }

        public bool VisiblePassword
        {
            get => cbxShowPwd.Checked;
            set => cbxShowPwd.Checked = value;
        }

        public string Password => txtPwd.Text.Trim();

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cbxShowPwd_CheckedChanged(object sender, EventArgs e)
        {
            AdjustPasswordChar();
        }

        private void AdjustPasswordChar()
        {
            txtPwd.PasswordChar = cbxShowPwd.Checked ? (char)0 : '*';
        }

        private void frmPassword_Load(object sender, EventArgs e)
        {
            AdjustPasswordChar();
        }

    }
}
