using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    public partial class frmPassword : Form
    {
        public frmPassword()
        {
            InitializeComponent();
        }

        public bool VisiblePassword
        {
            get { return this.cbxShowPwd.Checked; }
            set { this.cbxShowPwd.Checked = value; }
        }

        public string Password
        {
            get { return this.txtPwd.Text.Trim(); }
            set { this.txtPwd.Text = value; }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void cbxShowPwd_CheckedChanged(object sender, EventArgs e)
        {
            AdjustPasswordChar();
        }

        private void AdjustPasswordChar()
        {
            this.txtPwd.PasswordChar = this.cbxShowPwd.Checked ? (char)0 : '*';
        }

        private void frmPassword_Load(object sender, EventArgs e)
        {
            AdjustPasswordChar();
        }

    }
}
