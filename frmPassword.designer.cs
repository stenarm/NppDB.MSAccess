using System.ComponentModel;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    partial class FrmPassword
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtPwd = new System.Windows.Forms.TextBox();
            this.cbxShowPwd = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(120, 113);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(64, 36);
            this.btnClose.TabIndex = 8;
            this.btnClose.Text = "Cancel";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(43, 113);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(64, 36);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtPwd
            // 
            this.txtPwd.Location = new System.Drawing.Point(43, 39);
            this.txtPwd.Name = "txtPwd";
            this.txtPwd.Size = new System.Drawing.Size(141, 20);
            this.txtPwd.TabIndex = 9;
            // 
            // cbxShowPwd
            // 
            this.cbxShowPwd.AutoSize = true;
            this.cbxShowPwd.Location = new System.Drawing.Point(63, 78);
            this.cbxShowPwd.Name = "cbxShowPwd";
            this.cbxShowPwd.Size = new System.Drawing.Size(99, 17);
            this.cbxShowPwd.TabIndex = 10;
            this.cbxShowPwd.Text = "show password";
            this.cbxShowPwd.UseVisualStyleBackColor = true;
            this.cbxShowPwd.CheckedChanged += new System.EventHandler(this.cbxShowPwd_CheckedChanged);
            // 
            // frmPassword
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(232, 191);
            this.ControlBox = false;
            this.Controls.Add(this.cbxShowPwd);
            this.Controls.Add(this.txtPwd);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmPassword";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Password";
            this.Load += new System.EventHandler(this.frmPassword_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnClose;
        private Button btnOK;
        private TextBox txtPwd;
        private CheckBox cbxShowPwd;
    }
}