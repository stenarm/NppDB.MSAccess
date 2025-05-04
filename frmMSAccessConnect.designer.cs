using System.ComponentModel;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    partial class FrmMsAccessConnect
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
            this.rdoOpen = new System.Windows.Forms.RadioButton();
            this.rdoNew = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // btnClose
            // 
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
            // rdoOpen
            // 
            this.rdoOpen.AutoSize = true;
            this.rdoOpen.Checked = true;
            this.rdoOpen.Location = new System.Drawing.Point(43, 50);
            this.rdoOpen.Name = "rdoOpen";
            this.rdoOpen.Size = new System.Drawing.Size(136, 17);
            this.rdoOpen.TabIndex = 1;
            this.rdoOpen.TabStop = true;
            this.rdoOpen.Text = "Open an existing database";
            this.rdoOpen.UseVisualStyleBackColor = true;
            // 
            // rdoNew
            // 
            this.rdoNew.AutoSize = true;
            this.rdoNew.Location = new System.Drawing.Point(43, 27);
            this.rdoNew.Name = "rdoNew";
            this.rdoNew.Size = new System.Drawing.Size(137, 17);
            this.rdoNew.TabIndex = 0;
            this.rdoNew.Text = "Create an empty database ";
            this.rdoNew.UseVisualStyleBackColor = true;
            // 
            // frmMSAccessConnect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(232, 191);
            this.ControlBox = false;
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.rdoOpen);
            this.Controls.Add(this.rdoNew);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmMsAccessConnect";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MSAccess Connect";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Button btnClose;
        private Button btnOK;
        private RadioButton rdoOpen;
        private RadioButton rdoNew;
    }
}