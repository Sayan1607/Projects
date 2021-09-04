namespace ProjectHealthApplication
{
    partial class ChangePassword
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.txtRptChngdPasswd = new System.Windows.Forms.TextBox();
            this.lblRptNewPasswd = new System.Windows.Forms.Label();
            this.txtChngdPaswwd = new System.Windows.Forms.TextBox();
            this.txtOldPasswd = new System.Windows.Forms.TextBox();
            this.lblEnterNewPasswd = new System.Windows.Forms.Label();
            this.lblEnterOldPasswd = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSubmit);
            this.panel1.Controls.Add(this.txtRptChngdPasswd);
            this.panel1.Controls.Add(this.lblRptNewPasswd);
            this.panel1.Controls.Add(this.txtChngdPaswwd);
            this.panel1.Controls.Add(this.txtOldPasswd);
            this.panel1.Controls.Add(this.lblEnterNewPasswd);
            this.panel1.Controls.Add(this.lblEnterOldPasswd);
            this.panel1.Location = new System.Drawing.Point(1, 13);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(477, 235);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // btnSubmit
            // 
            this.btnSubmit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubmit.Location = new System.Drawing.Point(180, 174);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(138, 51);
            this.btnSubmit.TabIndex = 6;
            this.btnSubmit.Text = "Submit";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // txtRptChngdPasswd
            // 
            this.txtRptChngdPasswd.Location = new System.Drawing.Point(222, 125);
            this.txtRptChngdPasswd.Multiline = true;
            this.txtRptChngdPasswd.Name = "txtRptChngdPasswd";
            this.txtRptChngdPasswd.PasswordChar = '*';
            this.txtRptChngdPasswd.Size = new System.Drawing.Size(239, 31);
            this.txtRptChngdPasswd.TabIndex = 5;
            // 
            // lblRptNewPasswd
            // 
            this.lblRptNewPasswd.AutoSize = true;
            this.lblRptNewPasswd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRptNewPasswd.Location = new System.Drawing.Point(16, 128);
            this.lblRptNewPasswd.Name = "lblRptNewPasswd";
            this.lblRptNewPasswd.Size = new System.Drawing.Size(200, 20);
            this.lblRptNewPasswd.TabIndex = 4;
            this.lblRptNewPasswd.Text = "Reenter New Password:";
            // 
            // txtChngdPaswwd
            // 
            this.txtChngdPaswwd.Location = new System.Drawing.Point(222, 77);
            this.txtChngdPaswwd.Multiline = true;
            this.txtChngdPaswwd.Name = "txtChngdPaswwd";
            this.txtChngdPaswwd.PasswordChar = '*';
            this.txtChngdPaswwd.Size = new System.Drawing.Size(239, 31);
            this.txtChngdPaswwd.TabIndex = 3;
            // 
            // txtOldPasswd
            // 
            this.txtOldPasswd.Location = new System.Drawing.Point(222, 30);
            this.txtOldPasswd.Multiline = true;
            this.txtOldPasswd.Name = "txtOldPasswd";
            this.txtOldPasswd.PasswordChar = '*';
            this.txtOldPasswd.Size = new System.Drawing.Size(239, 31);
            this.txtOldPasswd.TabIndex = 2;
            // 
            // lblEnterNewPasswd
            // 
            this.lblEnterNewPasswd.AutoSize = true;
            this.lblEnterNewPasswd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEnterNewPasswd.Location = new System.Drawing.Point(16, 77);
            this.lblEnterNewPasswd.Name = "lblEnterNewPasswd";
            this.lblEnterNewPasswd.Size = new System.Drawing.Size(179, 20);
            this.lblEnterNewPasswd.TabIndex = 1;
            this.lblEnterNewPasswd.Text = "Enter New Password:";
            // 
            // lblEnterOldPasswd
            // 
            this.lblEnterOldPasswd.AutoSize = true;
            this.lblEnterOldPasswd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEnterOldPasswd.Location = new System.Drawing.Point(12, 30);
            this.lblEnterOldPasswd.Name = "lblEnterOldPasswd";
            this.lblEnterOldPasswd.Size = new System.Drawing.Size(172, 20);
            this.lblEnterOldPasswd.TabIndex = 0;
            this.lblEnterOldPasswd.Text = "Enter Old Password:";
            // 
            // ChangePassword
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 250);
            this.Controls.Add(this.panel1);
            this.Name = "ChangePassword";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ChangePassword";
            this.Load += new System.EventHandler(this.ChangePassword_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblEnterOldPasswd;
        private System.Windows.Forms.Label lblEnterNewPasswd;
        private System.Windows.Forms.TextBox txtOldPasswd;
        private System.Windows.Forms.TextBox txtChngdPaswwd;
        private System.Windows.Forms.TextBox txtRptChngdPasswd;
        private System.Windows.Forms.Label lblRptNewPasswd;
        private System.Windows.Forms.Button btnSubmit;
    }
}