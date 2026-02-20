namespace XDocUnlocker
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            BtnSelectFile = new Button();
            BtnUnlock = new Button();
            tbPath = new TextBox();
            openFileDialog1 = new OpenFileDialog();
            lblStatus = new Label();
            cbBackup = new CheckBox();
            SuspendLayout();
            // 
            // BtnSelectFile
            // 
            BtnSelectFile.Location = new Point(331, 37);
            BtnSelectFile.Name = "BtnSelectFile";
            BtnSelectFile.Size = new Size(75, 23);
            BtnSelectFile.TabIndex = 0;
            BtnSelectFile.Text = "...";
            BtnSelectFile.UseVisualStyleBackColor = true;
            BtnSelectFile.Click += BtnSelectFile_Click;
            // 
            // BtnUnlock
            // 
            BtnUnlock.Enabled = false;
            BtnUnlock.Location = new Point(331, 66);
            BtnUnlock.Name = "BtnUnlock";
            BtnUnlock.Size = new Size(75, 23);
            BtnUnlock.TabIndex = 1;
            BtnUnlock.Text = "Unlock";
            BtnUnlock.UseVisualStyleBackColor = true;
            BtnUnlock.Click += BtnUnlock_Click;
            // 
            // tbPath
            // 
            tbPath.Location = new Point(12, 37);
            tbPath.Name = "tbPath";
            tbPath.Size = new Size(313, 23);
            tbPath.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "Office (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx";
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Consolas", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblStatus.Location = new Point(12, 102);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(25, 13);
            lblStatus.TabIndex = 3;
            lblStatus.Text = "...";
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cbBackup
            // 
            cbBackup.AutoSize = true;
            cbBackup.Checked = true;
            cbBackup.CheckState = CheckState.Checked;
            cbBackup.Location = new Point(12, 69);
            cbBackup.Name = "cbBackup";
            cbBackup.Size = new Size(65, 19);
            cbBackup.TabIndex = 4;
            cbBackup.Text = "Backup";
            cbBackup.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(414, 121);
            Controls.Add(cbBackup);
            Controls.Add(lblStatus);
            Controls.Add(tbPath);
            Controls.Add(BtnUnlock);
            Controls.Add(BtnSelectFile);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "XDocUnlocker";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button BtnSelectFile;
        private Button BtnUnlock;
        private TextBox tbPath;
        private OpenFileDialog openFileDialog1;
        private Label lblStatus;
        private CheckBox cbBackup;
    }
}
