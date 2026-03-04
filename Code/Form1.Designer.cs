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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            BtnSelectFile = new Button();
            BtnUnlock = new Button();
            tbPath = new TextBox();
            openFileDialog1 = new OpenFileDialog();
            cbBackup = new CheckBox();
            btnClose = new Button();
            pictureBox1 = new PictureBox();
            label1 = new Label();
            progressBar1 = new ProgressBar();
            statusStrip1 = new StatusStrip();
            lblStatus = new ToolStripStatusLabel();
            btnPin = new Button();
            btnMini = new Button();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // BtnSelectFile
            // 
            BtnSelectFile.Location = new Point(328, 62);
            BtnSelectFile.Name = "BtnSelectFile";
            BtnSelectFile.Size = new Size(75, 23);
            BtnSelectFile.TabIndex = 0;
            BtnSelectFile.Text = "choose File";
            BtnSelectFile.UseVisualStyleBackColor = true;
            BtnSelectFile.Click += BtnSelectFile_Click;
            // 
            // BtnUnlock
            // 
            BtnUnlock.Enabled = false;
            BtnUnlock.Location = new Point(328, 91);
            BtnUnlock.Name = "BtnUnlock";
            BtnUnlock.Size = new Size(75, 23);
            BtnUnlock.TabIndex = 1;
            BtnUnlock.Text = "Unlock";
            BtnUnlock.UseVisualStyleBackColor = true;
            BtnUnlock.Click += BtnUnlock_Click;
            // 
            // tbPath
            // 
            tbPath.Location = new Point(12, 62);
            tbPath.Name = "tbPath";
            tbPath.Size = new Size(310, 23);
            tbPath.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "Office (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx";
            // 
            // cbBackup
            // 
            cbBackup.AutoSize = true;
            cbBackup.Checked = true;
            cbBackup.CheckState = CheckState.Checked;
            cbBackup.Location = new Point(12, 94);
            cbBackup.Name = "cbBackup";
            cbBackup.Size = new Size(65, 19);
            cbBackup.TabIndex = 4;
            cbBackup.Text = "Backup";
            cbBackup.UseVisualStyleBackColor = true;
            // 
            // btnClose
            // 
            btnClose.BackColor = Color.Transparent;
            btnClose.BackgroundImage = Properties.Resources.exit;
            btnClose.BackgroundImageLayout = ImageLayout.Zoom;
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.ForeColor = Color.Transparent;
            btnClose.Location = new Point(384, 5);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(19, 19);
            btnClose.TabIndex = 5;
            btnClose.UseVisualStyleBackColor = false;
            btnClose.Click += btnClose_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Image = Properties.Resources.xdoclogo;
            pictureBox1.Location = new Point(12, 5);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(34, 30);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.TabIndex = 6;
            pictureBox1.TabStop = false;
            pictureBox1.Click += pictureBox1_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 10F);
            label1.Location = new Point(52, 11);
            label1.Name = "label1";
            label1.Size = new Size(95, 19);
            label1.TabIndex = 7;
            label1.Text = "XDocUnlocker";
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(328, 91);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(75, 22);
            progressBar1.TabIndex = 8;
            progressBar1.Visible = false;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblStatus });
            statusStrip1.Location = new Point(0, 128);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(415, 22);
            statusStrip1.SizingGrip = false;
            statusStrip1.Stretch = false;
            statusStrip1.TabIndex = 9;
            statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(241, 17);
            lblStatus.Text = "💡 Drag 'n drop Office files or use the button";
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // btnPin
            // 
            btnPin.BackColor = Color.Transparent;
            btnPin.BackgroundImage = Properties.Resources.unpinned;
            btnPin.BackgroundImageLayout = ImageLayout.Zoom;
            btnPin.FlatAppearance.BorderSize = 0;
            btnPin.FlatStyle = FlatStyle.Flat;
            btnPin.ForeColor = Color.Transparent;
            btnPin.Location = new Point(334, 5);
            btnPin.Name = "btnPin";
            btnPin.Size = new Size(19, 19);
            btnPin.TabIndex = 10;
            btnPin.UseVisualStyleBackColor = false;
            btnPin.Click += btnPin_Click;
            // 
            // btnMini
            // 
            btnMini.BackColor = Color.Transparent;
            btnMini.BackgroundImage = Properties.Resources.mini;
            btnMini.BackgroundImageLayout = ImageLayout.Zoom;
            btnMini.FlatAppearance.BorderSize = 0;
            btnMini.FlatStyle = FlatStyle.Flat;
            btnMini.ForeColor = Color.Transparent;
            btnMini.Location = new Point(359, 5);
            btnMini.Name = "btnMini";
            btnMini.Size = new Size(19, 19);
            btnMini.TabIndex = 11;
            btnMini.UseVisualStyleBackColor = false;
            btnMini.Click += btnMini_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ClientSize = new Size(415, 150);
            Controls.Add(btnMini);
            Controls.Add(btnPin);
            Controls.Add(statusStrip1);
            Controls.Add(progressBar1);
            Controls.Add(label1);
            Controls.Add(pictureBox1);
            Controls.Add(btnClose);
            Controls.Add(cbBackup);
            Controls.Add(tbPath);
            Controls.Add(BtnUnlock);
            Controls.Add(BtnSelectFile);
            FormBorderStyle = FormBorderStyle.None;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "XDocUnlocker";
            MouseDown += Form1_MouseDown;
            MouseMove += Form1_MouseMove;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button BtnSelectFile;
        private Button BtnUnlock;
        private TextBox tbPath;
        private OpenFileDialog openFileDialog1;
        private CheckBox cbBackup;
        private Button btnClose;
        private PictureBox pictureBox1;
        private Label label1;
        private ProgressBar progressBar1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lblStatus;
        private Button btnPin;
        private Button btnMini;
    }
}
