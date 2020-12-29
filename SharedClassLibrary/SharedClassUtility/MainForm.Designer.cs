
namespace SharedClassUtility
{
    partial class MainForm
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
            this.tabcGeneral = new System.Windows.Forms.TabControl();
            this.tabHome = new System.Windows.Forms.TabPage();
            this.tabEncryption = new System.Windows.Forms.TabPage();
            this.btnEncrypt = new System.Windows.Forms.Button();
            this.tbText = new System.Windows.Forms.TextBox();
            this.tbPassPhrase = new System.Windows.Forms.TextBox();
            this.lblEncryptText = new System.Windows.Forms.Label();
            this.lblPassPhrase = new System.Windows.Forms.Label();
            this.tbResult = new System.Windows.Forms.TextBox();
            this.btnDecrypt = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.tabcGeneral.SuspendLayout();
            this.tabEncryption.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabcGeneral
            // 
            this.tabcGeneral.Controls.Add(this.tabHome);
            this.tabcGeneral.Controls.Add(this.tabEncryption);
            this.tabcGeneral.Location = new System.Drawing.Point(12, 12);
            this.tabcGeneral.Name = "tabcGeneral";
            this.tabcGeneral.SelectedIndex = 0;
            this.tabcGeneral.Size = new System.Drawing.Size(333, 228);
            this.tabcGeneral.TabIndex = 0;
            // 
            // tabHome
            // 
            this.tabHome.Location = new System.Drawing.Point(4, 22);
            this.tabHome.Name = "tabHome";
            this.tabHome.Padding = new System.Windows.Forms.Padding(3);
            this.tabHome.Size = new System.Drawing.Size(325, 202);
            this.tabHome.TabIndex = 0;
            this.tabHome.Text = "Home";
            this.tabHome.UseVisualStyleBackColor = true;
            // 
            // tabEncryption
            // 
            this.tabEncryption.Controls.Add(this.lblResult);
            this.tabEncryption.Controls.Add(this.btnDecrypt);
            this.tabEncryption.Controls.Add(this.tbResult);
            this.tabEncryption.Controls.Add(this.lblPassPhrase);
            this.tabEncryption.Controls.Add(this.lblEncryptText);
            this.tabEncryption.Controls.Add(this.tbPassPhrase);
            this.tabEncryption.Controls.Add(this.tbText);
            this.tabEncryption.Controls.Add(this.btnEncrypt);
            this.tabEncryption.Location = new System.Drawing.Point(4, 22);
            this.tabEncryption.Name = "tabEncryption";
            this.tabEncryption.Padding = new System.Windows.Forms.Padding(3);
            this.tabEncryption.Size = new System.Drawing.Size(325, 202);
            this.tabEncryption.TabIndex = 1;
            this.tabEncryption.Text = "Encryption";
            this.tabEncryption.UseVisualStyleBackColor = true;
            // 
            // btnEncrypt
            // 
            this.btnEncrypt.Location = new System.Drawing.Point(10, 96);
            this.btnEncrypt.Name = "btnEncrypt";
            this.btnEncrypt.Size = new System.Drawing.Size(86, 23);
            this.btnEncrypt.TabIndex = 0;
            this.btnEncrypt.Text = "&Encrypt";
            this.btnEncrypt.UseVisualStyleBackColor = true;
            this.btnEncrypt.Click += new System.EventHandler(this.btnEncrypt_Click);
            // 
            // tbText
            // 
            this.tbText.Location = new System.Drawing.Point(10, 27);
            this.tbText.Name = "tbText";
            this.tbText.Size = new System.Drawing.Size(189, 20);
            this.tbText.TabIndex = 1;
            // 
            // tbPassPhrase
            // 
            this.tbPassPhrase.Location = new System.Drawing.Point(10, 70);
            this.tbPassPhrase.Name = "tbPassPhrase";
            this.tbPassPhrase.Size = new System.Drawing.Size(189, 20);
            this.tbPassPhrase.TabIndex = 2;
            // 
            // lblEncryptText
            // 
            this.lblEncryptText.AutoSize = true;
            this.lblEncryptText.Location = new System.Drawing.Point(7, 11);
            this.lblEncryptText.Name = "lblEncryptText";
            this.lblEncryptText.Size = new System.Drawing.Size(125, 13);
            this.lblEncryptText.TabIndex = 3;
            this.lblEncryptText.Text = "Text To Encrypt/Decrypt";
            // 
            // lblPassPhrase
            // 
            this.lblPassPhrase.AutoSize = true;
            this.lblPassPhrase.Location = new System.Drawing.Point(10, 54);
            this.lblPassPhrase.Name = "lblPassPhrase";
            this.lblPassPhrase.Size = new System.Drawing.Size(66, 13);
            this.lblPassPhrase.TabIndex = 4;
            this.lblPassPhrase.Text = "Pass Phrase";
            // 
            // tbResult
            // 
            this.tbResult.Location = new System.Drawing.Point(10, 145);
            this.tbResult.Name = "tbResult";
            this.tbResult.Size = new System.Drawing.Size(189, 20);
            this.tbResult.TabIndex = 5;
            // 
            // btnDecrypt
            // 
            this.btnDecrypt.Location = new System.Drawing.Point(118, 97);
            this.btnDecrypt.Name = "btnDecrypt";
            this.btnDecrypt.Size = new System.Drawing.Size(81, 23);
            this.btnDecrypt.TabIndex = 6;
            this.btnDecrypt.Text = "&Decrypt";
            this.btnDecrypt.UseVisualStyleBackColor = true;
            this.btnDecrypt.Click += new System.EventHandler(this.btnDecrypt_Click);
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(10, 126);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(37, 13);
            this.lblResult.TabIndex = 7;
            this.lblResult.Text = "Result";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(353, 251);
            this.Controls.Add(this.tabcGeneral);
            this.Name = "MainForm";
            this.Text = "SharedClass Utility Software";
            this.tabcGeneral.ResumeLayout(false);
            this.tabEncryption.ResumeLayout(false);
            this.tabEncryption.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabcGeneral;
        private System.Windows.Forms.TabPage tabHome;
        private System.Windows.Forms.TabPage tabEncryption;
        private System.Windows.Forms.Button btnDecrypt;
        private System.Windows.Forms.TextBox tbResult;
        private System.Windows.Forms.Label lblPassPhrase;
        private System.Windows.Forms.Label lblEncryptText;
        private System.Windows.Forms.TextBox tbPassPhrase;
        private System.Windows.Forms.TextBox tbText;
        private System.Windows.Forms.Button btnEncrypt;
        private System.Windows.Forms.Label lblResult;
    }
}

