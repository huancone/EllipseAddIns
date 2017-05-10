namespace EllipseCommonsClassLibrary
{
    partial class SettingsBox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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
            this.labelProductName = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.gbDebugging = new System.Windows.Forms.GroupBox();
            this.cbDebugErrors = new System.Windows.Forms.CheckBox();
            this.cbDebugWarnings = new System.Windows.Forms.CheckBox();
            this.cbDebugQueries = new System.Windows.Forms.CheckBox();
            this.tbLocalDataPath = new System.Windows.Forms.TextBox();
            this.lblLocalDataPath = new System.Windows.Forms.Label();
            this.gbGlobalSettings = new System.Windows.Forms.GroupBox();
            this.gbDebugging.SuspendLayout();
            this.gbGlobalSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelProductName
            // 
            this.labelProductName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelProductName.Location = new System.Drawing.Point(9, 9);
            this.labelProductName.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.labelProductName.MaximumSize = new System.Drawing.Size(0, 17);
            this.labelProductName.Name = "labelProductName";
            this.labelProductName.Size = new System.Drawing.Size(298, 17);
            this.labelProductName.TabIndex = 20;
            this.labelProductName.Text = "Product Name";
            this.labelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Location = new System.Drawing.Point(229, 210);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 26);
            this.okButton.TabIndex = 30;
            this.okButton.Text = "&OK";
            // 
            // gbDebugging
            // 
            this.gbDebugging.Controls.Add(this.cbDebugErrors);
            this.gbDebugging.Controls.Add(this.cbDebugWarnings);
            this.gbDebugging.Controls.Add(this.cbDebugQueries);
            this.gbDebugging.Location = new System.Drawing.Point(12, 98);
            this.gbDebugging.Name = "gbDebugging";
            this.gbDebugging.Size = new System.Drawing.Size(296, 100);
            this.gbDebugging.TabIndex = 37;
            this.gbDebugging.TabStop = false;
            this.gbDebugging.Text = "Debugging";
            // 
            // cbDebugErrors
            // 
            this.cbDebugErrors.AutoSize = true;
            this.cbDebugErrors.Location = new System.Drawing.Point(6, 23);
            this.cbDebugErrors.Name = "cbDebugErrors";
            this.cbDebugErrors.Size = new System.Drawing.Size(126, 17);
            this.cbDebugErrors.TabIndex = 37;
            this.cbDebugErrors.Text = "Debug Errores/Fallas";
            this.cbDebugErrors.UseVisualStyleBackColor = true;
            this.cbDebugErrors.CheckedChanged += new System.EventHandler(this.cbDebugErrors_CheckedChanged);
            // 
            // cbDebugWarnings
            // 
            this.cbDebugWarnings.AutoSize = true;
            this.cbDebugWarnings.Location = new System.Drawing.Point(6, 46);
            this.cbDebugWarnings.Name = "cbDebugWarnings";
            this.cbDebugWarnings.Size = new System.Drawing.Size(143, 17);
            this.cbDebugWarnings.TabIndex = 38;
            this.cbDebugWarnings.Text = "Debug Warnings/Alertas";
            this.cbDebugWarnings.UseVisualStyleBackColor = true;
            this.cbDebugWarnings.CheckedChanged += new System.EventHandler(this.cbDebugWarnings_CheckedChanged);
            // 
            // cbDebugQueries
            // 
            this.cbDebugQueries.AutoSize = true;
            this.cbDebugQueries.Location = new System.Drawing.Point(6, 69);
            this.cbDebugQueries.Name = "cbDebugQueries";
            this.cbDebugQueries.Size = new System.Drawing.Size(148, 17);
            this.cbDebugQueries.TabIndex = 39;
            this.cbDebugQueries.Text = "Debug Queries/Consultas";
            this.cbDebugQueries.UseVisualStyleBackColor = true;
            this.cbDebugQueries.CheckedChanged += new System.EventHandler(this.cbDebugQueries_CheckedChanged);
            // 
            // tbLocalDataPath
            // 
            this.tbLocalDataPath.Location = new System.Drawing.Point(9, 32);
            this.tbLocalDataPath.Name = "tbLocalDataPath";
            this.tbLocalDataPath.Size = new System.Drawing.Size(251, 20);
            this.tbLocalDataPath.TabIndex = 41;
            this.tbLocalDataPath.TextChanged += new System.EventHandler(this.tbLocalDataPath_TextChanged);
            // 
            // lblLocalDataPath
            // 
            this.lblLocalDataPath.AutoSize = true;
            this.lblLocalDataPath.Location = new System.Drawing.Point(6, 16);
            this.lblLocalDataPath.Name = "lblLocalDataPath";
            this.lblLocalDataPath.Size = new System.Drawing.Size(103, 13);
            this.lblLocalDataPath.TabIndex = 40;
            this.lblLocalDataPath.Text = "Ruta de Data Local:";
            // 
            // gbGlobalSettings
            // 
            this.gbGlobalSettings.Controls.Add(this.tbLocalDataPath);
            this.gbGlobalSettings.Controls.Add(this.lblLocalDataPath);
            this.gbGlobalSettings.Location = new System.Drawing.Point(12, 31);
            this.gbGlobalSettings.Name = "gbGlobalSettings";
            this.gbGlobalSettings.Size = new System.Drawing.Size(296, 61);
            this.gbGlobalSettings.TabIndex = 38;
            this.gbGlobalSettings.TabStop = false;
            this.gbGlobalSettings.Text = "Configuración Global:";
            // 
            // SettingsBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(316, 248);
            this.Controls.Add(this.gbGlobalSettings);
            this.Controls.Add(this.gbDebugging);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.labelProductName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsBox";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "SettingsBox";
            this.Load += new System.EventHandler(this.SettingsBox_Load);
            this.gbDebugging.ResumeLayout(false);
            this.gbDebugging.PerformLayout();
            this.gbGlobalSettings.ResumeLayout(false);
            this.gbGlobalSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label labelProductName;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.GroupBox gbDebugging;
        private System.Windows.Forms.TextBox tbLocalDataPath;
        private System.Windows.Forms.CheckBox cbDebugErrors;
        private System.Windows.Forms.CheckBox cbDebugWarnings;
        private System.Windows.Forms.CheckBox cbDebugQueries;
        private System.Windows.Forms.Label lblLocalDataPath;
        private System.Windows.Forms.GroupBox gbGlobalSettings;

    }
}
