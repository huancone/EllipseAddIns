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
            this.gbEllipseSettings = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnDeleteLocalEllipseSettings = new System.Windows.Forms.Button();
            this.btnGenerateDefault = new System.Windows.Forms.Button();
            this.btnGenerateFromNetwork = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbServiceFileNetworkUrl = new System.Windows.Forms.TextBox();
            this.cbForceRegionConfig = new System.Windows.Forms.CheckBox();
            this.gbDebugging.SuspendLayout();
            this.gbGlobalSettings.SuspendLayout();
            this.gbEllipseSettings.SuspendLayout();
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
            this.okButton.Location = new System.Drawing.Point(229, 380);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 26);
            this.okButton.TabIndex = 30;
            this.okButton.Text = "&OK";
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
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
            this.tbLocalDataPath.Size = new System.Drawing.Size(263, 20);
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
            // gbEllipseSettings
            // 
            this.gbEllipseSettings.Controls.Add(this.cbForceRegionConfig);
            this.gbEllipseSettings.Controls.Add(this.label2);
            this.gbEllipseSettings.Controls.Add(this.btnDeleteLocalEllipseSettings);
            this.gbEllipseSettings.Controls.Add(this.btnGenerateDefault);
            this.gbEllipseSettings.Controls.Add(this.btnGenerateFromNetwork);
            this.gbEllipseSettings.Controls.Add(this.label1);
            this.gbEllipseSettings.Controls.Add(this.tbServiceFileNetworkUrl);
            this.gbEllipseSettings.Location = new System.Drawing.Point(12, 205);
            this.gbEllipseSettings.Name = "gbEllipseSettings";
            this.gbEllipseSettings.Size = new System.Drawing.Size(296, 142);
            this.gbEllipseSettings.TabIndex = 39;
            this.gbEllipseSettings.TabStop = false;
            this.gbEllipseSettings.Text = "Configuración Ellipse:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Generar Archivo Local:";
            // 
            // btnDeleteLocalEllipseSettings
            // 
            this.btnDeleteLocalEllipseSettings.Location = new System.Drawing.Point(186, 84);
            this.btnDeleteLocalEllipseSettings.Name = "btnDeleteLocalEllipseSettings";
            this.btnDeleteLocalEllipseSettings.Size = new System.Drawing.Size(87, 23);
            this.btnDeleteLocalEllipseSettings.TabIndex = 4;
            this.btnDeleteLocalEllipseSettings.Text = "&Eliminar Local";
            this.btnDeleteLocalEllipseSettings.UseVisualStyleBackColor = true;
            this.btnDeleteLocalEllipseSettings.Click += new System.EventHandler(this.btnDeleteLocalEllipseSettings_Click);
            // 
            // btnGenerateDefault
            // 
            this.btnGenerateDefault.Location = new System.Drawing.Point(87, 84);
            this.btnGenerateDefault.Name = "btnGenerateDefault";
            this.btnGenerateDefault.Size = new System.Drawing.Size(93, 23);
            this.btnGenerateDefault.TabIndex = 3;
            this.btnGenerateDefault.Text = "&Predeterminado";
            this.btnGenerateDefault.UseVisualStyleBackColor = true;
            this.btnGenerateDefault.Click += new System.EventHandler(this.btnGenerateDefault_Click);
            // 
            // btnGenerateFromNetwork
            // 
            this.btnGenerateFromNetwork.Location = new System.Drawing.Point(6, 84);
            this.btnGenerateFromNetwork.Name = "btnGenerateFromNetwork";
            this.btnGenerateFromNetwork.Size = new System.Drawing.Size(75, 23);
            this.btnGenerateFromNetwork.TabIndex = 2;
            this.btnGenerateFromNetwork.Text = "Desde &Url";
            this.btnGenerateFromNetwork.UseVisualStyleBackColor = true;
            this.btnGenerateFromNetwork.Click += new System.EventHandler(this.btnGenerateFromNetwork_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(174, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Ruta de Información de Servidores:";
            // 
            // tbServiceFileNetworkUrl
            // 
            this.tbServiceFileNetworkUrl.Location = new System.Drawing.Point(6, 36);
            this.tbServiceFileNetworkUrl.Name = "tbServiceFileNetworkUrl";
            this.tbServiceFileNetworkUrl.Size = new System.Drawing.Size(266, 20);
            this.tbServiceFileNetworkUrl.TabIndex = 0;
            // 
            // cbForceRegionConfig
            // 
            this.cbForceRegionConfig.AutoSize = true;
            this.cbForceRegionConfig.Location = new System.Drawing.Point(9, 114);
            this.cbForceRegionConfig.Name = "cbForceRegionConfig";
            this.cbForceRegionConfig.Size = new System.Drawing.Size(166, 17);
            this.cbForceRegionConfig.TabIndex = 6;
            this.cbForceRegionConfig.Text = "Forzar &Regionalización en-US";
            this.cbForceRegionConfig.UseVisualStyleBackColor = true;
            this.cbForceRegionConfig.CheckedChanged += new System.EventHandler(this.cbForceRegionConfig_CheckedChanged);
            // 
            // SettingsBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(316, 418);
            this.Controls.Add(this.gbEllipseSettings);
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
            this.gbEllipseSettings.ResumeLayout(false);
            this.gbEllipseSettings.PerformLayout();
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
        private System.Windows.Forms.GroupBox gbEllipseSettings;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnDeleteLocalEllipseSettings;
        private System.Windows.Forms.Button btnGenerateDefault;
        private System.Windows.Forms.Button btnGenerateFromNetwork;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbServiceFileNetworkUrl;
        private System.Windows.Forms.CheckBox cbForceRegionConfig;

    }
}
