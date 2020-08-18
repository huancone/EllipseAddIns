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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsBox));
            this.labelProductName = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.gbDebugging = new System.Windows.Forms.GroupBox();
            this.cbDebugErrors = new System.Windows.Forms.CheckBox();
            this.cbDebugWarnings = new System.Windows.Forms.CheckBox();
            this.cbDebugQueries = new System.Windows.Forms.CheckBox();
            this.tbLocalDataPath = new System.Windows.Forms.TextBox();
            this.lblLocalDataPath = new System.Windows.Forms.Label();
            this.gbGlobalSettings = new System.Windows.Forms.GroupBox();
            this.btnOpenLocalPath = new System.Windows.Forms.Button();
            this.btnRestoreLocalPath = new System.Windows.Forms.Button();
            this.cbForceRegionConfig = new System.Windows.Forms.CheckBox();
            this.gbEllipseSettings = new System.Windows.Forms.GroupBox();
            this.btnOpenServicesPath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRestoreServiceFile = new System.Windows.Forms.Button();
            this.btnGenerateServiceFileFromDefault = new System.Windows.Forms.Button();
            this.btnGenerateServiceFileFromUrl = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbServiceFileNetworkUrl = new System.Windows.Forms.TextBox();
            this.cbForceServerList = new System.Windows.Forms.CheckBox();
            this.btnOpenTnsnamesPath = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.tbTnsNameUrl = new System.Windows.Forms.TextBox();
            this.btnGenerateTnsnamesFile = new System.Windows.Forms.Button();
            this.gbDatabaseSettings = new System.Windows.Forms.GroupBox();
            this.btnRestoreTnsnamesUrl = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnDeleteCustomDb = new System.Windows.Forms.Button();
            this.btnGenerateCustomDb = new System.Windows.Forms.Button();
            this.ttSettings = new System.Windows.Forms.ToolTip();
            this.gbServiceDatabaseRelation = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.gbDebugging.SuspendLayout();
            this.gbGlobalSettings.SuspendLayout();
            this.gbEllipseSettings.SuspendLayout();
            this.gbDatabaseSettings.SuspendLayout();
            this.gbServiceDatabaseRelation.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelProductName
            // 
            this.labelProductName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelProductName.Location = new System.Drawing.Point(9, 9);
            this.labelProductName.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.labelProductName.MaximumSize = new System.Drawing.Size(0, 17);
            this.labelProductName.Name = "labelProductName";
            this.labelProductName.Size = new System.Drawing.Size(321, 17);
            this.labelProductName.TabIndex = 20;
            this.labelProductName.Text = "Product Name";
            this.labelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Location = new System.Drawing.Point(255, 588);
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
            this.gbDebugging.Location = new System.Drawing.Point(12, 456);
            this.gbDebugging.Name = "gbDebugging";
            this.gbDebugging.Size = new System.Drawing.Size(318, 103);
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
            // 
            // tbLocalDataPath
            // 
            this.tbLocalDataPath.Location = new System.Drawing.Point(9, 32);
            this.tbLocalDataPath.Name = "tbLocalDataPath";
            this.tbLocalDataPath.Size = new System.Drawing.Size(263, 20);
            this.tbLocalDataPath.TabIndex = 41;
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
            this.gbGlobalSettings.Controls.Add(this.btnOpenLocalPath);
            this.gbGlobalSettings.Controls.Add(this.btnRestoreLocalPath);
            this.gbGlobalSettings.Controls.Add(this.tbLocalDataPath);
            this.gbGlobalSettings.Controls.Add(this.cbForceRegionConfig);
            this.gbGlobalSettings.Controls.Add(this.lblLocalDataPath);
            this.gbGlobalSettings.Location = new System.Drawing.Point(12, 31);
            this.gbGlobalSettings.Name = "gbGlobalSettings";
            this.gbGlobalSettings.Size = new System.Drawing.Size(318, 88);
            this.gbGlobalSettings.TabIndex = 38;
            this.gbGlobalSettings.TabStop = false;
            this.gbGlobalSettings.Text = "Configuración Global:";
            // 
            // btnOpenLocalPath
            // 
            this.btnOpenLocalPath.Location = new System.Drawing.Point(278, 30);
            this.btnOpenLocalPath.Name = "btnOpenLocalPath";
            this.btnOpenLocalPath.Size = new System.Drawing.Size(29, 23);
            this.btnOpenLocalPath.TabIndex = 17;
            this.btnOpenLocalPath.Text = "O";
            this.ttSettings.SetToolTip(this.btnOpenLocalPath, "Abrir Ruta");
            this.btnOpenLocalPath.UseVisualStyleBackColor = true;
            this.btnOpenLocalPath.Click += new System.EventHandler(this.btnOpenLocalPath_Click);
            // 
            // btnRestoreLocalPath
            // 
            this.btnRestoreLocalPath.Location = new System.Drawing.Point(6, 58);
            this.btnRestoreLocalPath.Name = "btnRestoreLocalPath";
            this.btnRestoreLocalPath.Size = new System.Drawing.Size(75, 23);
            this.btnRestoreLocalPath.TabIndex = 14;
            this.btnRestoreLocalPath.Text = "Restaurar";
            this.ttSettings.SetToolTip(this.btnRestoreLocalPath, "Restaurar Ruta Predeterminada");
            this.btnRestoreLocalPath.UseVisualStyleBackColor = true;
            this.btnRestoreLocalPath.Click += new System.EventHandler(this.btnRestoreLocalPath_Click);
            // 
            // cbForceRegionConfig
            // 
            this.cbForceRegionConfig.AutoSize = true;
            this.cbForceRegionConfig.Location = new System.Drawing.Point(87, 62);
            this.cbForceRegionConfig.Name = "cbForceRegionConfig";
            this.cbForceRegionConfig.Size = new System.Drawing.Size(166, 17);
            this.cbForceRegionConfig.TabIndex = 6;
            this.cbForceRegionConfig.Text = "Forzar &Regionalización en-US";
            this.cbForceRegionConfig.UseVisualStyleBackColor = true;
            // 
            // gbEllipseSettings
            // 
            this.gbEllipseSettings.Controls.Add(this.btnOpenServicesPath);
            this.gbEllipseSettings.Controls.Add(this.label2);
            this.gbEllipseSettings.Controls.Add(this.btnRestoreServiceFile);
            this.gbEllipseSettings.Controls.Add(this.btnGenerateServiceFileFromDefault);
            this.gbEllipseSettings.Controls.Add(this.btnGenerateServiceFileFromUrl);
            this.gbEllipseSettings.Controls.Add(this.label1);
            this.gbEllipseSettings.Controls.Add(this.tbServiceFileNetworkUrl);
            this.gbEllipseSettings.Location = new System.Drawing.Point(12, 125);
            this.gbEllipseSettings.Name = "gbEllipseSettings";
            this.gbEllipseSettings.Size = new System.Drawing.Size(318, 107);
            this.gbEllipseSettings.TabIndex = 39;
            this.gbEllipseSettings.TabStop = false;
            this.gbEllipseSettings.Text = "Servicios Ellipse:";
            // 
            // btnOpenServicesPath
            // 
            this.btnOpenServicesPath.Location = new System.Drawing.Point(278, 34);
            this.btnOpenServicesPath.Name = "btnOpenServicesPath";
            this.btnOpenServicesPath.Size = new System.Drawing.Size(29, 23);
            this.btnOpenServicesPath.TabIndex = 16;
            this.btnOpenServicesPath.Text = "O";
            this.ttSettings.SetToolTip(this.btnOpenServicesPath, "Abrir Ruta");
            this.btnOpenServicesPath.UseVisualStyleBackColor = true;
            this.btnOpenServicesPath.Click += new System.EventHandler(this.btnOpenServicesPath_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(156, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Generar Archivo Personalizado:";
            // 
            // btnRestoreServiceFile
            // 
            this.btnRestoreServiceFile.Location = new System.Drawing.Point(6, 75);
            this.btnRestoreServiceFile.Name = "btnRestoreServiceFile";
            this.btnRestoreServiceFile.Size = new System.Drawing.Size(75, 23);
            this.btnRestoreServiceFile.TabIndex = 4;
            this.btnRestoreServiceFile.Text = "Re&staurar";
            this.btnRestoreServiceFile.UseVisualStyleBackColor = true;
            this.btnRestoreServiceFile.Click += new System.EventHandler(this.btnDeleteLocalEllipseSettings_Click);
            // 
            // btnGenerateServiceFileFromDefault
            // 
            this.btnGenerateServiceFileFromDefault.Location = new System.Drawing.Point(85, 75);
            this.btnGenerateServiceFileFromDefault.Name = "btnGenerateServiceFileFromDefault";
            this.btnGenerateServiceFileFromDefault.Size = new System.Drawing.Size(90, 23);
            this.btnGenerateServiceFileFromDefault.TabIndex = 3;
            this.btnGenerateServiceFileFromDefault.Text = "&Predeterminado";
            this.btnGenerateServiceFileFromDefault.UseVisualStyleBackColor = true;
            this.btnGenerateServiceFileFromDefault.Click += new System.EventHandler(this.btnGenerateDefault_Click);
            // 
            // btnGenerateServiceFileFromUrl
            // 
            this.btnGenerateServiceFileFromUrl.Location = new System.Drawing.Point(181, 75);
            this.btnGenerateServiceFileFromUrl.Name = "btnGenerateServiceFileFromUrl";
            this.btnGenerateServiceFileFromUrl.Size = new System.Drawing.Size(91, 23);
            this.btnGenerateServiceFileFromUrl.TabIndex = 2;
            this.btnGenerateServiceFileFromUrl.Text = "&Copiar a Local";
            this.btnGenerateServiceFileFromUrl.UseVisualStyleBackColor = true;
            this.btnGenerateServiceFileFromUrl.Click += new System.EventHandler(this.btnGenerateFromUrl_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Ruta de Información de Servicios:";
            // 
            // tbServiceFileNetworkUrl
            // 
            this.tbServiceFileNetworkUrl.Location = new System.Drawing.Point(6, 36);
            this.tbServiceFileNetworkUrl.Name = "tbServiceFileNetworkUrl";
            this.tbServiceFileNetworkUrl.Size = new System.Drawing.Size(266, 20);
            this.tbServiceFileNetworkUrl.TabIndex = 0;
            // 
            // cbForceServerList
            // 
            this.cbForceServerList.AutoSize = true;
            this.cbForceServerList.Location = new System.Drawing.Point(6, 32);
            this.cbForceServerList.Name = "cbForceServerList";
            this.cbForceServerList.Size = new System.Drawing.Size(290, 17);
            this.cbForceServerList.TabIndex = 17;
            this.cbForceServerList.Text = "Forzar lista de &servidores/Databases de los archivos xml";
            this.cbForceServerList.UseVisualStyleBackColor = true;
            // 
            // btnOpenTnsnamesPath
            // 
            this.btnOpenTnsnamesPath.Location = new System.Drawing.Point(278, 32);
            this.btnOpenTnsnamesPath.Name = "btnOpenTnsnamesPath";
            this.btnOpenTnsnamesPath.Size = new System.Drawing.Size(29, 23);
            this.btnOpenTnsnamesPath.TabIndex = 14;
            this.btnOpenTnsnamesPath.Text = "O";
            this.ttSettings.SetToolTip(this.btnOpenTnsnamesPath, "Abrir Ruta");
            this.btnOpenTnsnamesPath.UseVisualStyleBackColor = true;
            this.btnOpenTnsnamesPath.Click += new System.EventHandler(this.btnOpenTnsnamesPath_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Ruta de TnsNames:";
            // 
            // tbTnsNameUrl
            // 
            this.tbTnsNameUrl.Location = new System.Drawing.Point(6, 35);
            this.tbTnsNameUrl.Name = "tbTnsNameUrl";
            this.tbTnsNameUrl.Size = new System.Drawing.Size(266, 20);
            this.tbTnsNameUrl.TabIndex = 7;
            // 
            // btnGenerateTnsnamesFile
            // 
            this.btnGenerateTnsnamesFile.Location = new System.Drawing.Point(85, 58);
            this.btnGenerateTnsnamesFile.Name = "btnGenerateTnsnamesFile";
            this.btnGenerateTnsnamesFile.Size = new System.Drawing.Size(90, 23);
            this.btnGenerateTnsnamesFile.TabIndex = 15;
            this.btnGenerateTnsnamesFile.Text = "Generar";
            this.btnGenerateTnsnamesFile.UseVisualStyleBackColor = true;
            this.btnGenerateTnsnamesFile.Click += new System.EventHandler(this.btnGenerateTnsnamesFile_Click);
            // 
            // gbDatabaseSettings
            // 
            this.gbDatabaseSettings.Controls.Add(this.btnRestoreTnsnamesUrl);
            this.gbDatabaseSettings.Controls.Add(this.label4);
            this.gbDatabaseSettings.Controls.Add(this.btnOpenTnsnamesPath);
            this.gbDatabaseSettings.Controls.Add(this.btnGenerateTnsnamesFile);
            this.gbDatabaseSettings.Controls.Add(this.tbTnsNameUrl);
            this.gbDatabaseSettings.Location = new System.Drawing.Point(12, 238);
            this.gbDatabaseSettings.Name = "gbDatabaseSettings";
            this.gbDatabaseSettings.Size = new System.Drawing.Size(318, 91);
            this.gbDatabaseSettings.TabIndex = 40;
            this.gbDatabaseSettings.TabStop = false;
            this.gbDatabaseSettings.Text = "Bases de Datos:";
            // 
            // btnRestoreTnsnamesUrl
            // 
            this.btnRestoreTnsnamesUrl.Location = new System.Drawing.Point(6, 58);
            this.btnRestoreTnsnamesUrl.Name = "btnRestoreTnsnamesUrl";
            this.btnRestoreTnsnamesUrl.Size = new System.Drawing.Size(75, 23);
            this.btnRestoreTnsnamesUrl.TabIndex = 16;
            this.btnRestoreTnsnamesUrl.Text = "Restaurar";
            this.btnRestoreTnsnamesUrl.UseVisualStyleBackColor = true;
            this.btnRestoreTnsnamesUrl.Click += new System.EventHandler(this.btnRestoreTnsnamesUrl_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 65);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(284, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Es creado en el directorio global como EllipseDatabase.xml";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(303, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Relaciona la lista de servidores ellipse con las Bases de Datos:";
            // 
            // btnDeleteCustomDb
            // 
            this.btnDeleteCustomDb.Location = new System.Drawing.Point(105, 81);
            this.btnDeleteCustomDb.Name = "btnDeleteCustomDb";
            this.btnDeleteCustomDb.Size = new System.Drawing.Size(91, 23);
            this.btnDeleteCustomDb.TabIndex = 18;
            this.btnDeleteCustomDb.Text = "Eliminar Local";
            this.btnDeleteCustomDb.UseVisualStyleBackColor = true;
            this.btnDeleteCustomDb.Click += new System.EventHandler(this.btnDeleteCustomDb_Click);
            // 
            // btnGenerateCustomDb
            // 
            this.btnGenerateCustomDb.Location = new System.Drawing.Point(9, 81);
            this.btnGenerateCustomDb.Name = "btnGenerateCustomDb";
            this.btnGenerateCustomDb.Size = new System.Drawing.Size(90, 23);
            this.btnGenerateCustomDb.TabIndex = 17;
            this.btnGenerateCustomDb.Text = "Predeterminado";
            this.btnGenerateCustomDb.UseVisualStyleBackColor = true;
            this.btnGenerateCustomDb.Click += new System.EventHandler(this.btnGenerateCustomDb_Click);
            // 
            // gbServiceDatabaseRelation
            // 
            this.gbServiceDatabaseRelation.Controls.Add(this.label6);
            this.gbServiceDatabaseRelation.Controls.Add(this.btnDeleteCustomDb);
            this.gbServiceDatabaseRelation.Controls.Add(this.label5);
            this.gbServiceDatabaseRelation.Controls.Add(this.btnGenerateCustomDb);
            this.gbServiceDatabaseRelation.Controls.Add(this.cbForceServerList);
            this.gbServiceDatabaseRelation.Controls.Add(this.label3);
            this.gbServiceDatabaseRelation.Location = new System.Drawing.Point(12, 335);
            this.gbServiceDatabaseRelation.Name = "gbServiceDatabaseRelation";
            this.gbServiceDatabaseRelation.Size = new System.Drawing.Size(318, 115);
            this.gbServiceDatabaseRelation.TabIndex = 41;
            this.gbServiceDatabaseRelation.TabStop = false;
            this.gbServiceDatabaseRelation.Text = "Servicios/Bases de Datos:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(3, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(202, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "Generar Xml de Servers/Bases de Datos:";
            // 
            // SettingsBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(339, 626);
            this.Controls.Add(this.gbServiceDatabaseRelation);
            this.Controls.Add(this.gbDatabaseSettings);
            this.Controls.Add(this.gbEllipseSettings);
            this.Controls.Add(this.gbGlobalSettings);
            this.Controls.Add(this.gbDebugging);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.labelProductName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
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
            this.gbDatabaseSettings.ResumeLayout(false);
            this.gbDatabaseSettings.PerformLayout();
            this.gbServiceDatabaseRelation.ResumeLayout(false);
            this.gbServiceDatabaseRelation.PerformLayout();
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
        private System.Windows.Forms.Button btnRestoreServiceFile;
        private System.Windows.Forms.Button btnGenerateServiceFileFromDefault;
        private System.Windows.Forms.Button btnGenerateServiceFileFromUrl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbServiceFileNetworkUrl;
        private System.Windows.Forms.CheckBox cbForceRegionConfig;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbTnsNameUrl;
        private System.Windows.Forms.Button btnRestoreLocalPath;
        private System.Windows.Forms.Button btnOpenTnsnamesPath;
        private System.Windows.Forms.Button btnGenerateTnsnamesFile;
        private System.Windows.Forms.Button btnOpenLocalPath;
        private System.Windows.Forms.ToolTip ttSettings;
        private System.Windows.Forms.Button btnOpenServicesPath;
        private System.Windows.Forms.GroupBox gbDatabaseSettings;
        private System.Windows.Forms.Button btnRestoreTnsnamesUrl;
        private System.Windows.Forms.CheckBox cbForceServerList;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnDeleteCustomDb;
        private System.Windows.Forms.Button btnGenerateCustomDb;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox gbServiceDatabaseRelation;
        private System.Windows.Forms.Label label6;

    }
}
