namespace EllipseRequisitionServiceExcelAddIn
{
    partial class AssuranceSettingsBox
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AssuranceSettingsBox));
            this.labelProductName = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.gbSettingsSpecifics = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbMinItemValue = new System.Windows.Forms.TextBox();
            this.cbCostCenterMatch = new System.Windows.Forms.CheckBox();
            this.tbLocalDataPath = new System.Windows.Forms.TextBox();
            this.lblLocalDataPath = new System.Windows.Forms.Label();
            this.gbSettingsGeneral = new System.Windows.Forms.GroupBox();
            this.btnOpenLocalPath = new System.Windows.Forms.Button();
            this.btnRestoreLocalPath = new System.Windows.Forms.Button();
            this.ttSettings = new System.Windows.Forms.ToolTip(this.components);
            this.gbSettingsSpecifics.SuspendLayout();
            this.gbSettingsGeneral.SuspendLayout();
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
            this.labelProductName.Text = "Requisition Service - Garantías";
            this.labelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Location = new System.Drawing.Point(255, 246);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 26);
            this.okButton.TabIndex = 30;
            this.okButton.Text = "&OK";
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // gbSettingsSpecifics
            // 
            this.gbSettingsSpecifics.Controls.Add(this.label1);
            this.gbSettingsSpecifics.Controls.Add(this.tbMinItemValue);
            this.gbSettingsSpecifics.Controls.Add(this.cbCostCenterMatch);
            this.gbSettingsSpecifics.Location = new System.Drawing.Point(12, 125);
            this.gbSettingsSpecifics.Name = "gbSettingsSpecifics";
            this.gbSettingsSpecifics.Size = new System.Drawing.Size(318, 103);
            this.gbSettingsSpecifics.TabIndex = 37;
            this.gbSettingsSpecifics.TabStop = false;
            this.gbSettingsSpecifics.Text = "Configuración Específica:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(187, 13);
            this.label1.TabIndex = 39;
            this.label1.Text = "Valor Mínimo para seguimiento (USD):";
            // 
            // tbMinItemValue
            // 
            this.tbMinItemValue.Location = new System.Drawing.Point(6, 63);
            this.tbMinItemValue.Name = "tbMinItemValue";
            this.tbMinItemValue.Size = new System.Drawing.Size(100, 20);
            this.tbMinItemValue.TabIndex = 38;
            // 
            // cbCostCenterMatch
            // 
            this.cbCostCenterMatch.AutoSize = true;
            this.cbCostCenterMatch.Location = new System.Drawing.Point(6, 23);
            this.cbCostCenterMatch.Name = "cbCostCenterMatch";
            this.cbCostCenterMatch.Size = new System.Drawing.Size(181, 17);
            this.cbCostCenterMatch.TabIndex = 37;
            this.cbCostCenterMatch.Text = "Coincidencia de Centro de Costo";
            this.cbCostCenterMatch.UseVisualStyleBackColor = true;
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
            // gbSettingsGeneral
            // 
            this.gbSettingsGeneral.Controls.Add(this.btnOpenLocalPath);
            this.gbSettingsGeneral.Controls.Add(this.btnRestoreLocalPath);
            this.gbSettingsGeneral.Controls.Add(this.tbLocalDataPath);
            this.gbSettingsGeneral.Controls.Add(this.lblLocalDataPath);
            this.gbSettingsGeneral.Location = new System.Drawing.Point(12, 31);
            this.gbSettingsGeneral.Name = "gbSettingsGeneral";
            this.gbSettingsGeneral.Size = new System.Drawing.Size(318, 88);
            this.gbSettingsGeneral.TabIndex = 38;
            this.gbSettingsGeneral.TabStop = false;
            this.gbSettingsGeneral.Text = "Configuración General:";
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
            // AssuranceSettingsBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(339, 284);
            this.Controls.Add(this.gbSettingsGeneral);
            this.Controls.Add(this.gbSettingsSpecifics);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.labelProductName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AssuranceSettingsBox";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Seguimiento a Garantías";
            this.Load += new System.EventHandler(this.SettingsBox_Load);
            this.gbSettingsSpecifics.ResumeLayout(false);
            this.gbSettingsSpecifics.PerformLayout();
            this.gbSettingsGeneral.ResumeLayout(false);
            this.gbSettingsGeneral.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label labelProductName;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.GroupBox gbSettingsSpecifics;
        private System.Windows.Forms.TextBox tbLocalDataPath;
        private System.Windows.Forms.CheckBox cbCostCenterMatch;
        private System.Windows.Forms.Label lblLocalDataPath;
        private System.Windows.Forms.GroupBox gbSettingsGeneral;
        private System.Windows.Forms.Button btnRestoreLocalPath;
        private System.Windows.Forms.Button btnOpenLocalPath;
        private System.Windows.Forms.ToolTip ttSettings;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbMinItemValue;
    }
}
