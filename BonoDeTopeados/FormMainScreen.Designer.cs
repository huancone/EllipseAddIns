namespace BonoDeTopeados
{
    partial class FormMainScreen
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
            this.btnLoadEmployeeTurns = new System.Windows.Forms.Button();
            this.cbPeriodMode = new System.Windows.Forms.ComboBox();
            this.btnLoadEmployeeTurn886 = new System.Windows.Forms.Button();
            this.lblProgress = new System.Windows.Forms.Label();
            this.tbPeriod = new System.Windows.Forms.TextBox();
            this.tbYear = new System.Windows.Forms.TextBox();
            this.lblYear = new System.Windows.Forms.Label();
            this.lblPeriod = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnLoadEmployeeTurns
            // 
            this.btnLoadEmployeeTurns.Location = new System.Drawing.Point(13, 83);
            this.btnLoadEmployeeTurns.Name = "btnLoadEmployeeTurns";
            this.btnLoadEmployeeTurns.Size = new System.Drawing.Size(175, 23);
            this.btnLoadEmployeeTurns.TabIndex = 0;
            this.btnLoadEmployeeTurns.Text = "Cargar Empleados Turno";
            this.btnLoadEmployeeTurns.UseVisualStyleBackColor = true;
            this.btnLoadEmployeeTurns.Click += new System.EventHandler(this.btnLoadEmployeeTurns_Click);
            // 
            // cbPeriodMode
            // 
            this.cbPeriodMode.FormattingEnabled = true;
            this.cbPeriodMode.Location = new System.Drawing.Point(13, 141);
            this.cbPeriodMode.Name = "cbPeriodMode";
            this.cbPeriodMode.Size = new System.Drawing.Size(174, 21);
            this.cbPeriodMode.TabIndex = 1;
            // 
            // btnLoadEmployeeTurn886
            // 
            this.btnLoadEmployeeTurn886.Location = new System.Drawing.Point(14, 113);
            this.btnLoadEmployeeTurn886.Name = "btnLoadEmployeeTurn886";
            this.btnLoadEmployeeTurn886.Size = new System.Drawing.Size(174, 23);
            this.btnLoadEmployeeTurn886.TabIndex = 2;
            this.btnLoadEmployeeTurn886.Text = "Cargar Turnos Covid";
            this.btnLoadEmployeeTurn886.UseVisualStyleBackColor = true;
            this.btnLoadEmployeeTurn886.Click += new System.EventHandler(this.btnLoadEmployeeTurn886_Click);
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(84, 181);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(24, 13);
            this.lblProgress.TabIndex = 3;
            this.lblProgress.Text = "0/0";
            // 
            // tbPeriod
            // 
            this.tbPeriod.Location = new System.Drawing.Point(65, 57);
            this.tbPeriod.Name = "tbPeriod";
            this.tbPeriod.Size = new System.Drawing.Size(122, 20);
            this.tbPeriod.TabIndex = 4;
            // 
            // tbYear
            // 
            this.tbYear.Location = new System.Drawing.Point(65, 31);
            this.tbYear.Name = "tbYear";
            this.tbYear.Size = new System.Drawing.Size(122, 20);
            this.tbYear.TabIndex = 5;
            // 
            // lblYear
            // 
            this.lblYear.AutoSize = true;
            this.lblYear.Location = new System.Drawing.Point(13, 34);
            this.lblYear.Name = "lblYear";
            this.lblYear.Size = new System.Drawing.Size(29, 13);
            this.lblYear.TabIndex = 6;
            this.lblYear.Text = "Año:";
            // 
            // lblPeriod
            // 
            this.lblPeriod.AutoSize = true;
            this.lblPeriod.Location = new System.Drawing.Point(13, 60);
            this.lblPeriod.Name = "lblPeriod";
            this.lblPeriod.Size = new System.Drawing.Size(46, 13);
            this.lblPeriod.TabIndex = 7;
            this.lblPeriod.Text = "Periodo:";
            // 
            // FormMainScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(219, 195);
            this.Controls.Add(this.lblPeriod);
            this.Controls.Add(this.lblYear);
            this.Controls.Add(this.tbYear);
            this.Controls.Add(this.tbPeriod);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.btnLoadEmployeeTurn886);
            this.Controls.Add(this.cbPeriodMode);
            this.Controls.Add(this.btnLoadEmployeeTurns);
            this.Name = "FormMainScreen";
            this.Text = "Bono de Topeados  - Manejo de Carbón";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoadEmployeeTurns;
        private System.Windows.Forms.ComboBox cbPeriodMode;
        private System.Windows.Forms.Button btnLoadEmployeeTurn886;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.TextBox tbPeriod;
        private System.Windows.Forms.TextBox tbYear;
        private System.Windows.Forms.Label lblYear;
        private System.Windows.Forms.Label lblPeriod;
    }
}