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
            this.SuspendLayout();
            // 
            // btnLoadEmployeeTurns
            // 
            this.btnLoadEmployeeTurns.Location = new System.Drawing.Point(12, 12);
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
            this.cbPeriodMode.Location = new System.Drawing.Point(12, 70);
            this.cbPeriodMode.Name = "cbPeriodMode";
            this.cbPeriodMode.Size = new System.Drawing.Size(174, 21);
            this.cbPeriodMode.TabIndex = 1;
            // 
            // btnLoadEmployeeTurn886
            // 
            this.btnLoadEmployeeTurn886.Location = new System.Drawing.Point(13, 42);
            this.btnLoadEmployeeTurn886.Name = "btnLoadEmployeeTurn886";
            this.btnLoadEmployeeTurn886.Size = new System.Drawing.Size(174, 23);
            this.btnLoadEmployeeTurn886.TabIndex = 2;
            this.btnLoadEmployeeTurn886.Text = "Cargar Turnos Covid";
            this.btnLoadEmployeeTurn886.UseVisualStyleBackColor = true;
            this.btnLoadEmployeeTurn886.Click += new System.EventHandler(this.btnLoadEmployeeTurn886_Click);
            // 
            // FormMainScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(219, 144);
            this.Controls.Add(this.btnLoadEmployeeTurn886);
            this.Controls.Add(this.cbPeriodMode);
            this.Controls.Add(this.btnLoadEmployeeTurns);
            this.Name = "FormMainScreen";
            this.Text = "Bono de Topeados  - Manejo de Carbón";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnLoadEmployeeTurns;
        private System.Windows.Forms.ComboBox cbPeriodMode;
        private System.Windows.Forms.Button btnLoadEmployeeTurn886;
    }
}