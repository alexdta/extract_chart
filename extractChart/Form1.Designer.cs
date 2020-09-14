namespace extractChart
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnExtraer = new System.Windows.Forms.Button();
            this.imgGrafico = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbGrafico = new System.Windows.Forms.ComboBox();
            this.btnExportar = new System.Windows.Forms.Button();
            this.lblHoja = new System.Windows.Forms.Label();
            this.lblGraf = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.imgGrafico)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExtraer
            // 
            this.btnExtraer.Location = new System.Drawing.Point(41, 12);
            this.btnExtraer.Name = "btnExtraer";
            this.btnExtraer.Size = new System.Drawing.Size(109, 32);
            this.btnExtraer.TabIndex = 0;
            this.btnExtraer.Text = "extraer grafs";
            this.btnExtraer.UseVisualStyleBackColor = true;
            this.btnExtraer.Click += new System.EventHandler(this.btnExtraer_Click);
            // 
            // imgGrafico
            // 
            this.imgGrafico.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.imgGrafico.Location = new System.Drawing.Point(41, 145);
            this.imgGrafico.Name = "imgGrafico";
            this.imgGrafico.Size = new System.Drawing.Size(709, 277);
            this.imgGrafico.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.imgGrafico.TabIndex = 1;
            this.imgGrafico.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Grafico";
            // 
            // cmbGrafico
            // 
            this.cmbGrafico.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbGrafico.Enabled = false;
            this.cmbGrafico.FormattingEnabled = true;
            this.cmbGrafico.Location = new System.Drawing.Point(41, 91);
            this.cmbGrafico.Name = "cmbGrafico";
            this.cmbGrafico.Size = new System.Drawing.Size(121, 24);
            this.cmbGrafico.TabIndex = 4;
            // 
            // btnExportar
            // 
            this.btnExportar.Enabled = false;
            this.btnExportar.Location = new System.Drawing.Point(621, 91);
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.Size = new System.Drawing.Size(129, 32);
            this.btnExportar.TabIndex = 6;
            this.btnExportar.Text = "exportar a word";
            this.btnExportar.UseVisualStyleBackColor = true;
            this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
            // 
            // lblHoja
            // 
            this.lblHoja.AutoSize = true;
            this.lblHoja.Location = new System.Drawing.Point(41, 125);
            this.lblHoja.Name = "lblHoja";
            this.lblHoja.Size = new System.Drawing.Size(37, 17);
            this.lblHoja.TabIndex = 7;
            this.lblHoja.Text = "Hoja";
            // 
            // lblGraf
            // 
            this.lblGraf.AutoSize = true;
            this.lblGraf.Location = new System.Drawing.Point(108, 125);
            this.lblGraf.Name = "lblGraf";
            this.lblGraf.Size = new System.Drawing.Size(54, 17);
            this.lblGraf.TabIndex = 8;
            this.lblGraf.Text = "Grafico";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lblGraf);
            this.Controls.Add(this.lblHoja);
            this.Controls.Add(this.btnExportar);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.imgGrafico);
            this.Controls.Add(this.btnExtraer);
            this.Controls.Add(this.cmbGrafico);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.imgGrafico)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExtraer;
        private System.Windows.Forms.PictureBox imgGrafico;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbGrafico;
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.Label lblHoja;
        private System.Windows.Forms.Label lblGraf;
    }
}

