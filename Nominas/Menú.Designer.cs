namespace Nominas
{
    partial class Menú
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menú));
            this.bARtOP = new MaterialSkin.Controls.MaterialDivider();
            this.monoFlat_ControlBox1 = new MonoFlat.MonoFlat_ControlBox();
            this.butAdministracion = new System.Windows.Forms.PictureBox();
            this.butProduccion = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.butAdministracion)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.butProduccion)).BeginInit();
            this.SuspendLayout();
            // 
            // bARtOP
            // 
            this.bARtOP.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(41)))), ((int)(((byte)(50)))));
            this.bARtOP.Depth = 0;
            this.bARtOP.Dock = System.Windows.Forms.DockStyle.Top;
            this.bARtOP.Location = new System.Drawing.Point(0, 0);
            this.bARtOP.MouseState = MaterialSkin.MouseState.HOVER;
            this.bARtOP.Name = "bARtOP";
            this.bARtOP.Size = new System.Drawing.Size(554, 41);
            this.bARtOP.TabIndex = 5;
            this.bARtOP.Text = "materialDivider1";
            // 
            // monoFlat_ControlBox1
            // 
            this.monoFlat_ControlBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.monoFlat_ControlBox1.EnableHoverHighlight = false;
            this.monoFlat_ControlBox1.EnableMaximizeButton = true;
            this.monoFlat_ControlBox1.EnableMinimizeButton = true;
            this.monoFlat_ControlBox1.Location = new System.Drawing.Point(458, 15);
            this.monoFlat_ControlBox1.Name = "monoFlat_ControlBox1";
            this.monoFlat_ControlBox1.Size = new System.Drawing.Size(100, 25);
            this.monoFlat_ControlBox1.TabIndex = 6;
            this.monoFlat_ControlBox1.Text = "monoFlat_ControlBox1";
            // 
            // butAdministracion
            // 
            this.butAdministracion.BackColor = System.Drawing.Color.Transparent;
            this.butAdministracion.Image = global::Nominas.Properties.Resources.Administracion2;
            this.butAdministracion.Location = new System.Drawing.Point(311, 56);
            this.butAdministracion.Name = "butAdministracion";
            this.butAdministracion.Size = new System.Drawing.Size(171, 162);
            this.butAdministracion.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.butAdministracion.TabIndex = 24;
            this.butAdministracion.TabStop = false;
            this.butAdministracion.Click += new System.EventHandler(this.ButAdministracion_Click);
            // 
            // butProduccion
            // 
            this.butProduccion.BackColor = System.Drawing.Color.Transparent;
            this.butProduccion.Image = global::Nominas.Properties.Resources.Producción;
            this.butProduccion.Location = new System.Drawing.Point(72, 56);
            this.butProduccion.Name = "butProduccion";
            this.butProduccion.Size = new System.Drawing.Size(171, 162);
            this.butProduccion.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.butProduccion.TabIndex = 23;
            this.butProduccion.TabStop = false;
            this.butProduccion.Click += new System.EventHandler(this.ButProduccion_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(41)))), ((int)(((byte)(50)))));
            this.label1.Font = new System.Drawing.Font("Segoe UI Semibold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(247, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 28);
            this.label1.TabIndex = 25;
            this.label1.Text = "MENÚ";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Utsaah", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(84, 227);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(145, 29);
            this.label2.TabIndex = 26;
            this.label2.Text = "PRODUCCIÓN";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Utsaah", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(306, 227);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(182, 29);
            this.label3.TabIndex = 27;
            this.label3.Text = "ADMINISTRACIÓN";
            // 
            // Menú
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(554, 263);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.butAdministracion);
            this.Controls.Add(this.butProduccion);
            this.Controls.Add(this.monoFlat_ControlBox1);
            this.Controls.Add(this.bARtOP);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Menú";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Menú";
            ((System.ComponentModel.ISupportInitialize)(this.butAdministracion)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.butProduccion)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MaterialSkin.Controls.MaterialDivider bARtOP;
        private MonoFlat.MonoFlat_ControlBox monoFlat_ControlBox1;
        private System.Windows.Forms.PictureBox butAdministracion;
        private System.Windows.Forms.PictureBox butProduccion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}