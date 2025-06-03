namespace LimpiadorExcel
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        // Método para limpiar recursos
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        // Método de inicialización del formulario
        private void InitializeComponent()
        {
            this.btnLimpiar = new System.Windows.Forms.Button();
            this.SuspendLayout();

            // 
            // btnLimpiar
            // 
            this.btnLimpiar.Location = new System.Drawing.Point(100, 100);
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.Size = new System.Drawing.Size(120, 30);
            this.btnLimpiar.TabIndex = 0;
            this.btnLimpiar.Text = "Limpiar Excel";
            this.btnLimpiar.UseVisualStyleBackColor = true;
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);

            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(320, 240);
            this.Controls.Add(this.btnLimpiar);
            this.Name = "Form1";
            this.Text = "Limpiador de Excel";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Button btnLimpiar;
    }
}
