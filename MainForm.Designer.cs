namespace LimpiadorExcel
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnLimpiar;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnLimpiar = new System.Windows.Forms.Button();
            this.SuspendLayout();

            // btnLimpiar
            this.btnLimpiar.Location = new System.Drawing.Point(100, 100);
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.Size = new System.Drawing.Size(200, 50);
            this.btnLimpiar.TabIndex = 0;
            this.btnLimpiar.Text = "Limpiar Excel";
            this.btnLimpiar.UseVisualStyleBackColor = true;
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);

            // MainForm
            this.ClientSize = new System.Drawing.Size(400, 200);
            this.Controls.Add(this.btnLimpiar);
            this.Name = "MainForm";
            this.Text = "Limpiador Excel";
            this.ResumeLayout(false);
        }
    }
}
