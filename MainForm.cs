using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace LimpiadorExcel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            // Cargar el archivo Excel
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos de Excel|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string rutaArchivo = openFileDialog.FileName;

                if (!File.Exists(rutaArchivo))
                {
                    MessageBox.Show("Archivo no encontrado.");
                    return;
                }

                // Procesar el archivo Excel
                try
                {
                    LimpiarExcel(rutaArchivo);
                    MessageBox.Show("Archivo limpio guardado con éxito.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}");
                }
            }
        }

        private void LimpiarExcel(string rutaArchivo)
        {
            using var workbook = new XLWorkbook(rutaArchivo);
            var worksheet = workbook.Worksheet(1); // Primera hoja

            // Obtener el rango utilizado
            var filas = worksheet.RangeUsed().RowsUsed().ToList();

            // Limpiar duplicados (suponiendo que toda la fila es el criterio)
            var filasUnicas = filas
                .GroupBy(f => string.Join("|", f.Cells().Select(c => c.GetValue<string>().Trim())))
                .Select(g => g.First())
                .ToList();

            // Crear nuevo libro
            var nuevoLibro = new XLWorkbook();
            var nuevaHoja = nuevoLibro.AddWorksheet("Limpio");

            // Escribir las filas limpias
            for (int i = 0; i < filasUnicas.Count; i++)
            {
                var fila = filasUnicas[i];
                for (int j = 0; j < fila.CellCount(); j++)
                {
                    string valor = fila.Cell(j + 1).GetValue<string>().Trim();
                    nuevaHoja.Cell(i + 1, j + 1).Value = valor;
                }
            }

            // Guardar el archivo limpio
            string rutaSalida = Path.Combine(Path.GetDirectoryName(rutaArchivo)!, "LIMPIO_" + Path.GetFileName(rutaArchivo));
            nuevoLibro.SaveAs(rutaSalida);

            Console.WriteLine("Archivo limpio guardado en:");
            Console.WriteLine(rutaSalida);
        }
    }
}
