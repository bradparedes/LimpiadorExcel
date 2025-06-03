using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LimpiadorExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Evento de botón para limpiar el archivo Excel
        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            // Mostrar diálogo para seleccionar archivo
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string rutaArchivo = openFileDialog.FileName;
                    LimpiarExcel(rutaArchivo);
                }
            }
        }

        private void LimpiarExcel(string rutaArchivo)
        {
            try
            {
                // Abrir el libro de Excel
                using var workbook = new XLWorkbook(rutaArchivo);
                var worksheet = workbook.Worksheet(1); // Seleccionar la primera hoja

                // Obtener el rango utilizado
                var rango = worksheet.RangeUsed();

                // Verificar si el rango no es null o si tiene filas
                if (rango == null || !rango.RowsUsed().Any())
                {
                    MessageBox.Show("El archivo Excel está vacío o no tiene filas con datos.");
                    return;
                }

                // Obtener las filas usadas
                var filas = rango.RowsUsed().ToList();

                // Limpiar duplicados (usando toda la fila como criterio)
                var filasUnicas = filas
                    .GroupBy(f => string.Join("|", f.Cells().Select(c => c.GetValue<string>().Trim())))
                    .Select(g => g.First())
                    .ToList();

                // Crear un nuevo libro de Excel
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

                MessageBox.Show($"Archivo limpio guardado en: {rutaSalida}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al procesar el archivo: {ex.Message}");
            }
        }
    }
}
