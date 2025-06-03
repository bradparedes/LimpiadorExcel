using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LimpiadorExcel
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Configuración de la aplicación
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Selecciona un archivo Excel",
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("No se seleccionó ningún archivo.", "Operación cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string rutaArchivo = openFileDialog.FileName;

            if (!File.Exists(rutaArchivo))
            {
                MessageBox.Show("Archivo no encontrado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var workbook = new XLWorkbook(rutaArchivo);
            var worksheet = workbook.Worksheet(1); // Primera hoja

            var rango = worksheet.RangeUsed();
            if (rango == null)
            {
                MessageBox.Show("La hoja seleccionada no contiene datos.", "Hoja vacía", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var filas = rango.RowsUsed().ToList();

            var filasUnicas = filas
                .GroupBy(f => string.Join("|", f.Cells().Select(c => c.GetValue<string>().Trim())))
                .Select(g => g.First())
                .ToList();

            var nuevoLibro = new XLWorkbook();
            var nuevaHoja = nuevoLibro.AddWorksheet("Limpio");

            for (int i = 0; i < filasUnicas.Count; i++)
            {
                var fila = filasUnicas[i];
                for (int j = 0; j < fila.CellCount(); j++)
                {
                    string valor = fila.Cell(j + 1).GetValue<string>().Trim();
                    nuevaHoja.Cell(i + 1, j + 1).Value = valor;
                }
            }

            string carpeta = Path.GetDirectoryName(rutaArchivo)!;
            string nombreSalida = "LIMPIO_" + Path.GetFileName(rutaArchivo);
            string rutaSalida = Path.Combine(carpeta, nombreSalida);
            nuevoLibro.SaveAs(rutaSalida);

            MessageBox.Show($"Archivo limpio guardado como:\n\n{rutaSalida}", "Proceso completado", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
