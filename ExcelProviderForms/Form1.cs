using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProviderForms
{
    public partial class Form1 : Form
    {
        private DatabaseManager dbManager;

        public Form1()
        {
            InitializeComponent();
            txtOldPath.Text = @"\\fs-plaza1";
            txtNewPath.Text = @"\\nas-plaza1";

            // Inicializar DB
            dbManager = new DatabaseManager();

            // Inicializar ComboBox
            cmbTipoArchivo.Items.Add("Excel");
            cmbTipoArchivo.Items.Add("PowerPoint");
            cmbTipoArchivo.Items.Add("Ambos");
            cmbTipoArchivo.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Title = "Selecciona una carpeta con archivos";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string seleccion = cmbTipoArchivo.SelectedItem.ToString();

                    string[] archivos = Directory.GetFiles(dialog.FileName, "*.*", SearchOption.AllDirectories)
                        .Where(f =>
                            (seleccion == "Excel" &&
                                (f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))) ||
                            (seleccion == "PowerPoint" &&
                                (f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase))) ||
                            (seleccion == "Ambos" &&
                                (f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                 f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase)))
                        )
                        .ToArray();

                    listBox1.Items.Clear();

                    foreach (string archivo in archivos)
                    {
                        ArchivoOffice archivoObj;

                        if (archivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || archivo.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                            archivoObj = new ExcelArchivo(Path.GetFileName(archivo), archivo);
                        else
                            archivoObj = new PowerPointArchivo(Path.GetFileName(archivo), archivo);

                        dbManager.RegistrarArchivo(archivoObj.Nombre, archivoObj.RutaCompleta);
                        listBox1.Items.Add(archivoObj);
                    }
                }
            }
        }

        private async void btnProcesar_Click(object sender, EventArgs e)
        {
            string oldLink = txtOldPath.Text.Trim();
            string newLink = txtNewPath.Text.Trim();

            if (string.IsNullOrWhiteSpace(oldLink) || string.IsNullOrWhiteSpace(newLink))
            {
                MessageBox.Show("Por favor ingresa ambas rutas: antigua y nueva.", "Error");
                return;
            }

            listBox2.Items.Clear();

            await Task.Run(() =>
            {
                foreach (ArchivoOffice item in listBox1.Items)
                {
                    try
                    {
                        item.Procesar(oldLink, newLink, dbManager,
                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add(mensaje))));

                        dbManager.MarcarProcesado(item.RutaCompleta);
                    }
                    catch (Exception ex)
                    {
                        dbManager.MarcarFallido(item.RutaCompleta);
                        Invoke((MethodInvoker)(() => listBox2.Items.Add($"❌ ERROR en {item.Nombre}: {ex.Message}")));
                    }
                }
            });
            MessageBox.Show("✅ Procesamiento completado.");
        }


        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            dbManager?.Dispose();
            base.OnFormClosed(e);
        }
    }
}
