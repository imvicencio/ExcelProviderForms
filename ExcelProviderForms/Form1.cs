using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProviderForms
{
    public partial class Form1 : Form
    {
        private DatabaseManager dbManager;

        public Form1()
        {
            InitializeComponent();
            txtOldPath.Text = @"\\nas-plaza1";
            txtNewPath.Text = @"\\stmplfsrgprdeastus2.file.core.windows.net\fs-mpl-chile\Gerencia_Analisis_del_Negocio";

            // Inicializar el manager de base de datos
            dbManager = new DatabaseManager();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Title = "Selecciona una carpeta con archivos Excel";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    //string[] archivosExcel = Directory.GetFiles(dialog.FileName, "*.xlsx", SearchOption.AllDirectories);

                    string[] archivosExcel = Directory.GetFiles(dialog.FileName, "*.*", SearchOption.AllDirectories)
    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
    .ToArray();

                    listBox1.Items.Clear();
                    int archivosPendientes = 0;
                    int archivosYaProcesados = 0;

                    foreach (string archivo in archivosExcel)
                    {
                        var excelArchivo = new ExcelArchivo(Path.GetFileName(archivo), archivo);

                        // Registrar el archivo en la base de datos si no existe
                        dbManager.RegistrarArchivo(excelArchivo.Nombre, archivo);

                        // Verificar si ya fue procesado
                        if (dbManager.EstaProcesado(archivo))
                        {
                            archivosYaProcesados++;
                        }
                        else
                        {
                            archivosPendientes++;
                        }

                        listBox1.Items.Add(excelArchivo);
                    }

                    if (archivosExcel.Length == 0)
                    {
                        MessageBox.Show("No se encontraron archivos .xlsx en la carpeta seleccionada.", "Información");
                    }
                    else
                    {
                        lblTotalArchivos.Text = $"Total: {archivosExcel.Length} | Pendientes: {archivosPendientes} | Ya procesados: {archivosYaProcesados}";
                    }
                }
            }
        }

        private async void btnProcesar_Click(object sender, EventArgs e)
        {
            string oldLink = txtOldPath.Text.Trim();
            string newLink = txtNewPath.Text.Trim();
            Dictionary<string, string> listLink = new Dictionary<string, string> 
            {
                {@"\\nas-plaza1" , @"\\stmplfsrgprdeastus2.file.core.windows.net\fs-mpl-chile\Gerencia_Analisis_del_Negocio" },
                {@"\\fs-plaza1" , @"\\stmplfsrgprdeastus2.file.core.windows.net\fs-mpl-chile\Gerencia_Analisis_del_Negocio" },
                {@"file:///" , "" }
            }; 

            if (string.IsNullOrWhiteSpace(oldLink) || string.IsNullOrWhiteSpace(newLink))
            {
                MessageBox.Show("Por favor ingresa ambas rutas: antigua y nueva.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            listBox2.Items.Clear();
            int totalArchivos = listBox1.Items.Count;
            int archivosOmitidos = 0;
            int archivosProcesados = 0;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = totalArchivos;
            progressBar1.Value = 0;

            await Task.Run(() =>
            {
                int contador = 0;

                foreach (ExcelArchivo item in listBox1.Items)
                {
                    contador++;
                    string progreso = $"Procesando archivo {item.Nombre} -- {contador} de {totalArchivos}";

                    Invoke((MethodInvoker)(() =>
                    {
                        lblProgreso.Text = progreso;
                        progressBar1.Value = contador;
                    }));

                    // Verificar si ya fue procesado
                    if (dbManager.EstaProcesado(item.RutaCompleta))
                    {
                        archivosOmitidos++;
                        dbManager.Loguear(item.Nombre, "Archivo ya procesado previamente, omitiendo", "OMITIDO",
                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"⏭️ {mensaje}"))));
                        continue;
                    }

                    string excelPath = item.RutaCompleta;
                    Excel.Application excelApp = null;
                    Excel.Workbook workbook = null;
                    bool procesamientoExitoso = true;

                    try
                    {
                        dbManager.Loguear(item.Nombre, "Iniciando procesamiento", "INICIO",
                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔄 {mensaje}"))));

                        excelApp = new Excel.Application
                        {
                            DisplayAlerts = false,
                            AskToUpdateLinks = false,
                            AlertBeforeOverwriting = false
                        };

                        workbook = excelApp.Workbooks.Open(
                            excelPath,
                            UpdateLinks: 0,
                            ReadOnly: false,
                            CorruptLoad: Excel.XlCorruptLoad.xlRepairFile
                        );

                        var rawLinks = workbook.LinkSources(Excel.XlLink.xlExcelLinks);

                        if (rawLinks is Array links)
                        {
                            int vinculosActualizados = 0;
                            int vinculosRotos = 0;
                            int vinculosSinCambios = 0;

                            foreach (var obj in links)
                            {
                                string link = obj.ToString();

                                foreach (var ruta in listLink)
                                {
                                    if (link.StartsWith(ruta.Key, StringComparison.OrdinalIgnoreCase))
                                    {
                                        string updatedLink = Regex.Replace(link, Regex.Escape(ruta.Key), ruta.Value, RegexOptions.IgnoreCase);

                                        if (File.Exists(updatedLink))
                                        {
                                            workbook.ChangeLink(link, updatedLink, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                            vinculosActualizados++;

                                            dbManager.Loguear(item.Nombre, $"Vínculo actualizado: {link} → {updatedLink}", "ACTUALIZADO",
                                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"✅ {mensaje}"))));
                                        }
                                        else
                                        {
                                            workbook.BreakLink(link, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                            vinculosRotos++;

                                            dbManager.Loguear(item.Nombre, $"Vínculo roto eliminado: {link}", "ROTO",
                                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"⚠️ {mensaje}"))));
                                        }
                                    }
                                    else
                                    {
                                        vinculosSinCambios++;
                                        dbManager.Loguear(item.Nombre, $"Vínculo sin cambios: {link}", "SIN_CAMBIOS",
                                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔗 {mensaje}"))));
                                    }
                                }

 
                            }

                            // Log resumen
                            dbManager.Loguear(item.Nombre,
                                $"Resumen: {vinculosActualizados} actualizados, {vinculosRotos} eliminados, {vinculosSinCambios} sin cambios",
                                "RESUMEN");
                        }
                        else
                        {
                            dbManager.Loguear(item.Nombre, "Sin vínculos externos", "SIN_VINCULOS",
                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"📭 {mensaje}"))));
                        }

                        workbook.SaveAs(excelPath, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
                        archivosProcesados++;

                        dbManager.Loguear(item.Nombre, "Procesamiento completado exitosamente", "COMPLETADO");
                    }
                    catch (Exception ex)
                    {
                        procesamientoExitoso = false;
                        dbManager.Loguear(item.Nombre, $"ERROR: {ex.Message}", "ERROR",
                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"❌ {mensaje}"))));
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        // Marcar como procesado solo si fue exitoso
                        if (procesamientoExitoso)
                        {
                            dbManager.MarcarProcesado(item.RutaCompleta);
                        }
                        else
                        {
                            dbManager.MarcarFallido(item.RutaCompleta);
                        }
                    }
                }
            });

            string mensajeResumen = $"✅ Procesamiento completado.\n" +
                                  $"Procesados: {archivosProcesados}\n" +
                                  $"Omitidos (ya procesados): {archivosOmitidos}";

            MessageBox.Show(mensajeResumen, "Listo");
            lblProgreso.Text = "Procesamiento completado.";

            // Log final del proceso
            dbManager.Loguear("SISTEMA", $"Lote completado - {archivosProcesados} procesados, {archivosOmitidos} omitidos", "LOTE_COMPLETADO");
        }

        // Método para limpiar recursos al cerrar la aplicación
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            dbManager?.Dispose();
            base.OnFormClosed(e);
        }
    }
}