using ExcelProviderForms;
using Microsoft.Office.Core;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
                dialog.Title = "Selecciona una carpeta con archivos PowerPoint";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string[] archivosPowerPoint = Directory.GetFiles(dialog.FileName, "*.*", SearchOption.AllDirectories)
        .Where(f => f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".pptm", StringComparison.OrdinalIgnoreCase))
        .ToArray();

                    listBox1.Items.Clear();
                    int archivosPendientes = 0;
                    int archivosYaProcesados = 0;

                    foreach (string archivo in archivosPowerPoint)
                    {
                        var pptArchivo = new PowerPointArchivo(Path.GetFileName(archivo), archivo);

                        // Registrar el archivo en la base de datos si no existe
                        dbManager.RegistrarArchivo(pptArchivo.Nombre, archivo);

                        // Verificar si ya fue procesado
                        if (dbManager.EstaProcesado(archivo))
                        {
                            archivosYaProcesados++;
                        }
                        else
                        {
                            archivosPendientes++;
                        }

                        listBox1.Items.Add(pptArchivo);
                    }

                    if (archivosPowerPoint.Length == 0)
                    {
                        MessageBox.Show("No se encontraron archivos PowerPoint en la carpeta seleccionada.", "Información");
                    }
                    else
                    {
                        lblTotalArchivos.Text = $"Total: {archivosPowerPoint.Length} | Pendientes: {archivosPendientes} | Ya procesados: {archivosYaProcesados}";
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

                foreach (PowerPointArchivo item in listBox1.Items)
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

                    string pptPath = item.RutaCompleta;
                    PowerPoint.Application pptApp = null;
                    PowerPoint.Presentation presentation = null;
                    bool procesamientoExitoso = true;

                    try
                    {
                        dbManager.Loguear(item.Nombre, "Iniciando procesamiento", "INICIO",
                            mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔄 {mensaje}"))));

                        pptApp = new PowerPoint.Application();
                        presentation = pptApp.Presentations.Open(pptPath);

                        int vinculosActualizados = 0;
                        int vinculosRotos = 0;
                        int vinculosSinCambios = 0;

                        // Procesar vínculos en diapositivas
                        ProcessSlides(presentation, oldLink, newLink, item.Nombre, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);

                        // Procesar vínculos en masters
                        ProcessMasters(presentation, oldLink, newLink, item.Nombre, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);

                        // Log resumen
                        dbManager.Loguear(item.Nombre,
                            $"Resumen: {vinculosActualizados} actualizados, {vinculosRotos} eliminados, {vinculosSinCambios} sin cambios",
                            "RESUMEN");

                        if (vinculosActualizados == 0 && vinculosRotos == 0 && vinculosSinCambios == 0)
                        {
                            dbManager.Loguear(item.Nombre, "Sin vínculos encontrados", "SIN_VINCULOS",
                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"📭 {mensaje}"))));
                        }

                        presentation.Save();
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
                        if (presentation != null)
                        {
                            presentation.Close();
                            Marshal.ReleaseComObject(presentation);
                        }
                        if (pptApp != null)
                        {
                            pptApp.Quit();
                            Marshal.ReleaseComObject(pptApp);
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

        private void ProcessSlides(PowerPoint.Presentation presentation, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                ProcessShapes(slide.Shapes, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
            }
        }

        private void ProcessMasters(PowerPoint.Presentation presentation, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                foreach (PowerPoint.Master master in presentation.Designs.Cast<PowerPoint.Design>().Select(d => d.SlideMaster))
                {
                    ProcessShapes(master.Shapes, oldLink, newLink, nombreArchivo,
                                  ref vinculosActualizados,
                                  ref vinculosRotos,
                                  ref vinculosSinCambios);

                    foreach (PowerPoint.CustomLayout layout in master.CustomLayouts)
                    {
                        ProcessShapes(layout.Shapes, oldLink, newLink, nombreArchivo,
                                      ref vinculosActualizados,
                                      ref vinculosRotos,
                                      ref vinculosSinCambios);
                    }
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error procesando masters: {ex.Message}", "ERROR");
            }
        }

        private void ProcessShapes(PowerPoint.Shapes shapes, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            foreach (PowerPoint.Shape shape in shapes)
            {
                try
                {
                    // Procesar objetos OLE vinculados
                    if (shape.Type == MsoShapeType.msoLinkedOLEObject)
                    {
                        ProcessOLELink(shape, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
                    }

                    // Procesar imágenes vinculadas
                    if (shape.Type == MsoShapeType.msoLinkedPicture)
                    {
                        ProcessPictureLink(shape, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
                    }

                    // Procesar hipervínculos en formas
                    ProcessHyperlinks(shape, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);

                    // Procesar formas agrupadas recursivamente
                    if (shape.Type == MsoShapeType.msoGroup)
                    {
                        ProcessShapes((PowerPoint.Shapes)shape.GroupItems, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
                    }

                    // Procesar vínculos en texto
                    ProcessTextFrameLinks(shape, oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
                }
                catch (Exception ex)
                {
                    // Log del error pero continuar con las siguientes formas
                    dbManager.Loguear(nombreArchivo, $"Error procesando forma: {ex.Message}", "ERROR_FORMA");
                }
            }
        }

        private void ProcessOLELink(PowerPoint.Shape shape, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                if (shape.LinkFormat != null)
                {
                    string currentPath = shape.LinkFormat.SourceFullName;
                    if (!string.IsNullOrEmpty(currentPath))
                    {
                        dbManager.Loguear(nombreArchivo, $"Comparando OLE: [Ruta Actual: '{currentPath}'] con [Ruta Antigua: '{oldLink}']", "DEBUG");

                        if (currentPath.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                        {
                            string updatedPath = Regex.Replace(currentPath, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);

                            bool newPathExists = File.Exists(updatedPath);
                            dbManager.Loguear(nombreArchivo, $"Verificando existencia de nueva ruta OLE: {updatedPath} -> {(newPathExists ? "ENCONTRADO" : "NO ENCONTRADO")}", "DEBUG");

                            if (newPathExists)
                            {
                                shape.LinkFormat.SourceFullName = updatedPath;
                                vinculosActualizados++;
                                dbManager.Loguear(nombreArchivo, $"OLE actualizado: {currentPath} → {updatedPath}", "ACTUALIZADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"✅ {mensaje}"))));
                            }
                            else
                            {
                                shape.LinkFormat.BreakLink();
                                vinculosRotos++;
                                dbManager.Loguear(nombreArchivo, $"Vínculo OLE roto, objeto incrustado: {currentPath}", "EMBEBIDO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔗 {mensaje}"))));
                            }
                        }
                        else
                        {
                            vinculosSinCambios++;
                            dbManager.Loguear(nombreArchivo, $"OLE sin cambios: {currentPath}", "SIN_CAMBIOS",
                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔗 {mensaje}"))));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error en OLE: {ex.Message}", "ERROR");
            }
        }

        private void ProcessPictureLink(PowerPoint.Shape shape, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                if (shape.LinkFormat != null)
                {
                    string currentPath = shape.LinkFormat.SourceFullName;
                    if (!string.IsNullOrEmpty(currentPath))
                    {
                        dbManager.Loguear(nombreArchivo, $"Comparando Imagen: [Ruta Actual: '{currentPath}'] con [Ruta Antigua: '{oldLink}']", "DEBUG");

                        if (currentPath.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                        {
                            string updatedPath = Regex.Replace(currentPath, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);

                            bool newPathExists = File.Exists(updatedPath);
                            dbManager.Loguear(nombreArchivo, $"Verificando existencia de nueva ruta de imagen: {updatedPath} -> {(newPathExists ? "ENCONTRADO" : "NO ENCONTRADO")}", "DEBUG");

                            if (newPathExists)
                            {
                                shape.LinkFormat.SourceFullName = updatedPath;
                                vinculosActualizados++;
                                dbManager.Loguear(nombreArchivo, $"Imagen actualizada: {currentPath} → {updatedPath}", "ACTUALIZADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"✅ {mensaje}"))));
                            }
                            else
                            {
                                shape.LinkFormat.BreakLink();
                                vinculosRotos++;
                                dbManager.Loguear(nombreArchivo, $"Vínculo de imagen roto, objeto incrustado: {currentPath}", "EMBEBIDO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔗 {mensaje}"))));
                            }
                        }
                        else
                        {
                            vinculosSinCambios++;
                            dbManager.Loguear(nombreArchivo, $"Imagen sin cambios: {currentPath}", "SIN_CAMBIOS",
                                mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🔗 {mensaje}"))));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error en imagen: {ex.Message}", "ERROR");
            }
        }

        private void ProcessHyperlinks(PowerPoint.Shape shape, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                if (shape.ActionSettings != null)
                {
                    ProcessActionSetting(shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick],
                        oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);

                    ProcessActionSetting(shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseOver],
                        oldLink, newLink, nombreArchivo, ref vinculosActualizados, ref vinculosRotos, ref vinculosSinCambios);
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error en hipervínculos: {ex.Message}", "ERROR");
            }
        }

        private void ProcessActionSetting(PowerPoint.ActionSetting actionSetting, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                if (actionSetting.Action == PowerPoint.PpActionType.ppActionHyperlink)
                {
                    string currentAddress = actionSetting.Hyperlink.Address;
                    if (!string.IsNullOrEmpty(currentAddress))
                    {
                        dbManager.Loguear(nombreArchivo, $"Comparando Hipervínculo: [Ruta Actual: '{currentAddress}'] con [Ruta Antigua: '{oldLink}']", "DEBUG");

                        if (currentAddress.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                        {
                            string updatedAddress = Regex.Replace(currentAddress, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);

                            bool newPathExists = File.Exists(updatedAddress) || Directory.Exists(updatedAddress);
                            dbManager.Loguear(nombreArchivo, $"Verificando existencia de nueva ruta de hipervínculo: {updatedAddress} -> {(newPathExists ? "ENCONTRADO" : "NO ENCONTRADO")}", "DEBUG");

                            if (newPathExists)
                            {
                                actionSetting.Hyperlink.Address = updatedAddress;
                                vinculosActualizados++;
                                dbManager.Loguear(nombreArchivo, $"Hipervínculo actualizado: {currentAddress} → {updatedAddress}", "ACTUALIZADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"✅ {mensaje}"))));
                            }
                            else
                            {
                                actionSetting.Action = PowerPoint.PpActionType.ppActionNone;
                                vinculosRotos++;
                                dbManager.Loguear(nombreArchivo, $"Hipervínculo roto, se eliminó la acción: {currentAddress}", "ELIMINADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🚫 {mensaje}"))));
                            }
                        }
                        else
                        {
                            vinculosSinCambios++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error en ActionSetting: {ex.Message}", "ERROR");
            }
        }

        private void ProcessTextFrameLinks(PowerPoint.Shape shape, string oldLink, string newLink, string nombreArchivo,
            ref int vinculosActualizados, ref int vinculosRotos, ref int vinculosSinCambios)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;

                    // Fix: Use the Runs method to get each run, not the property
                    for (int i = 1; i <= textRange.Runs().Count; i++)
                    {
                        try
                        {
                            var run = textRange.Runs(i);
                            if (run.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action == PowerPoint.PpActionType.ppActionHyperlink)
                            {
                                var hyperlink = run.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink;
                                string currentAddress = hyperlink.Address;

                                dbManager.Loguear(nombreArchivo, $"Comparando Texto-link: [Ruta Actual: '{currentAddress}'] con [Ruta Antigua: '{oldLink}']", "DEBUG");

                                if (!string.IsNullOrEmpty(currentAddress) &&
                                    currentAddress.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                                {
                                    string updatedAddress = Regex.Replace(currentAddress, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);

                                    bool newPathExists = File.Exists(updatedAddress) || Directory.Exists(updatedAddress);
                                    dbManager.Loguear(nombreArchivo, $"Verificando existencia de nueva ruta de texto-link: {updatedAddress} -> {(newPathExists ? "ENCONTRADO" : "NO ENCONTRADO")}", "DEBUG");

                            if (newPathExists)
                            {
                                hyperlink.Address = updatedAddress;
                                vinculosActualizados++;
                                dbManager.Loguear(nombreArchivo, $"Texto-link actualizado: {currentAddress} → {updatedAddress}", "ACTUALIZADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"✅ {mensaje}"))));
                            }
                            else
                            {
                                run.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionNone;
                                vinculosRotos++;
                                dbManager.Loguear(nombreArchivo, $"Texto-link roto, se eliminó la acción: {currentAddress}", "ELIMINADO",
                                    mensaje => Invoke((MethodInvoker)(() => listBox2.Items.Add($"🚫 {mensaje}"))));
                            }
                                }
                            }
                        }
                        catch
                        {
                            // Continuar con el siguiente run si hay error
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbManager.Loguear(nombreArchivo, $"Error en texto: {ex.Message}", "ERROR");
            }
        }

        // Método para limpiar recursos al cerrar la aplicación
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            dbManager?.Dispose();
            base.OnFormClosed(e);
        }
    }

}