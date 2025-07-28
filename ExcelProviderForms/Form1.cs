using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProviderForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtOldPath.Text = @"C:\Nas";
            txtNewPath.Text = @"C:\FsPlaza";
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
                    string[] archivosExcel = Directory.GetFiles(dialog.FileName, "*.xlsx");

                    listBox1.Items.Clear();
                    foreach (string archivo in archivosExcel)
                    {
                        var excelArchivo = new ExcelArchivo(Path.GetFileName(archivo), archivo);
                        listBox1.Items.Add(excelArchivo);
                    }

                    if (archivosExcel.Length == 0)
                    {
                        MessageBox.Show("No se encontraron archivos .xlsx en la carpeta seleccionada.", "Información");
                    }
                }
            }
        }


        private void btnProcesar_Click(object sender, EventArgs e)
        {
            string oldLink = txtOldPath.Text.Trim();
            string newLink = txtNewPath.Text.Trim();
            bool breakLinkIfNotFound = true;

            if (string.IsNullOrWhiteSpace(oldLink) || string.IsNullOrWhiteSpace(newLink))
            {
                MessageBox.Show("Por favor ingresa ambas rutas: antigua y nueva.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            listBox2.Items.Clear();

            foreach (ExcelArchivo item in listBox1.Items)
            {
                string excelPath = item.RutaCompleta;
                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;

                try
                {
                    excelApp = new Excel.Application();
                    excelApp.DisplayAlerts = false;
                    excelApp.AskToUpdateLinks = false;
                    excelApp.AlertBeforeOverwriting = false;

                    workbook = excelApp.Workbooks.Open(excelPath, UpdateLinks: 0, ReadOnly: false);
                    var rawLinks = workbook.LinkSources(Excel.XlLink.xlExcelLinks);

                    if (rawLinks is Array links)
                    {
                        foreach (var obj in links)
                        {
                            string link = obj.ToString();

                            if (link.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                            {
                                string updatedLink = link.Replace(oldLink, newLink);

                                if (File.Exists(updatedLink))
                                {
                                    workbook.ChangeLink(link, updatedLink, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                    listBox2.Items.Add($"{item.Nombre}: ✅ {link} → {updatedLink}");
                                }
                                else if (breakLinkIfNotFound)
                                {
                                    workbook.BreakLink(link, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                    listBox2.Items.Add($"{item.Nombre}: ⚠️ {link} roto (archivo no encontrado en {updatedLink})");
                                }
                                else
                                {
                                    listBox2.Items.Add($"{item.Nombre}: ⚠️ {link} no encontrado, vínculo original mantenido");
                                }
                            }
                            else
                            {
                                listBox2.Items.Add($"{item.Nombre}: 🔗 Vínculo sin cambios: {link}");
                            }
                        }
                    }
                    else
                    {
                        listBox2.Items.Add($"{item.Nombre}: 📭 Sin vínculos externos");
                    }

                    workbook.SaveAs(excelPath, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
                }
                catch (Exception ex)
                {
                    listBox2.Items.Add($"{item.Nombre}: ❌ ERROR - {ex.Message}");
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
                }
            }

            MessageBox.Show("✅ Procesamiento completado.", "Listo");
        }

    }
}
