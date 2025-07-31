using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProviderForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtOldPath.Text = @"\\fs-plaza1";
            txtNewPath.Text = @"\\nas-plaza1";

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
                    //string[] archivosExcel = Directory.GetFiles(dialog.FileName, "*.xlsx");
                    string[] archivosExcel = Directory.GetFiles(dialog.FileName, "*.xlsx", SearchOption.AllDirectories);

                    listBox1.Items.Clear();
                    foreach (string archivo in archivosExcel)
                    {
                        var excelArchivo = new ExcelArchivo(Path.GetFileName(archivo), archivo);
                        listBox1.Items.Add(excelArchivo);
                    }

                    if (archivosExcel.Length == 0)
                    {
                        MessageBox.Show("No se encontraron archivos .xlsx en la carpeta seleccionada.", "Información");
                    }else
                    {
                        lblTotalArchivos.Text = $"Total archivos encontrados: {archivosExcel.Length}";
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

            progressBar1.Minimum = 0;
            progressBar1.Maximum = totalArchivos;
            progressBar1.Value = 0;

            await Task.Run(() =>
            {
                int contador = 0;

                foreach (ExcelArchivo item in listBox1.Items)
                {
                    contador++;
                    string progreso = $"Procesando archivo {contador} de {totalArchivos}";

                    Invoke((MethodInvoker)(() =>
                    {
                        lblProgreso.Text = progreso;
                        progressBar1.Value = contador;
                    }));

                    string excelPath = item.RutaCompleta;
                    Excel.Application excelApp = null;
                    Excel.Workbook workbook = null;

                    try
                    {
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
                            foreach (var obj in links)
                            {
                                string link = obj.ToString();

                                if (link.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                                {
                                    string updatedLink = Regex.Replace(link, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);

                                    if (File.Exists(updatedLink))
                                    {
                                        workbook.ChangeLink(link, updatedLink, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                        Invoke((MethodInvoker)(() =>
                                        {
                                            listBox2.Items.Add($"{item.Nombre}: ✅ {link} → {updatedLink}");
                                        }));
                                    }
                                    else
                                    {
                                        workbook.BreakLink(link, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                        Invoke((MethodInvoker)(() =>
                                        {
                                            listBox2.Items.Add($"{item.Nombre}: ⚠️ {link} roto → vínculo eliminado, valores conservados");
                                        }));
                                    }
                                }
                                else
                                {
                                    Invoke((MethodInvoker)(() =>
                                    {
                                        listBox2.Items.Add($"{item.Nombre}: 🔗 Vínculo sin cambios: {link}");
                                    }));
                                }
                            }
                        }
                        else
                        {
                            Invoke((MethodInvoker)(() =>
                            {
                                listBox2.Items.Add($"{item.Nombre}: 📭 Sin vínculos externos");
                            }));
                        }

                        workbook.SaveAs(excelPath, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
                    }
                    catch (Exception ex)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            listBox2.Items.Add($"{item.Nombre}: ❌ ERROR - {ex.Message}");
                        }));
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
            });

            MessageBox.Show("✅ Procesamiento completado.", "Listo");
            lblProgreso.Text = "Procesamiento completado.";

        }

    }
}
