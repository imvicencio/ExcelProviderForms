using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProviderForms
{
    public class ExcelArchivo : ArchivoOffice
    {
        public ExcelArchivo(string nombre, string rutaCompleta) : base(nombre, rutaCompleta) { }

        public override void Procesar(string oldLink, string newLink, DatabaseManager dbManager, Action<string> log)
        {
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
                    RutaCompleta,
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
                                log($"✅ Vínculo actualizado en Excel: {link} → {updatedLink}");
                            }
                            else
                            {
                                workbook.BreakLink(link, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                log($"⚠️ Vínculo roto eliminado en Excel: {link}");
                            }
                        }
                        else
                        {
                            log($"🔗 Vínculo sin cambios en Excel: {link}");
                        }
                    }
                }
                else
                {
                    log("📭 Sin vínculos externos en Excel");
                }

                workbook.SaveAs(RutaCompleta, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
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
    }
}
