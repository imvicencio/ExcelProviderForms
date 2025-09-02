using System;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ExcelProviderForms
{
    public class PowerPointArchivo : ArchivoOffice
    {
        public PowerPointArchivo(string nombre, string rutaCompleta) : base(nombre, rutaCompleta) { }

        public override void Procesar(string oldLink, string newLink, DatabaseManager dbManager, Action<string> log)
        {
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                pptApp = new PowerPoint.Application();
                presentation = pptApp.Presentations.Open(RutaCompleta,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        var hyperlink = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink;
                        if (hyperlink != null && !string.IsNullOrEmpty(hyperlink.Address))
                        {
                            string link = hyperlink.Address;

                            if (link.StartsWith(oldLink, StringComparison.OrdinalIgnoreCase))
                            {
                                string updatedLink = Regex.Replace(link, Regex.Escape(oldLink), newLink, RegexOptions.IgnoreCase);
                                hyperlink.Address = updatedLink;
                                log($"✅ Vínculo actualizado en PPT: {link} → {updatedLink}");
                            }
                            else
                            {
                                log($"🔗 Vínculo sin cambios en PPT: {link}");
                            }
                        }
                    }
                }

                presentation.Save();
            }
            finally
            {
                if (presentation != null) presentation.Close();
                if (pptApp != null) pptApp.Quit();
            }
        }
    }
}
