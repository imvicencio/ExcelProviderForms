using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ExcelProviderForms
{
    public class PowerPointArchivo
    {
        public string Nombre { get; set; }
        public string RutaCompleta { get; set; }
        public List<string> Hipervinculos { get; set; }

        public PowerPointArchivo(string nombre, string rutaCompleta)
        {
            Nombre = nombre;
            RutaCompleta = rutaCompleta;
            Hipervinculos = new List<string>();
        }

        public void Procesar()
        {
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                pptApp = new PowerPoint.Application();
                presentation = AbrirConReintentos(pptApp, RutaCompleta);

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // 1. Intentar obtener Hyperlinks directos
                        try
                        {
                            foreach (PowerPoint.Hyperlink hl in shape.Hyperlinks)
                            {
                                if (!string.IsNullOrEmpty(hl.Address))
                                {
                                    Hipervinculos.Add(hl.Address);
                                }
                            }
                        }
                        catch
                        {
                            // Algunos shapes no tienen Hyperlinks → ignorar
                        }

                        // 2. Intentar obtener ActionSettings (si aplica)
                        try
                        {
                            if (shape.ActionSettings != null)
                            {
                                var action = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                                if (action != null && action.Hyperlink != null &&
                                    !string.IsNullOrEmpty(action.Hyperlink.Address))
                                {
                                    Hipervinculos.Add(action.Hyperlink.Address);
                                }
                            }
                        }
                        catch
                        {
                            // No todos los shapes soportan ActionSettings → ignorar
                        }
                    }
                }
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
            }
        }

        private PowerPoint.Presentation AbrirConReintentos(PowerPoint.Application pptApp, string ruta)
        {
            int retries = 3;
            while (true)
            {
                try
                {
                    return pptApp.Presentations.Open(ruta, WithWindow: MsoTriState.msoFalse);
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x8001010A) // RPC_E_SERVERCALL_RETRYLATER
                {
                    retries--;
                    if (retries <= 0)
                        throw;
                    System.Threading.Thread.Sleep(500); // Esperar medio segundo y reintentar
                }
            }
        }

        public override string ToString()
        {
            return Nombre;
        }
    }

    // Requerido para el flag WithWindow
    public enum MsoTriState
    {
        msoFalse = 0,
        msoTrue = -1,
        msoCTrue = 1,
        msoTriStateToggle = -3
    }
}
