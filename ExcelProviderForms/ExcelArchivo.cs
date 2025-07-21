using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProviderForms
{
    public class ExcelArchivo
    {
        public string Nombre { get; set; }
        public string RutaCompleta { get; set; }

        public ExcelArchivo(string nombre, string rutaCompleta)
        {
            Nombre = nombre;
            RutaCompleta = rutaCompleta;
        }

        public override string ToString()
        {
            return Nombre; // Esto es lo que se mostrará en el ListBox
        }
    }
}
