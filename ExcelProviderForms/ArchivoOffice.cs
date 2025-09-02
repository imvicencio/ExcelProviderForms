using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProviderForms
{
    public abstract class ArchivoOffice
    {
        public string Nombre { get; set; }
        public string RutaCompleta { get; set; }

        protected ArchivoOffice(string nombre, string rutaCompleta)
        {
            Nombre = nombre;
            RutaCompleta = rutaCompleta;
        }

        // Procesamiento de vínculos (implementado en cada derivada)
        public abstract void Procesar(string oldLink, string newLink, DatabaseManager dbManager, Action<string> log);

        // Lo que se muestra en el ListBox
        public override string ToString()
        {
            return Nombre;
        }
    }
}
