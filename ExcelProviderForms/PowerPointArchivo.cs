namespace ExcelProviderForms
{
    public class PowerPointArchivo
    {
        public string Nombre { get; set; }
        public string RutaCompleta { get; set; }

        public PowerPointArchivo(string nombre, string rutaCompleta)
        {
            Nombre = nombre;
            RutaCompleta = rutaCompleta;
        }

        public override string ToString()
        {
            return Nombre;
        }
    }
}
