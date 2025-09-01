using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace ExcelProviderForms
{
    public class DatabaseManager : IDisposable
    {
        private readonly string dbPath;
        private readonly string connectionString;
        private bool disposed = false;

        public DatabaseManager(string fileName = "procesamiento.db")
        {
            dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            connectionString = $"Data Source={dbPath};Version=3;";
            InicializarBase();
        }

        private void InicializarBase()
        {
            if (!File.Exists(dbPath))
            {
                SQLiteConnection.CreateFile(dbPath);
            }

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                string sqlArchivos = @"CREATE TABLE IF NOT EXISTS ArchivosProcesados (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    nombre TEXT,
                                    ruta TEXT UNIQUE,
                                    procesado INTEGER,
                                    fecha_procesado TEXT,
                                    fecha_creado TEXT DEFAULT CURRENT_TIMESTAMP
                                  );";

                string sqlLogs = @"CREATE TABLE IF NOT EXISTS Logs (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    archivo TEXT,
                                    mensaje TEXT,
                                    estado TEXT,
                                    fecha TEXT DEFAULT CURRENT_TIMESTAMP
                               );";

                using (var cmd = new SQLiteCommand(sqlArchivos, conn))
                    cmd.ExecuteNonQuery();

                using (var cmd = new SQLiteCommand(sqlLogs, conn))
                    cmd.ExecuteNonQuery();
            }
        }

        public bool EstaProcesado(string ruta)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT procesado FROM ArchivosProcesados WHERE ruta = @ruta", conn))
                {
                    cmd.Parameters.AddWithValue("@ruta", ruta);
                    var result = cmd.ExecuteScalar();
                    return result != null && Convert.ToInt32(result) == 1;
                }
            }
        }

        public void RegistrarArchivo(string nombre, string ruta)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "INSERT OR IGNORE INTO ArchivosProcesados (nombre, ruta, procesado, fecha_creado) VALUES (@nombre, @ruta, 0, datetime('now'))", conn))
                {
                    cmd.Parameters.AddWithValue("@nombre", nombre);
                    cmd.Parameters.AddWithValue("@ruta", ruta);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void MarcarProcesado(string ruta)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "UPDATE ArchivosProcesados SET procesado = 1, fecha_procesado = datetime('now') WHERE ruta = @ruta", conn))
                {
                    cmd.Parameters.AddWithValue("@ruta", ruta);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void Loguear(string archivo, string mensaje, string estado, Action<string> mostrarEnUI = null)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            // Mostrar en UI si se pasa un método delegado
            mostrarEnUI?.Invoke($"{archivo}: {mensaje}");

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "INSERT INTO Logs (archivo, mensaje, estado, fecha) VALUES (@archivo, @mensaje, @estado, datetime('now'))", conn))
                {
                    cmd.Parameters.AddWithValue("@archivo", archivo);
                    cmd.Parameters.AddWithValue("@mensaje", mensaje);
                    cmd.Parameters.AddWithValue("@estado", estado);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void MarcarFallido(string ruta)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "UPDATE ArchivosProcesados SET procesado = 1, fecha_procesado = datetime('now') WHERE ruta = @ruta", conn))
                {
                    cmd.Parameters.AddWithValue("@ruta", ruta);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // Método adicional para obtener estadísticas
        public (int total, int procesados, int pendientes) ObtenerEstadisticas()
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM ArchivosProcesados", conn))
                {
                    int total = Convert.ToInt32(cmd.ExecuteScalar());

                    using (var cmd2 = new SQLiteCommand("SELECT COUNT(*) FROM ArchivosProcesados WHERE procesado = 1", conn))
                    {
                        int procesados = Convert.ToInt32(cmd2.ExecuteScalar());
                        return (total, procesados, total - procesados);
                    }
                }
            }
        }

        // Método para limpiar logs antiguos (opcional)
        public void LimpiarLogsAntiguos(int diasAntiguedad = 30)
        {
            if (disposed) throw new ObjectDisposedException(nameof(DatabaseManager));

            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "DELETE FROM Logs WHERE fecha < datetime('now', '-' || @dias || ' days')", conn))
                {
                    cmd.Parameters.AddWithValue("@dias", diasAntiguedad);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
            }
        }

        ~DatabaseManager()
        {
            Dispose(false);
        }
    }
}