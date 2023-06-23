using ConvertToPDF;
using System.Text.RegularExpressions;

namespace InformacionDocumento
{
    public class InfoDocument
    {
        string rutaCarpeta;
        //inicializamos el constructor con la ruta de la carpeta auxiliar que utilizaremos.
        public InfoDocument(string rutaCarpeta)
        {
            this.rutaCarpeta = rutaCarpeta;
        }
        public int? obtenerCantPaginas(byte[] bytes, string extension, string pathSofficeExe)
        {
            string? rutaArchivoNuevo = null;
            if (
                    extension.Equals("docx", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("xls", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("xlsx", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("csv", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("pptx", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("ppt", StringComparison.OrdinalIgnoreCase)
                )
            {
                ConvertToPDF.ConvertToPDF convert = new ConvertToPDF.ConvertToPDF();//referencia interna al dll de conversion a pdf
                rutaArchivoNuevo = convert.OfficeToPDF(bytes, extension, pathSofficeExe);
            }
            else if (extension.Equals("pdf", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    crearCarpeta(rutaCarpeta);
                    string nameFile = Guid.NewGuid().ToString();//usamos el guid para darle un nombre aleatorio
                    rutaArchivoNuevo = Path.Combine(rutaCarpeta, nameFile + "." + extension);//mediante el pathcombine acoplamos la ruta de la carpeta con el nombre del archivo junto con su extension
                    File.WriteAllBytes(rutaArchivoNuevo, bytes);//convierte byte a formato PDF
                }
                catch (Exception)
                {
                    return null;
                }
            }
            if (rutaArchivoNuevo == null)
            {
                return null;
            }
            //var pdfDocument = new PdfDocument(reader: new PdfReader(rutaArchivoNuevo)); //usamos el pdfreader con la nueva ruta
            //int cantPaginas = pdfDocument.GetNumberOfPages();//extraemos la cantidad de paginas con la funcion que contiene dicha libreria
            //pdfDocument.Close();//cerramos el pdfDocument para limpiar el buffer
            int cantPaginas = extraerNumeroPaginas(rutaArchivoNuevo);
            eliminarArchivoCarpeta(rutaArchivoNuevo);
            return cantPaginas;
        }
        public double obtenerTamanioArchivo(byte[] bytes, string extension)
        {
            crearCarpeta(rutaCarpeta);
            string nombreArchivo = Guid.NewGuid().ToString();//usamos el guid para darle un nombre aleatorio
            string rutaArchivoNuevo = Path.Combine(rutaCarpeta, nombreArchivo + "." + extension);//destino y nombre del archivoPrueba acoplamos con la extension en este caso pdf
            File.WriteAllBytes(rutaArchivoNuevo, bytes);//convierte byte a formato correspondiente
            FileInfo fi = new FileInfo(rutaArchivoNuevo);
            string tamaniokb = fi.Length.ToString();
            double resultadoKB = double.Parse(tamaniokb) / 1024;
            eliminarArchivoCarpeta(rutaArchivoNuevo);
            return resultadoKB;
        }

        public int extraerNumeroPaginas(string rutaArchivoNuevo)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(rutaArchivoNuevo)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());
                return matches.Count;
            }
        }

        private void crearCarpeta(string rutaCarpeta)
        {
            if (!Directory.Exists(rutaCarpeta))
            {
                Directory.CreateDirectory(rutaCarpeta);//sino existe el directorio carpetaPrueba lo creara.
            }
        }

        private void eliminarArchivoCarpeta(string rutaArchivo)
        {
            if (Directory.Exists(rutaCarpeta))
            {
                File.Delete(rutaArchivo);//en caso de existir dicho directorio elimina el contenido
                Directory.Delete(rutaCarpeta);//posteriormente elimina la carpeta
            }
        }
    }
}