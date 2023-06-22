using iText;
using iText.Kernel.Pdf;
using Microsoft.WindowsAPICodePack.Shell;

namespace InformacionDocumento
{
    public class InfoDocument
    {
        string rutaCarpeta;
        //inicializamos el constructor con la ruta de la carpeta auxiliar que utilizaremos.
        public InfoDocument(string rutaCarpeta) {
            this.rutaCarpeta = rutaCarpeta; 
        }
        public int obtenerCantPaginas(byte[] bytes , string extension)
        {
            if (extension == "pdf") //si la extension es .pdf 
            {
                if (!Directory.Exists(rutaCarpeta))
                {
                    Directory.CreateDirectory(rutaCarpeta);//sino existe el directorio carpetaPrueba lo creara.

                }
                string rutaArchivoNuevo = Path.Combine(rutaCarpeta, "archivoaux." + extension);//destino y nombre del archivoPrueba acoplamos con la extension en este caso pdf
                File.WriteAllBytes(rutaArchivoNuevo, bytes);//convierte byte a formato PDF lleva como parametro
                var pdfDocument = new PdfDocument(reader: new PdfReader(rutaArchivoNuevo)); //usamos el pdfreader con la nueva ruta
                int cantPaginas = pdfDocument.GetNumberOfPages();//extraemos la cantidad de paginas con la funcion que contiene dicha libreria
                pdfDocument.Close();//cerramos el pdfDocument para limpiar el buffer y posteriormente
                eliminarArchivoCarpeta(rutaArchivoNuevo);
                return cantPaginas;
            }
            else
            {
                return -1;
            }
        }
        public double obtenerTamanioArchivo(byte[] bytes , string extension)
        {
            if (!Directory.Exists(rutaCarpeta))
            {
                Directory.CreateDirectory(rutaCarpeta);//sino existe el directorio carpetaPrueba lo creara.

            }
            string rutaArchivoNuevo = Path.Combine(rutaCarpeta, "archivoaux." + extension);//destino y nombre del archivoPrueba acoplamos con la extension en este caso pdf
            File.WriteAllBytes(rutaArchivoNuevo, bytes);//convierte byte a formato correspondiente
            FileInfo fi = new FileInfo(rutaArchivoNuevo);
            string tamaniokb = fi.Length.ToString();
            double resultadoKB = double.Parse(tamaniokb)/1000;
            eliminarArchivoCarpeta(rutaArchivoNuevo);
            return resultadoKB;
            
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