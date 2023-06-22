using InformacionDocumento;
string rutaCarpeta = @"C:\carpetaPrueba"; //directorio de prueba
InfoDocument infodoc = new InfoDocument(rutaCarpeta);
string rutaPrueba = @"C:\rutaPruebaPaginas\redes.pdf";
byte[] bytes = File.ReadAllBytes(rutaPrueba);
string extension = "pdf";
int resultado = infodoc.obtenerCantPaginas(bytes,extension);
double tamanio = infodoc.obtenerTamanioArchivo(bytes,extension);
Console.WriteLine("la cantidad de paginas es : "+resultado);
Console.WriteLine("el tamanio del archivo es: "+tamanio);
