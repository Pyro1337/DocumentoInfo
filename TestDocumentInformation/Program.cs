using InformacionDocumento;
string rutaCarpeta = @"C:\" + Guid.NewGuid().ToString(); //directorio de prueba
InfoDocument infodoc = new InfoDocument(rutaCarpeta);
string rutaPrueba = @"C:\rutaPruebaPaginas\Instalación local del Zcriptum.docx";
byte[] bytes = File.ReadAllBytes(rutaPrueba);
string extension = "docx";
string pathSofficeExe = @"C:\Program Files\LibreOffice\program\soffice.exe";
int? resultado = infodoc.obtenerCantPaginas(bytes, extension, pathSofficeExe);
double tamanio = infodoc.obtenerTamanioArchivo(bytes, extension);
Console.WriteLine("la cantidad de paginas es : " + resultado ?? "-1");
Console.WriteLine("el tamanio del archivo es: " + tamanio);