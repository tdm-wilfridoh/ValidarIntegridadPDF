using iTextSharp.text.pdf;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using Tdm.APILF_Net6.Domain.Models;
using ValidarIntegridadPDF.ExtensionsMethods;
using Document = Tdm.APILF_Net6.Domain.Models.Document;
using log4net.Config;
using log4net;
using Microsoft.Extensions.Configuration;
using ValidarIntegridadPDF.Models;
using NPOI.SS.Formula.Functions;

internal class Program
{
    private static HttpClient client = new HttpClient();

    private static string? exportPath;

    private static string? reportPath;

    private static string? reportName;

    private static ILog log = LogManager.GetLogger(typeof(Program));

    private static API _api;
    private static SearchDates _searchDates;

    private static void Main(string[] args)
    {
        try
        {

            XmlConfigurator.Configure(new FileInfo("log4net.config"));

            var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("config.json", optional: true, reloadOnChange: true);

            IConfiguration configuration = builder.Build();
            _api = configuration.GetSection("API").Get<API>();
            _searchDates = configuration.GetSection("SearchDates").Get<SearchDates>();

            exportPath = configuration.GetSection("ExportPath").Value;
            reportPath = configuration.GetSection("ReportPath").Value;
            reportName = configuration.GetSection("ReportName").Value;
            reportName = String.Format(reportName, DateTime.Now.ToString("dd/MM/yyyy").Replace('/', '-'));



            client.BaseAddress = new Uri(_api.APIBaseAddress);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.Timeout = new TimeSpan(1, 0, 0);

            InitAsync().GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            log.Error("Error general: " + ex.Message);
        }
        Console.ReadLine();
    }

    static async Task InitAsync()
    {
        log.Debug("------------ Iniciando proceso ------------");

        if (Directory.Exists(exportPath))
            DeleteExportPath();


        Directory.CreateDirectory(exportPath);


        SearchByCreationDateRequest request = new SearchByCreationDateRequest()
        {
            InitialDate = _searchDates.Initial,
            FinalDate = _searchDates.Final
        };

        log.Debug($"Consultando al repositorio: Fecha inicial = {_searchDates.Initial},  Fecha final = {_searchDates.Final}");
        var response = await client.PostAsJsonAsync(_api.URISearchRequest, request);
        List<Document> documentos = new List<Document>();

        if (response.IsSuccessStatusCode)
        {
            var result = await response.Content.ReadAsAsync<ApiResponse<SearchResult>>();
            if (result.Succeded)
            {
                documentos = result.Data.Documents;
                log.Debug($"Total de documentos encontrados: {documentos.Count}");
                await ExportarDocumentosAsync(documentos);
            }
            else
                await Console.Out.WriteLineAsync("Error la consultar la API: " + result.ErrorMessage);
        }

    }

    static async Task ExportarDocumentosAsync(List<Document> docs)
    {
        int i = 1;

        // create a new workbook
        IWorkbook wb = new XSSFWorkbook();

        ISheet ws = wb.CrearEncabezado();

        log.Debug("Iniciando proceso de exportación de documentos y validación de integridad");
        foreach (var doc in docs)
        {
            try
            {
                if (i++ % 10 == 0)
                    Thread.Sleep(_api.SleepTime);

                log.Debug($"==== Procesando {doc.ID} - {doc.Name} ====");
                var response = await client.GetAsync($"{_api.URIExportRequest}/{doc.ID}");

                if (response.IsSuccessStatusCode)
                {

                    string filePath;

                    var result = response.Content.ReadAsAsync<ApiResponse<ExportResult>>().Result;

                    if (!result.Succeded)
                    {
                        Console.WriteLine($"Error al exportar - ID {doc.ID}: {result.ErrorMessage}");
                        log.Error($"Error al exportar: {result.ErrorMessage}");
                        ws.AdicionarFila(doc, result.ErrorMessage);
                        continue;
                    }

                    //if (i > 0 && docs[i-1].Name == doc.Name)

                    if (BuscarArchivo(exportPath, result.Data.FileName) > 0)
                    {
                        /*Tamaño del nombre del documento*/
                        var lengthDocumentName = doc.Path.Split('\\').Last().Length;

                        /*Ruta en LF donde se encuentra el documento*/
                        var lfDirectory = doc.Path.Substring(1, doc.Path.Length - (lengthDocumentName + 2));

                        var newDirectory = $"{exportPath}\\{doc.Name}\\{lfDirectory}";
                        Directory.CreateDirectory(newDirectory);
                        filePath = $"{newDirectory}\\{result.Data.FileName}";
                        File.WriteAllBytes(filePath, result.Data.Stream);
                    }
                    else
                    {
                        filePath = $"{exportPath}\\{result.Data.FileName}";
                        File.WriteAllBytes(filePath, result.Data.Stream);
                        log.Debug("Archivo exportado correctamente");
                    }

                    bool integrity = ValidarIntegridad(filePath);

                    if (!integrity)
                    {
                        ws.AdicionarFila(doc, "Error de integridad");
                    }

                    DeleteFile(filePath);

                    log.Debug("Prueba de integridad: " + (integrity ? "Aprobada" : "No aprobada"));
                }
                else
                {
                    log.Error($"Error al requerir la API: {response.ReasonPhrase}");
                }

                log.Debug($"==== Fin de proceso {doc.ID} - {doc.Name}  ====");
            }
            catch(Exception e)
            {
                ws.AdicionarFila(doc, e.Message);
                log.Error($"Export error {doc.ID} - {doc.Name}: " + e.Message);
            }

        }

        DeleteExportPath();

        log.Debug("Fin proceso de exportación de documentos y validación de integridad");

        var corruptFiles = ws.TotalArchivosCorruptos();

        log.Debug($"Número de documentos procesados: {i-1}");

        log.Debug($"Cantidad de archivos corruptos: {corruptFiles}");

        if (corruptFiles > 0)
        {
            var numberDocuments = BuscarArchivo(reportPath, reportName + "*");
            wb.GuardarReporte(reportPath + reportName + (numberDocuments > 0 ? $"({numberDocuments})" : "") + ".xlsx");
            log.Debug("Archivo de reporte generado correctamente");
        }

    }

    //Validar cuántos documentos "fileName" existen con el mismo nombre en la ruta "path". Si no existen duplicados devuleve 0
    static int BuscarArchivo(string path, string fileName) => Directory.GetFiles(path, fileName, SearchOption.TopDirectoryOnly).Length;


    static bool ValidarIntegridad(string file)
    {
        try
        {
            using PdfReader reader = new PdfReader(file);
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error Validar Integridad  -  {file}: {e.Message}");
            log.Error($"Error al validar integridad: {e.Message}");
            return false;
        }

        return true;
    }

    static void DeleteFile(string file) => File.Delete(file);

    static void DeleteExportPath() => Directory.Delete(exportPath, true);

}