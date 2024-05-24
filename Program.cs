using iTextSharp.text.pdf;
using Newtonsoft.Json;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using Tdm.APILF.Application.Services;
using Tdm.APILF.Domain.Models;
using ValidarIntegridadPDF.ExtensionsMethods;
using Document = Tdm.APILF.Domain.Models.Document;

internal class Program
{
    private static HttpClient client = new HttpClient();

    private static string exportPath = "C:\\WorkSpace\\TGI\\ValidarIntegridadPDF\\bin\\docs";

    private static string reportPath = "C:\\WorkSpace\\TGI\\ValidarIntegridadPDF\\bin\\report\\Archivos_Corruptos.xlsx";


    private static void Main(string[] args)
    {
        client.BaseAddress = new Uri("https://localhost:44395/");
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        client.Timeout = new TimeSpan(1, 0, 0);

        InitAsync().GetAwaiter().GetResult();
        Console.ReadLine();
    }

    private static async Task InitAsync()
    {
        SearchByCreationDateRequest request = new SearchByCreationDateRequest()
        {
            InitialDate = "2018/06/22 00:00:00",
            FinalDate = "2018/06/22 15:00:00"
        };

        var response = await client.PostAsJsonAsync("api/ApiLF/SearchByCreationDate", request);
        List<Document> documentos = new List<Document>();

        if (response.IsSuccessStatusCode)
        {
            var result = await response.Content.ReadAsAsync<ApiResponse<SearchResult>>();
            if(result.Succeded)
            {
                documentos = result.Data.Documents;
                await ExportarDocumentosAsync(documentos);
            }
            else
                await Console.Out.WriteLineAsync("Error la consultar la API: " + result.ErrorMessage);
        }

    }

    private static async Task ExportarDocumentosAsync(List<Document> docs)
    {
        int i = 1;

        //for (int i = 0; i< docs.Count; i++)
        //{
        //    if (i+1 % 10 == 0)
        //        Thread.Sleep(1000);

        //    var doc = docs[i];


        //    var response = await client.GetAsync($"api/ApiLF/export/{doc.ID}");
        //    if (response.IsSuccessStatusCode)
        //    {

        //        var result = await response.Content.ReadAsAsync<ApiResponse<ExportResult>>();


        //        //if (i > 0 && docs[i-1].Name == doc.Name)

        //        if (BuscarArchivo(result.Data.FileName))
        //        {
        //            /*Tamaño del nombre del documento*/
        //            var lengthDocumentName = doc.Path.Split('\\').Last().Length;

        //            /*Ruta en LF donde se encuentra el documento*/
        //            var lfDirectory = doc.Path.Substring(1, doc.Path.Length - (lengthDocumentName + 2));

        //            var newDirectory = $"{exportPath}\\{doc.Name}\\{lfDirectory}";
        //            Directory.CreateDirectory(newDirectory);
        //            File.WriteAllBytes($"{newDirectory}\\{result.Data.FileName}", result.Data.Stream);


        //        }
        //        else
        //            File.WriteAllBytes($"{exportPath}\\{result.Data.FileName}", result.Data.Stream);

        //    }
        //}

        // create a new workbook
        IWorkbook wb = new XSSFWorkbook();

        ISheet ws = wb.CrearEncabezado();


        foreach(var doc in docs)
        {
            if (i++ % 10 == 0)
                Thread.Sleep(1000);

            var response = await client.GetAsync($"api/ApiLF/export/{doc.ID}");

            if (response.IsSuccessStatusCode)
            {

                string filePath;

                var result = response.Content.ReadAsAsync<ApiResponse<ExportResult>>().Result;

                if (!result.Succeded)
                {
                    Console.WriteLine($"Error al exportar - ID {doc.ID}: {result.ErrorMessage}");
                    continue;
                }

                //if (i > 0 && docs[i-1].Name == doc.Name)

                if (BuscarArchivo(result.Data.FileName))
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
                }

                if (!ValidarIntegridad(filePath))
                {
                    ws.AdicionarFila(doc);
                }

            }
        }

        if(ws.ExistenArchivosCorruptos())
        {
            wb.GuardarReporte(reportPath);
        }
    }

    //Validar si existe un documento con el mismo nombre en la ruta de exportación
    private static bool BuscarArchivo(string file) => Directory.GetFiles(exportPath, file, SearchOption.TopDirectoryOnly).Length > 0;

    private static bool ValidarIntegridad(string file)
    {
        try
        {
            using PdfReader reader = new PdfReader(file);
        }
        catch(Exception e)
        {
            Console.WriteLine($"Error Validar Integridad  -  {file}: {e.Message}");
            return false;
        }

        return true;
    }
}