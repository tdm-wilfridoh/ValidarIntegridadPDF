using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidarIntegridadPDF.Models
{
    public class API
    {
        public string? APIBaseAddress { get; set; }
        public string? URISearchRequest { get; set; }
        public string? URIExportRequest { get; set; }
        public int SleepTime { get; set; }
    }
}
