using iTextSharp.text;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tdm.APILF.Domain.Models;
using static iTextSharp.text.pdf.AcroFields;
using Document = Tdm.APILF.Domain.Models.Document;

namespace ValidarIntegridadPDF.ExtensionsMethods
{
    public static class WorkBookExtensions
    {
        public static int numeroFila = 0;
        

        public static ISheet CrearEncabezado(this IWorkbook wb)
        {

            // create a new worksheet
            ISheet ws = wb.CreateSheet("Documentos corruptos");

            //Create style
            //ICellStyle style = wb.CreateCellStyle();
            ////Set border style 
            ////style.BorderBottom = BorderStyle.Double;
            ////style.BottomBorderColor = HSSFColor.Black.Index;
            ////Set font style
            //IFont font = wb.CreateFont();
            ////font.Color = HSSFColor.White.Index;
            //font.FontName = "Arial";
            //font.FontHeight = 11;
            ////font.Boldweight = 2;
            //style.SetFont(font);
            //Set background color
            //style.FillBackgroundColor = IndexedColors.BlueGrey.Index;
            //style.FillPattern = FillPattern.SolidForeground;

            // create a new row
            IRow row = ws.CreateRow(numeroFila);

            // Encabezado "ID"
            ICell encabezadoID = row.CreateCell(0);
            encabezadoID.SetCellValue("ID");
            //Apply the style
            //encabezadoID.CellStyle = style;

            // Encabezado "NOMBRE"
            ICell encabezadoNombre = row.CreateCell(1);
            encabezadoNombre.SetCellValue("NOMBRE");
            //Apply the style
            //encabezadoNombre.CellStyle = style;

            // Encabezado "RUTA"
            ICell encabezadoRuta = row.CreateCell(2);
            encabezadoRuta.SetCellValue("RUTA");
            //Apply the style
            //encabezadoRuta.CellStyle = style;

            // Encabezado "FECHA DE CREACIÓN"
            ICell encabezadoFechaCreacion = row.CreateCell(3);
            encabezadoFechaCreacion.SetCellValue("FECHA DE CREACIÓN");
            //Apply the style
            //encabezadoFechaCreacion.CellStyle = style;

            return ws;
        }

        public static void AdicionarFila(this ISheet ws, Document document)
        {

            // create a new row
            IRow row = ws.CreateRow(++numeroFila);

            // create a new cell and set its value
            ICell celdaTabla = row.CreateCell(0);
            celdaTabla.SetCellValue(document.ID);

            ICell celdaCaja = row.CreateCell(1);
            celdaCaja.SetCellValue(document.Name);

            ICell celdaEncontrado = row.CreateCell(2);
            celdaEncontrado.SetCellValue(document.Path);

            ICell celdaFolderCode = row.CreateCell(3);
            celdaFolderCode.SetCellValue(document.CreationDate.ToString());
        }

        public static void GuardarReporte(this IWorkbook wb, string path)
        {
            using (var archivoResultado = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                wb.Write(archivoResultado);
            }
        }

        public static bool ExistenArchivosCorruptos(this ISheet ws) => numeroFila > 0;
    }

}
