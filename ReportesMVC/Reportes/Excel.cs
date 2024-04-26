using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO.Compression;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text;

namespace ReportesMVC.Reportes
{
    public class Excel
    {
        public static string EnviarCorreo(string correoenvia, List<string> correosAEnviar, string asunto,
        string contenido, List<string> nombresArchivos = null, List<byte[]> listabyte = null)
        {
            string str;
            try
            {
                string Host = "Smtp.Gmail.com";
                string appSetting2 = "587";//587
                string appSetting3 = "wapcamargo@gmail.com";
                string password = "Samupi7185467*";
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential(appSetting3, password);
                smtpClient.Port = int.Parse(appSetting2);
                smtpClient.Host = Host;
                smtpClient.EnableSsl = true;
                MailMessage message = new MailMessage();
                message.From = new MailAddress(correoenvia);
                foreach (string address in correosAEnviar)
                    message.To.Add(new MailAddress(address));
                if (nombresArchivos != null && listabyte != null)
                {
                    for (int i = 0; i < nombresArchivos.Count; i++)
                    {
                        MemoryStream memoryStream = new MemoryStream(listabyte[i]);
                        message.Attachments.Add(new Attachment(memoryStream, nombresArchivos[i],
                            MediaTypeNames.Application.Octet));
                    }
                }
                message.Subject = asunto;
                message.IsBodyHtml = true;
                message.Body = contenido;
                smtpClient.Send(message);
                str = "Se envio el Correo satisfactoriamente";
            }
            catch (Exception ex)
            {
                str = "Error al Enviar Correo: " + ex.Message;
            }
            return str;
        }


        public static void tituloHorizontal(ExcelWorksheet ew1, string titulo = "Reporte", int postFila = 1,
            int posInicioColumna = 1, int postFinColumna = 4, int fuenteTexto = 20,
            Color? fondo = null, Color? colortexto = null)
        {
            using (ExcelRange rango = ew1.Cells[postFila, posInicioColumna, postFila, postFinColumna])
            {
                rango.Merge = true;
                rango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rango.Style.Font.Size = fuenteTexto;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                if (fondo == null)
                    rango.Style.Fill.BackgroundColor.SetColor(Color.Teal);
                else
                    rango.Style.Fill.BackgroundColor.SetColor((Color)fondo);
                if (colortexto == null)
                    rango.Style.Font.Color.SetColor(Color.White);
                else
                    rango.Style.Font.Color.SetColor((Color)colortexto);
            }
            ew1.Cells[postFila, posInicioColumna].Value = titulo;
        }

        public static void tituloVertical(ExcelWorksheet ew1, string titulo = "Reporte", int postColumna = 1,
            int posInicioFila = 1, int postFinFila = 4, int fuenteTexto = 20,
            Color? fondo = null, Color? colortexto = null)
        {
            using (ExcelRange rango = ew1.Cells[posInicioFila, postColumna, postFinFila, postColumna])
            {
                rango.Merge = true;
                rango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rango.Style.Font.Size = fuenteTexto;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                if (fondo == null)
                    rango.Style.Fill.BackgroundColor.SetColor(Color.Teal);
                else
                    rango.Style.Fill.BackgroundColor.SetColor((Color)fondo);
                if (colortexto == null)
                    rango.Style.Font.Color.SetColor(Color.White);
                else
                    rango.Style.Font.Color.SetColor((Color)colortexto);
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < titulo.Length; i++)
            {
                sb.AppendLine(titulo.Substring(i, 1));
            }
            ew1.Cells[posInicioFila, postColumna].Value = sb.ToString();
            ew1.Cells[posInicioFila, postColumna].Style.WrapText = true;
        }

        public static void anchosColumnas(ExcelWorksheet ew1, int postInicioColumna = 1,
            List<int> anchos = null)
        {
            int postFinColumna = postInicioColumna + anchos.Count - 1;
            int indiceAncho = 0;
            for (int i = postInicioColumna; i <= postFinColumna; i++)
            {
                ew1.Column(i).Width = anchos[indiceAncho];
                indiceAncho++;
            }


        }

        public static void cabecerasTabla(ExcelWorksheet ew1, int posFila, int postInicioColumna,
            List<string> cabeceras, bool bordes = true, bool centrar = true, bool negrita = true)
        {
            int postFinColumna = postInicioColumna + cabeceras.Count - 1;
            int indicecabecera = 0;
            for (int i = postInicioColumna; i <= postFinColumna; i++)
            {
                ew1.Cells[posFila, i].Value = cabeceras[indicecabecera];
                if (negrita == true)
                    ew1.Cells[posFila, i].Style.Font.Bold = true;
                if (bordes == true)
                    Excel.border(ew1, posFila, i);
                if (centrar == true)
                    ew1.Cells[posFila, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                indicecabecera++;
            }
        }

        public static void filasTabla<T>(ExcelWorksheet ew1, List<T> data, int posFila, int postInicioColumna,
            List<string> campos, bool bordes = true, bool centrar = true, bool negrita = true)
        {
            int postFinColumna = postInicioColumna + campos.Count - 1;

            for (int j = 0; j < data.Count; j++)
            {
                int indicecampos = 0;
                for (int i = postInicioColumna; i <= postFinColumna; i++)
                {
                    ew1.Cells[posFila, i].Value = data[j].GetType().GetProperty(campos[indicecampos]).GetValue(data[j], null).ToString();
                    indicecampos++;
                }
                posFila++;
            }
            /*
			for (int i = 0; i < listaPersona.Count; i++)
			{
				ew1.Cells[iniciofila, 1].Value = listaPersona[i].Nombre;
				ew1.Cells[iniciofila, 2].Value = listaPersona[i].Appaterno;
				ew1.Cells[iniciofila, 3].Value = listaPersona[i].Apmaterno;
				ew1.Cells[iniciofila, 4].Value = listaPersona[i].Telefonoocelular1;
				iniciofila++;
			}*/
        }

        public static void border(ExcelWorksheet ew1, int fila, int columna)
        {
            ew1.Cells[fila, columna].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ew1.Cells[fila, columna].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ew1.Cells[fila, columna].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ew1.Cells[fila, columna].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        public static void cabecerasFila(ExcelWorksheet ew1, int posColumna, int postInicioFila,
        List<string> cabeceras, bool bordes = true, bool centrar = true, bool negrita = true)
        {
            int postFinFila = postInicioFila + cabeceras.Count - 1;
            int indicecabecera = 0;
            for (int i = postInicioFila; i <= postFinFila; i++)
            {
                ew1.Cells[i, posColumna].Value = cabeceras[indicecabecera];
                if (negrita == true)
                    ew1.Cells[i, posColumna].Style.Font.Bold = true;
                if (bordes == true)
                    Excel.border(ew1, i, posColumna);
                if (centrar == true)
                    ew1.Cells[i, posColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                indicecabecera++;
            }
        }
        public static void objetoFila<T>(ExcelWorksheet ew1, T data, int posColumna, int postInicioFila,
            List<string> campos, bool bordes = true, bool centrar = true, bool negrita = false)
        {
            int postFinFila = postInicioFila + campos.Count - 1;


            int indicecampos = 0;
            for (int i = postInicioFila; i <= postFinFila; i++)
            {
                ew1.Cells[i, posColumna].Value =
                data.GetType().GetProperty(campos[indicecampos]).GetValue(data, null) == null ? "" :
                data.GetType().GetProperty(campos[indicecampos]).GetValue(data, null).ToString();
                if (bordes == true)
                    border(ew1, i, posColumna);
                if (centrar == true)
                    ew1.Cells[i, posColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ew1.Cells[i, posColumna].Style.Font.Bold = negrita;
                indicecampos++;
            }


        }

        public static string comprimirExcel(List<string> listanombres, List<byte[]> listaarraybyte)
        {
            byte[] compressedBytes;
            using (var outStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(outStream, ZipArchiveMode.Create, true))
                {
                    for (int i = 0; i < listaarraybyte.Count; i++)
                    {
                        var fileInArchive = archive.CreateEntry(listanombres[i] + ".xlsx",
                            System.IO.Compression.CompressionLevel.Optimal);
                        using (var entryStream = fileInArchive.Open())
                        using (var fileToCompressStream = new MemoryStream(listaarraybyte[i]))
                        {
                            fileToCompressStream.CopyTo(entryStream);
                        }
                    }
                }
                compressedBytes = outStream.ToArray();
                return "data:application/zip;base64," + Convert.ToBase64String(compressedBytes);
            }
        }
    }
}
