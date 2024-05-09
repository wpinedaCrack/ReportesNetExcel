using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportesMVC.Models;
using ReportesMVC.Reportes;

namespace ReportesMVC.Controllers
{
    public class PersonaController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ExcelCheck()
        {
            return View();
        }
        public IActionResult ExcelCheckMultiplesHojas()
        {
            return View();
        }
        public IActionResult ExcelCheckComprimido()
        {
            return View();
        }
        public IActionResult CorreoEnvioExcel()
        {
            return View();
        }

        public IActionResult PreviewExcel()
        {
            return View();
        }

        public string leerExcel(IFormFile excel)
        {
            string rpta = "", negrita = "", bordeTop = "", bordeBottom = "", bordeRight = "", bordeLeft = "",
                horizontalCenter = "", horizontalRight = "";

            byte[] buffer;
            using (MemoryStream ms = new MemoryStream())
            {
                excel.CopyTo(ms);
                buffer = ms.ToArray();
            }
            using (MemoryStream ms = new MemoryStream(buffer))
            {
                using (ExcelPackage ep = new ExcelPackage(ms))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var ew1 = ep.Workbook.Worksheets[0];
                    int ncolumnas = ew1.Dimension.End.Column;
                    int nfilas = ew1.Dimension.End.Row;
                    //    | -> Columnas
                    //    ¬ ->Filas
                    for (int i = 1; i <= nfilas; i++)
                    {
                        for (int j = 1; j <= ncolumnas; j++)
                        {
                            if (ew1.Cells[i, j].Style.Font.Bold == true) negrita += "n" + i + "¬" + j + "|";
                            rpta += ew1.Cells[i, j].Value;
                            if (ew1.Cells[i, j].Style.Border.Top.Style != ExcelBorderStyle.None)
                                bordeTop += "bt" + i + "¬" + j + "|";
                            if (ew1.Cells[i, j].Style.Border.Bottom.Style != ExcelBorderStyle.None)
                                bordeBottom += "bb" + i + "¬" + j + "|";
                            if (ew1.Cells[i, j].Style.Border.Right.Style != ExcelBorderStyle.None)
                                bordeRight += "br" + i + "¬" + j + "|";
                            if (ew1.Cells[i, j].Style.Border.Left.Style != ExcelBorderStyle.None)
                                bordeLeft += "bl" + i + "¬" + j + "|";
                            if (ew1.Cells[i, j].Style.HorizontalAlignment == ExcelHorizontalAlignment.Center)
                                horizontalCenter += "hc" + i + "¬" + j + "|";
                            if (ew1.Cells[i, j].Style.HorizontalAlignment == ExcelHorizontalAlignment.Right)
                                horizontalRight += "hr" + i + "¬" + j + "|";
                            rpta += "|";
                        }
                        rpta = rpta.Substring(0, rpta.Length - 1);
                        rpta += "¬";
                    }
                    rpta = rpta.Substring(0, rpta.Length - 1);
                    rpta += "_";
                    //Estilos
                    if (negrita != "") rpta += negrita.Substring(0, negrita.Length - 1) + "";
                    if (bordeTop != "") rpta += "|" + bordeTop.Substring(0, bordeTop.Length - 1) + "";
                    if (bordeBottom != "") rpta += "|" + bordeBottom.Substring(0, bordeBottom.Length - 1) + "";
                    if (bordeRight != "") rpta += "|" + bordeRight.Substring(0, bordeRight.Length - 1) + "";
                    if (bordeLeft != "") rpta += "|" + bordeLeft.Substring(0, bordeLeft.Length - 1) + "";
                    if (horizontalCenter != "") rpta += "|" + horizontalCenter.Substring(0,
                        horizontalCenter.Length - 1) + "";
                    if (horizontalRight != "") rpta += "|" + horizontalRight.Substring(0,
                        horizontalRight.Length - 1) + "";

                }
            }
            return rpta;
        }


        public string enviarCorreo(int id, string correo)
        {
            string cadenabase = generarReportePorId(id);
            byte[] buffer = Convert.FromBase64String(cadenabase);
            string rpta = Excel.EnviarCorreo("wilberthpinedacamargo@gmail.com", new List<string> { correo },
                  "Curso Reportes en Net Core", "Te adjunto tu informacion completa registrada",
                  new List<string> { "Informacion.xlsx" }, new List<byte[]> { buffer });
            return rpta;
        }


        public string generarReportePorId(int id)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                //Data
                List<Persona> listaImprimir = new List<Persona>();

                BDReportesContext bDReportesContext = new BDReportesContext();
                listaImprimir = bDReportesContext.Personas.Where(p => p.Iidpersona == id).ToList();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ep = new ExcelPackage();
                int indice = 0;
                foreach (Persona persona in listaImprimir)
                {
                    ep.Workbook.Worksheets.Add(persona.Nombre + " " + persona.Appaterno + " " + persona.Apmaterno);
                    ExcelWorksheet ew1 = ep.Workbook.Worksheets[indice];
                    ew1.Column(2).Width = 50;
                    ew1.Column(3).Width = 50;
                    Excel.tituloHorizontal(ew1, persona.Nombre + " " + persona.Appaterno, 1, 2, 3);
                    Excel.cabecerasFila(ew1, 2, 2, new List<string> { "Identificacion" , "Nombre","Apellido Paterno",
                        "Apellido Materno","Correo","Celular 1","Celular 2","Calle","Numero exterior",
                    "Numero interior","Colonia","Cp","Municipio","Estado","Registro Unico Contribuyente"
                    });
                    Excel.objetoFila<Persona>(ew1, persona, 3, 2, new List<string> { "Numeroidentificacion", "Nombre",
                    "Appaterno","Apmaterno","Correo","Telefonoocelular1","Telefonoocelular2","Calle","Nexterior",
                    "Ninterior","Colonia","Cp","Municipiopais","Estadopais","Numeroregistrounicocontribuyente"
                    });
                    indice++;
                }
                ep.SaveAs(ms);
                byte[] buffer = ms.ToArray();
                return Convert.ToBase64String(buffer);
            }
        }
        public string generarExcelComprimido(string checks)
        {
            //Data
            List<Persona> listaImprimir = new List<Persona>();
            List<string> listanombres = new List<string>();
            List<byte[]> listabytes = new List<byte[]>();

            List<int> ids = checks.Split(",").ToList().Select(int.Parse).ToList();
            BDReportesContext bDReportesContext = new BDReportesContext();
            List<Persona> listaPersona;

            listaPersona = bDReportesContext.Personas.ToList();
            foreach (int id in ids)
            {
                listaImprimir.Add(listaPersona.Where(p => p.Iidpersona == id).First());
            }

            foreach (Persona persona in listaImprimir)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelPackage ep = new ExcelPackage();

                    ep.Workbook.Worksheets.Add(persona.Nombre + " " + persona.Appaterno + " " + persona.Apmaterno);
                    listanombres.Add(persona.Nombre + " " + persona.Appaterno + " " + persona.Apmaterno);
                    ExcelWorksheet ew1 = ep.Workbook.Worksheets[0];
                    ew1.Column(2).Width = 50;
                    ew1.Column(3).Width = 50;
                    Excel.tituloHorizontal(ew1, persona.Nombre + " " + persona.Appaterno, 1, 2, 3);
                    Excel.cabecerasFila(ew1, 2, 2, new List<string> { "Identificacion" , "Nombre","Apellido Paterno",
                        "Apellido Materno","Correo","Celular 1","Celular 2","Calle","Numero exterior",
                    "Numero interior","Colonia","Cp","Municipio","Estado","Registro Unico Contribuyente"
                    });
                    Excel.objetoFila<Persona>(ew1, persona, 3, 2, new List<string> { "Numeroidentificacion", "Nombre",
                    "Appaterno","Apmaterno","Correo","Telefonoocelular1","Telefonoocelular2","Calle","Nexterior",
                    "Ninterior","Colonia","Cp","Municipiopais","Estadopais","Numeroregistrounicocontribuyente"
                    });

                    ep.SaveAs(ms);
                    byte[] buffer = ms.ToArray();
                    listabytes.Add(buffer);
                }
            }
            return Excel.comprimirExcel(listanombres, listabytes);
        }

        public string generarReporteCheckMultipleHoja(string checks)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                //Data
                List<Persona> listaImprimir = new List<Persona>();
                List<int> ids = checks.Split(",").ToList().Select(int.Parse).ToList();
                BDReportesContext bDReportesContext = new BDReportesContext();
                List<Persona> listaPersona;

                listaPersona = bDReportesContext.Personas.ToList();
                foreach (int id in ids)
                {
                    listaImprimir.Add(listaPersona.Where(p => p.Iidpersona == id).First());
                }
                //
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ep = new ExcelPackage();
                int indice = 0;
                foreach (Persona persona in listaImprimir)
                {
                    ep.Workbook.Worksheets.Add(persona.Nombre + " " + persona.Appaterno + " " + persona.Apmaterno);
                    ExcelWorksheet ew1 = ep.Workbook.Worksheets[indice];
                    ew1.Column(2).Width = 50;
                    ew1.Column(3).Width = 50;
                    Excel.tituloHorizontal(ew1, persona.Nombre + " " + persona.Appaterno, 1, 2, 3);
                    Excel.cabecerasFila(ew1, 2, 2, new List<string> { "Identificacion" , "Nombre","Apellido Paterno",
                        "Apellido Materno","Correo","Celular 1","Celular 2","Calle","Numero exterior",
                    "Numero interior","Colonia","Cp","Municipio","Estado","Registro Unico Contribuyente"
                    });
                    Excel.objetoFila<Persona>(ew1, persona, 3, 2, new List<string> { "Numeroidentificacion", "Nombre",
                    "Appaterno","Apmaterno","Correo","Telefonoocelular1","Telefonoocelular2","Calle","Nexterior",
                    "Ninterior","Colonia","Cp","Municipiopais","Estadopais","Numeroregistrounicocontribuyente"
                    });
                    indice++;
                }

                ep.SaveAs(ms);
                byte[] buffer = ms.ToArray();
                return Convert.ToBase64String(buffer);
            }
        }

        //4,5,6,7,9
        public string generarReporteCheck(string checks)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                List<Persona> listaImprimir = new List<Persona>();
                List<int> ids = checks.Split(",").ToList().Select(int.Parse).ToList();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ep = new ExcelPackage();
                ep.Workbook.Worksheets.Add("Hoja de Prueba");

                ExcelWorksheet ew1 = ep.Workbook.Worksheets[0];

                Excel.tituloHorizontal(ew1, "Reportes Persona", 1, 1, 4, 22);
                Excel.anchosColumnas(ew1, 1, new List<int> { 25, 25, 25, 30 });
                Excel.cabecerasTabla(ew1, 2, 1, new List<string> { "Nombres", "Apellido Paterno", "Apellido Materno"
                    , "Telefono" });
                BDReportesContext bDReportesContext = new BDReportesContext();
                List<Persona> listaPersona;

                listaPersona = bDReportesContext.Personas.ToList();
                foreach (int id in ids)
                {
                    listaImprimir.Add(listaPersona.Where(p => p.Iidpersona == id).First());
                }
                Excel.filasTabla<Persona>(ew1, listaImprimir, 3, 1, new List<string> { "Nombre", "Appaterno", "Apmaterno"
                    , "Telefonoocelular1" });
                ep.SaveAs(ms);
                byte[] buffer = ms.ToArray();
                return Convert.ToBase64String(buffer);
            }
        }

        public string generarReporte(string nombre)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ep = new ExcelPackage();
                ep.Workbook.Worksheets.Add("Hoja de Prueba");

                ExcelWorksheet ew1 = ep.Workbook.Worksheets[0];

                Excel.tituloHorizontal(ew1, "Reportes Persona", 1, 1, 4, 22);
                Excel.anchosColumnas(ew1, 1, new List<int> { 25, 25, 25, 30 });
                Excel.cabecerasTabla(ew1, 2, 1, new List<string> { "Nombres", "Apellido Paterno", "Apellido Materno"
                    , "Telefono" });
                BDReportesContext bDReportesContext = new BDReportesContext();
                List<Persona> listaPersona;
                if (nombre == null)
                    listaPersona = bDReportesContext.Personas.ToList();
                else
                    listaPersona = bDReportesContext.Personas.Where(p => p.Nombre.Contains(nombre)).ToList();
                Excel.filasTabla<Persona>(ew1, listaPersona, 3, 1, new List<string> { "Nombre", "Appaterno", "Apmaterno"
                    , "Telefonoocelular1" });
                ep.SaveAs(ms);
                byte[] buffer = ms.ToArray();
                return Convert.ToBase64String(buffer);
            }
        }

        public List<Persona> listaPersonas(string nombre)
        {
            BDReportesContext bDReportesContext = new BDReportesContext();
            if (nombre == null)
                return bDReportesContext.Personas.ToList();
            else
                return bDReportesContext.Personas.Where(p => p.Nombre.Contains(nombre)).ToList();
        }

    }
}
