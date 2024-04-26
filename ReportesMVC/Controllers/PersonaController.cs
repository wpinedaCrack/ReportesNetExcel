using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
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
