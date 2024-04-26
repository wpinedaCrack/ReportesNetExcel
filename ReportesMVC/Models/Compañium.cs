using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class Compañium
    {
        public Compañium()
        {
            Facturas = new HashSet<Factura>();
        }

        public int Iidcompania { get; set; }
        public string? Personaafacturar { get; set; }
        public string? Nombrecompañia { get; set; }
        public string? Direccion { get; set; }
        public string? Ciudad { get; set; }
        public string? Estado { get; set; }
        public string? Codigopostal { get; set; }
        public string? Telefono { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual ICollection<Factura> Facturas { get; set; }
    }
}
