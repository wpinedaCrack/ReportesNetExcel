using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class Factura
    {
        public Factura()
        {
            DetalleFacturas = new HashSet<DetalleFactura>();
        }

        public int Iidfactura { get; set; }
        public int? Iidcompania { get; set; }
        public int? Numerofactura { get; set; }
        public decimal? Subtotalfactura { get; set; }
        public decimal? Impuesto { get; set; }
        public decimal? Otro { get; set; }
        public decimal? Total { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual Compañium? IidcompaniaNavigation { get; set; }
        public virtual ICollection<DetalleFactura> DetalleFacturas { get; set; }
    }
}
