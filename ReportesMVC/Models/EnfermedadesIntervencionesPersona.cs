using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class EnfermedadesIntervencionesPersona
    {
        public int Iidenfermedadesintervenciones { get; set; }
        public int? Iidpersona { get; set; }
        public string? Descripcion { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual Persona? IidpersonaNavigation { get; set; }
    }
}
