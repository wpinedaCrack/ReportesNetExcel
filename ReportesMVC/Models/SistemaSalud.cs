using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class SistemaSalud
    {
        public SistemaSalud()
        {
            FichaMedicas = new HashSet<FichaMedica>();
        }

        public int Iidsistemasalud { get; set; }
        public string? Nombre { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual ICollection<FichaMedica> FichaMedicas { get; set; }
    }
}
