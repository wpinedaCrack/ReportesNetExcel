using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class Hora
    {
        public Hora()
        {
            HoraDiaActividadPersonas = new HashSet<HoraDiaActividadPersona>();
        }

        public int Iidhora { get; set; }
        public string? Hora1 { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual ICollection<HoraDiaActividadPersona> HoraDiaActividadPersonas { get; set; }
    }
}
