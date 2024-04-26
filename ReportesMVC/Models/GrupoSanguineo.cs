using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class GrupoSanguineo
    {
        public GrupoSanguineo()
        {
            FichaMedicas = new HashSet<FichaMedica>();
        }

        public int Iidgruposanguineo { get; set; }
        public string? Nombresanguineo { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual ICollection<FichaMedica> FichaMedicas { get; set; }
    }
}
