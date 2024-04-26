using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class MarcaVacuna
    {
        public MarcaVacuna()
        {
            ExpedienteVacunacions = new HashSet<ExpedienteVacunacion>();
        }

        public int Iidmarcavacuna { get; set; }
        public string? Nombremarca { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual ICollection<ExpedienteVacunacion> ExpedienteVacunacions { get; set; }
    }
}
