using System;
using System.Collections.Generic;

namespace ReportesMVC.Models
{
    public partial class Persona
    {
        public Persona()
        {
            EnfermedadesIntervencionesPersonas = new HashSet<EnfermedadesIntervencionesPersona>();
            ExpedienteVacunacions = new HashSet<ExpedienteVacunacion>();
            FichaMedicas = new HashSet<FichaMedica>();
            HoraDiaActividadPersonas = new HashSet<HoraDiaActividadPersona>();
            MedicamentoConsumoPersonas = new HashSet<MedicamentoConsumoPersona>();
            MedicoTelefonoPersonas = new HashSet<MedicoTelefonoPersona>();
            ObservacionesPersonas = new HashSet<ObservacionesPersona>();
            TratamientoMedicoPersonas = new HashSet<TratamientoMedicoPersona>();
        }

        public int Iidpersona { get; set; }
        public string? Numeroidentificacion { get; set; }
        public string? Nombre { get; set; }
        public string? Appaterno { get; set; }
        public string? Apmaterno { get; set; }
        public int? Iidsexo { get; set; }
        public string? Correo { get; set; }
        public string? Telefonoocelular1 { get; set; }
        public string? Telefonoocelular2 { get; set; }
        public string? Calle { get; set; }
        public string? Nexterior { get; set; }
        public string? Ninterior { get; set; }
        public string? Colonia { get; set; }
        public string? Cp { get; set; }
        public string? Municipiopais { get; set; }
        public string Estadopais { get; set; } = null!;
        public int? Iidtipodocumento { get; set; }
        public string? Numeroregistrounicocontribuyente { get; set; }
        public string? Nombrefoto { get; set; }
        public int? Bhabilitado { get; set; }

        public virtual Sexo? IidsexoNavigation { get; set; }
        public virtual TipoDocumentoIdentificacion? IidtipodocumentoNavigation { get; set; }
        public virtual ICollection<EnfermedadesIntervencionesPersona> EnfermedadesIntervencionesPersonas { get; set; }
        public virtual ICollection<ExpedienteVacunacion> ExpedienteVacunacions { get; set; }
        public virtual ICollection<FichaMedica> FichaMedicas { get; set; }
        public virtual ICollection<HoraDiaActividadPersona> HoraDiaActividadPersonas { get; set; }
        public virtual ICollection<MedicamentoConsumoPersona> MedicamentoConsumoPersonas { get; set; }
        public virtual ICollection<MedicoTelefonoPersona> MedicoTelefonoPersonas { get; set; }
        public virtual ICollection<ObservacionesPersona> ObservacionesPersonas { get; set; }
        public virtual ICollection<TratamientoMedicoPersona> TratamientoMedicoPersonas { get; set; }
    }
}
