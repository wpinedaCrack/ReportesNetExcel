window.onload = function () {
	listarPersonas()
}


function listarPersonas() {
	pintar({
		divPintado: "divPersona",
		url: "Persona/listaPersonas",
		cabeceras: ["Nombre", "Apellido paterno", "Apellido materno"],
		propiedades: ["nombre", "appaterno", "apmaterno"],
		propiedadId: "iidpersona"
	}, {
		legend: "",
		idformulario: "frmPersonaBusqueda",
		url: "Persona/listaPersonas",
		formulario: [
			[
				{
					class: "col-md-6",
					label: "Nombre Persona",
					type: "text",
					name: "nombre"

				},

			]
		]
	})
}

function ExportarExcel() {
	var nombre = getN("nombre")
	fetchGet("Persona/generarReporte/?nombre=" + nombre, "text", function (data) {
		var a = document.createElement("a");
		a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + data;
		a.click();
	})
}