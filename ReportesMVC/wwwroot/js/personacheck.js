window.onload = function () {
	listarPersonas()
}

function listarPersonas() {
	pintar({
		divPintado: "divPersona",
		url: "Persona/listaPersonas",
		cabeceras: ["Nombre", "Apellido paterno", "Apellido materno"],
		propiedades: ["nombre", "appaterno", "apmaterno"],
		propiedadId: "iidpersona",
		check: true
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
	//alert(idsChecks) [3,4,5,6,7] -> "3,4,5,6,7"
	if (idsChecks.length == 0) {
		Error("Seleccione al menos check")
	}
	var checks = idsChecks.join(",")

	fetchGet("Persona/generarReporteCheck/?checks=" + checks, "text", function (data) {
		var a = document.createElement("a");
		a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + data;
		a.click();
	})
}