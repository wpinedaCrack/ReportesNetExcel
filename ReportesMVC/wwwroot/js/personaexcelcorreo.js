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
		rowClick: true,
		cursor: true

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

function rowClickEvent(obj) {
	Confirmacion("Confirmacion", "Desea enviar el correo a " + obj.nombre + " " + obj.appaterno + " " + obj.apmaterno, function () {
		fetchGet("Persona/enviarCorreo/?id=" + obj.iidpersona + "&correo=" + obj.correo, "text", function (data) {
			if (data == "Se envio el Correo satisfactoriamente") {
				Exito(data)
			} else {
				Error(data)
			}
		})
	})
	//fetchGet("Persona/enviarCorreo/?id=" + obj.iidpersona + "&correo=" + obj.correo, "text", function (data) {
	//	if (data == "Se envio el Correo satisfactoriamente") {
	//		Exito(data)
	//	} else {
	//		Error(data)
	//	}
	//})

	//fetchGet("Persona/generarReportePorId/?id=" + obj.iidpersona, "text", function (data) {
	//	var a = document.createElement("a");
	//	a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + data;
	//	a.download = obj.nombre + " " + obj.appaterno + " " + obj.apmaterno // Nombre del Archivo
	//	a.click();
	//})

	//alert(obj.iidpersona)
}


function ExportarExcel() {
	//alert(idsChecks) [3,4,5,6,7] -> "3,4,5,6,7"
	if (idsChecks.length == 0) {
		Error("Seleccione al menos check")
	}
	var checks = idsChecks.join(",")

	fetchGet("Persona/generarReporteCheckMultipleHoja/?checks=" + checks, "text", function (data) {
		var a = document.createElement("a");
		a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + data;
		a.click();
	})
}