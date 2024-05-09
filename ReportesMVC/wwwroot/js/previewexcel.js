function EnviarArchivoExcel() {
	var frmPreview = document.getElementById("frmPreview");
	var frm = new FormData(frmPreview)
	fetchPost("Persona/leerExcel", "text", frm, function (data) {
		pintarExcel(data)
	})
}


function pintarExcel(data) {

	var contenido = "<table class='table mt-5'>";
	var array = data.split("_")
	var estilos = array[1].split("|")
	var datosFila = array[0].split("¬");
	var nfilas = datosFila.length;
	var ncolumnas = datosFila[0].split("|").length
	var filaActual;
	var campos
	contenido += "<tr>";
	contenido += "<td></td>"
	for (var j = 0; j < ncolumnas; j++) {

		contenido += "<td style='background-color:#bdbdbd;border:1px solid #d3d3d3 !important'>";
		contenido += String.fromCharCode(65 + j)
		contenido += "</td>";
	}
	contenido += "</tr>";
	for (var i = 0; i < datosFila.length; i++) {
		filaActual = datosFila[i]
		campos = filaActual.split("|")
		contenido += "<tr>";
		contenido += "<td style='background-color:#bdbdbd;border:1px solid #d3d3d3 !important' >";
		contenido += (i + 1)
		contenido += "</td>";
		for (var j = 0; j < ncolumnas; j++) {
			var estilo = "";
			var row = i + 1;
			var column = j + 1
			estilo += estilos.includes("n" + row + "¬" + column) ? "font-weight:bold;" : ""
			estilo += estilos.includes("bt" + row + "¬" + column) ? "border-top:1.1px solid;" : ""
			estilo += estilos.includes("bb" + row + "¬" + column) ? "border-bottom:1px solid;" : ""
			estilo += estilos.includes("br" + row + "¬" + column) ? "border-right:1px solid;" : ""
			estilo += estilos.includes("bl" + row + "¬" + column) ? "border-left:1px solid;" : ""
			estilo += estilos.includes("hc" + row + "¬" + column) ? "text-align:center;" : ""
			estilo += estilos.includes("hr" + row + "¬" + column) ? "text-align:right;" : ""

			contenido += `<td style='${estilo}' >`;
			contenido += campos[j]
			contenido += "</td>";
		}
		contenido += "</tr>";
	}


	contenido += "</table>"
	setI("divTabla", contenido)
}