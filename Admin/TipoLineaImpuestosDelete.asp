<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->

<%

dim idTipoImpuesto, Mindex
idTipoImpuesto=Trim(Request.Form("idTipoImpuesto"))
Mindex=Trim(Request.Form("index"))

dim sqlQuery, rsPlantilla, rsAnalisis

sqlQuery="SELECT Plantilla,Concepto " _
	& " FROM Plantilla_lineas " _
	& " INNER JOIN Plantillas " _
	& "    ON Plantilla_lineas.idPlantilla=Plantillas.idPlantilla " _
	& " WHERE Plantilla_lineas.idTipoImpuesto=" & idTipoImpuesto _
	& " ORDER BY Plantilla, Concepto "
AbrirRecordSet rsPlantilla, sqlQuery, cn_STRING

sqlQuery="SELECT Proyecto, Operacion, Analisis, Concepto " _
	& " FROM Analisis_lineas " _
	& " INNER JOIN Analisis_Vista " _
	& "    ON Analisis_lineas.idAnalisis=Analisis_Vista.idAnalisis " _
	& " WHERE Analisis_lineas.idTipoImpuesto=" & idTipoImpuesto _
	& " ORDER BY Proyecto, Operacion, Analisis, Concepto "
AbrirRecordSet rsAnalisis, sqlQuery, cn_STRING

if rsAnalisis.eof  and rsPlantilla.eof then 
	dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	cmdDelete.CommandText = "Delete from dbo.AUX_TipoLineaImpuestos where idTipoImpuesto=" & idTipoImpuesto
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	CerrarRecordSet rsPlantilla
	CerrarRecordSet rsAnalisis
	
	response.Redirect("TipoLineaImpuestos.asp?index=" & Mindex)
else
%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Cerrar(){
	window.history.go(-1);
}
</script>
</head>

<body class="Texto" >
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Línea Impuestos ", "Borrar Tipo Línea Impuestos" %>
<%Cabecera2 "Cerrar"%>
<br>
<table>
	<tr>
  	<td colspan="4" class="TextoNegrita"><%= TraducirTexto("No puede borrarse el tipo de línea de impuestos porque está asignado a:")%></td>
	<%
	if not rsPlantilla.eof then
	%>
		<tr><td></td></tr>
		<tr>
			<td colspan="3" class="Cabecera"><%= TraducirTexto("Plantilla")%></td>
			<td  class="Cabecera"><%= TraducirTexto("Línea")%></td>
		</tr>
	<%
		while not rsPlantilla.eof 
		%>
		<tr>
			<td colspan="3" class="Texto"><%=(rsPlantilla.Fields.Item("Plantilla").Value)%></td>
			<td  class="Texto"><%=(rsPlantilla.Fields.Item("Concepto").Value)%></td>
		</tr>
		<%
			rsPlantilla.movenext
		wend
	end if

	if not rsAnalisis.eof then
	%>
		<tr><td></td></tr>
		<tr>
			<td class="Cabecera"><%= TraducirTexto("Proyecto")%></td>
			<td class="Cabecera"><%= TraducirTexto("Operación")%></td>
			<td class="Cabecera"><%= TraducirTexto("Análisis")%></td>
			<td class="Cabecera"><%= TraducirTexto("Línea")%></td>
		</tr>
	<%
		while not rsAnalisis.eof 
		%>
		<tr>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Proyecto").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Operacion").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Analisis").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Concepto").Value)%></td>
		</tr>
		<%
			rsAnalisis.movenext
		wend
	end if

	%>
</table>
<%
end if

CerrarRecordSet rsPlantilla
CerrarRecordSet rsAnalisis

%>