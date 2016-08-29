<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idHito, Mindex
idHito=Trim(Request.Form("idHito"))
Mindex=Trim(Request.Form("index"))


dim sqlQuery
dim rsAnalisis, rsPlantillas, rsFuncionesHitos, rsTareas

sqlQuery="SELECT Operacion, Analisis, Estado, Version, Concepto, OrdenAp, Orden, Apartado  FROM Analisis_lineas_vista  WHERE idHito = " & idHito & " or idHitoFIn=" & idHito & "   ORDER BY Operacion, Analisis, Estado, Version,OrdenAp, Orden"
AbrirRecordset rsAnalisis, sqlQuery, cn_STRING


sqlQuery="SELECT distinct Plantilla,Apartado, Concepto, OrdenAp, Orden  FROM dbo.Plantilla_lineas_Vista  WHERE idHito=" & idHito & " or idHitoFin=" & idHito & "  ORDER BY Plantilla,OrdenAp, Orden"
AbrirRecordset rsPlantillas, sqlQuery, cn_STRING

sqlQuery="SELECT FuncionHito  FROM FuncionHitos   INNER JOIN FuncionHitos_Valores   ON FuncionHitos.idFuncionHito = FuncionHitos_Valores.idFuncionHito  where idHito=" & idHito &""
AbrirRecordset rsFuncionesHitos, sqlQuery, cn_STRING


sqlQuery="SELECT distinct PlantillaTareas FROM Tareas_Hitos INNER JOIN Tareas_Plantillas ON Tareas_Hitos.idPlantillaTareas=Tareas_Plantillas.idPlantillaTareas WHERE Tareas_Hitos.idHito=14"
AbrirRecordset rsTareas, sqlQuery, cn_STRING



dim cmdDelete
Set cmdDelete = Server.CreateObject("ADODB.Command")
cmdDelete.ActiveConnection = cn_STRING

'hay que borrr previamente las fucniones que ya no pertenecen a ninsgún análisis
cmdDelete.CommandText = "Delete from FuncionHitos where General=0 and idFuncionHito not in (Select idFuncionHito from Analisis_FuncionHitos)"
cmdDelete.Execute




if rsAnalisis.eof and rsPlantillas.eof and rsFuncionesHitos.eof and rsTareas.eof then 
	cmdDelete.CommandText = "Delete from dbo.Hitos where idHito=" & idHito
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	cerrarRecordSet rsAnalisis
	cerrarRecordSet rsPlantillas
	cerrarRecordSet rsFuncionesHitos
	cerrarRecordSet rsTareas

	response.Redirect("Hitos.asp?index=" & Mindex)
else

	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing

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
<%Cabecera "../img/Maestros.gif", "Maestros: Hitos ", "Modificar hito" %>
<%Cabecera2 "Cerrar" %>
<br />
<table width="100%">
	<tr>
  	<td colspan="6" class="TextoNegrita"><%= TraducirTexto("No puede borrarse el hito, porque se utiliza en:")%></td>
	</tr>
	<tr>
		<td colspan="4" class="TextoInverso"><%= TraducirTexto("Plantilla")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Apartado")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Línea")%></td>
	</tr>
	<%while not rsPlantillas.eof  %>
		<tr>
			<td colspan="4" class="Texto"><%=(rsPlantillas.Fields.Item("Plantilla").Value)%></td>
			<td class="Texto"><%=(rsPlantillas.Fields.Item("Apartado").Value)%></td>
			
    <td class="Texto">&nbsp;<%=(rsPlantillas.Fields.Item("Concepto").Value)%></td>
		</tr>
	<%
		rsPlantillas.movenext
	wend
	%>
	<tr><td><br>&nbsp;</td></tr>
	<tr>
		<td class="TextoInverso"><%= TraducirTexto("Operaci&oacute;n")%></td>
		<td class="TextoInverso"><%= TraducirTexto("An&aacute;lisis")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Estado")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Versión")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Apartado")%></td>
		<td class="TextoInverso"><%= TraducirTexto("Línea")%></td>
	</tr>
	
  <%while not rsAnalisis.eof  %>
		<tr>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Operacion").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Analisis").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Estado").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Version").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Apartado").Value)%></td>
    		<td class="Texto">&nbsp;<%=(rsAnalisis.Fields.Item("Concepto").Value)%></td>
		</tr>
	<%
		rsAnalisis.movenext
	wend
	%>
	<tr><td><br>&nbsp;</td></tr>
	<tr>
		<td class="TextoInverso" colspan="6"><%= TraducirTexto("Función desglose hito")%></td>
	</tr>
	
  <%while not rsFuncionesHitos.eof  %>
		<tr>
			<td class="Texto" colspan="6"><%=(rsFuncionesHitos.Fields.Item("FuncionHito").Value)%></td>
		</tr>
	<%
		rsFuncionesHitos.movenext
	wend
	%>
	<tr><td><br>&nbsp;</td></tr>
	<tr>
		<td class="TextoInverso" colspan="6"><%= TraducirTexto("Plantilla de Tareas")%></td>
	</tr>
	
  <%while not rsTareas.eof  %>
		<tr>
			<td class="Texto" colspan="6"><%=(rsTareas.Fields.Item("PlantillaTareas").Value)%></td>
		</tr>
	<%
		rsTareas.movenext
	wend
	%>
	
	
	
	</table>


<%
end if
%>
<%
	cerrarRecordSet rsAnalisis
	cerrarRecordSet rsPlantillas
	cerrarRecordSet rsFuncionesHitos
	cerrarRecordSet rsTareas
%>
