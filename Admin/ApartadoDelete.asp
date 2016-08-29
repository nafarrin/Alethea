<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idApartado, Mindex
idApartado=Trim(Request.Form("idApartado"))
Mindex=Trim(Request.Form("index"))

%>
<%
Dim rsLineasPlantilla__par1
rsLineasPlantilla__par1 = "-1"
If (idApartado <> "") Then 
  rsLineasPlantilla__par1 = idApartado
End If
%>
<%
Dim rsLineasPlantilla
Dim rsLineasPlantilla_numRows

Set rsLineasPlantilla = Server.CreateObject("ADODB.Recordset")
rsLineasPlantilla.ActiveConnection = cn_STRING
rsLineasPlantilla.Source = "SELECT *  FROM dbo.Plantilla_lineas_Vista  WHERE idApartado = " + Replace(rsLineasPlantilla__par1, "'", "''") + ""
rsLineasPlantilla.CursorType = 0
rsLineasPlantilla.CursorLocation = 2
rsLineasPlantilla.LockType = 1
rsLineasPlantilla.Open()

rsLineasPlantilla_numRows = 0
%>
<%
Dim rsLineasAnalisis__par1
rsLineasAnalisis__par1 = "-1"
If (idApartado <> "") Then 
  rsLineasAnalisis__par1 = idApartado
End If
%>
<%
Dim rsLineasAnalisis
Dim rsLineasAnalisis_numRows

Set rsLineasAnalisis = Server.CreateObject("ADODB.Recordset")
rsLineasAnalisis.ActiveConnection = cn_STRING
rsLineasAnalisis.Source = "SELECT *  FROM dbo.Analisis_lineas_Vista  WHERE idApartado = " + Replace(rsLineasAnalisis__par1, "'", "''") + ""
rsLineasAnalisis.CursorType = 0
rsLineasAnalisis.CursorLocation = 2
rsLineasAnalisis.LockType = 1
rsLineasAnalisis.Open()

rsLineasAnalisis_numRows = 0
%>
<%
Dim rsSubapartados__par1
rsSubapartados__par1 = "-1"
If (idApartado <> "") Then 
  rsSubapartados__par1 = idApartado
End If
%>
<%
Dim rsSubapartados
Dim rsSubapartados_numRows

Set rsSubapartados = Server.CreateObject("ADODB.Recordset")
rsSubapartados.ActiveConnection = cn_STRING
rsSubapartados.Source = "SELECT *  FROM dbo.SubApartados  WHERE idApartado = " + Replace(rsSubapartados__par1, "'", "''") + ""
rsSubapartados.CursorType = 0
rsSubapartados.CursorLocation = 2
rsSubapartados.LockType = 1
rsSubapartados.Open()

rsSubapartados_numRows = 0
%>

<%
if rsLineasAnalisis.eof and rsLineasPlantilla.eof and rsSubApartados.eof then 
	dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	
	'cambiar orden 
	cmdDelete.CommandText = "Update Lineas_Apartados set OrdenConsulta=OrdenConsulta-1 where OrdenConsulta> (select OrdenConsulta from Lineas_Apartados where idApartado=" & idApartado & ") "
	cmdDelete.Execute
	
	cmdDelete.CommandText = "Delete from dbo.Lineas_Apartados where idApartado=" & idApartado
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	response.Redirect("Apartados.asp?index=" & Mindex)
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
<%Cabecera "../img/Maestros.gif", "Maestros: Apartados ", "Borrar apartado" %>
<%Cabecera2 "Cerrar"%>
<table width="100%">
	<tr>
  	<td colspan="6" class="TextoNegrita"><%=TraducirTexto("No puede borrarse el apartado, porque  tienen asignadas las siguientes líneas:")%></td>
	</tr>
	<tr>
		<td colspan="4" class="TextoInverso"><%=TraducirTexto("Plantilla")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Apartado")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Línea")%></td>
	</tr>
	<%while not rsLineasPlantilla.eof  %>
		<tr>
			<td colspan="4" class="Texto"><%=(rsLineasPlantilla.Fields.Item("Plantilla").Value)%></td>
			<td class="Texto"><%=(rsLineasPlantilla.Fields.Item("Apartado").Value)%></td>
			
    <td class="Texto">&nbsp;<%=(rsLineasPlantilla.Fields.Item("Concepto").Value)%></td>
		</tr>
	<%
		rsLineasPlantilla.movenext
	wend
	%>
	<tr><td><br>&nbsp;</td></tr>
	<tr>
		<td class="TextoInverso"><%=TraducirTexto("Operación")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Análisis")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Estado")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Versión")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Apartado")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Línea")%></td>
	</tr>
	
  <%while not rsLineasAnalisis.eof  %>
		<tr>
			<td class="Texto"><%=(rsLineasAnalisis.Fields.Item("Operacion").Value)%></td>
			<td class="Texto"><%=(rsLineasAnalisis.Fields.Item("Analisis").Value)%></td>
			<td class="Texto"><%=(rsLineasAnalisis.Fields.Item("Estado").Value)%></td>
			<td class="Texto"><%=(rsLineasAnalisis.Fields.Item("Version").Value)%></td>
			<td class="Texto"><%=(rsLineasAnalisis.Fields.Item("Apartado").Value)%></td>
			
    <td class="Texto">&nbsp;<%=(rsLineasAnalisis.Fields.Item("Concepto").Value)%></td>
		</tr>
	<%
		rsLineasAnalisis.movenext
	wend
	%>
	<%if not rsSubapartados.eof then%>
	<tr><td class="TextoNegrita"><br><%=TraducirTexto("Tiene asignados los siguientes subapartados")%></td></tr>
	<tr>
		<td class="TextoInverso" colspan="5"><%=TraducirTexto("Subapartado")%></td>
	</tr>
	<%
	end if
	while not rsSubApartados.eof
	%>
		<tr>
		<td class="Texto" colspan="5"><%=(rsSubapartados.Fields.Item("SubApartado").Value)%></td>
	</tr>
<%
		rsSubApartados.movenext
	wend
	%>
</table>
<%
end if
%>
<%
rsLineasPlantilla.Close()
Set rsLineasPlantilla = Nothing
%>
<%
rsLineasAnalisis.Close()
Set rsLineasAnalisis = Nothing
%>
<%
rsSubapartados.Close()
Set rsSubapartados = Nothing
%>
