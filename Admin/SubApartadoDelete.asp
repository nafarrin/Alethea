<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idSubApartado, Mindex
idSubApartado=Trim(Request.Form("idSubApartado"))
Mindex=Trim(Request.Form("index"))

%>
<%
Dim rsPlantillas__par1
rsPlantillas__par1 = "-1"
If (idSubapartado <> "") Then 
  rsPlantillas__par1 = idSubapartado
End If
%>
<%
Dim rsPlantillas
Dim rsPlantillas_numRows

Set rsPlantillas = Server.CreateObject("ADODB.Recordset")
rsPlantillas.ActiveConnection = cn_STRING
rsPlantillas.Source = "SELECT distinct Plantilla,Apartado, Concepto, OrdenAp, Orden  FROM dbo.plantilla_lineas_vista  WHERE idSubapartado=" + Replace(rsPlantillas__par1, "'", "''") + "  ORDER BY Plantilla,OrdenAp, Orden"
rsPlantillas.CursorType = 0
rsPlantillas.CursorLocation = 2
rsPlantillas.LockType = 1
rsPlantillas.Open()

rsPlantillas_numRows = 0
%>
<%
Dim rsAnalisis__par1
rsAnalisis__par1 = "-1"
If (idSubapartado <> "") Then 
  rsAnalisis__par1 = idSubapartado
End If
%>
<%
Dim rsAnalisis
Dim rsAnalisis_numRows

Set rsAnalisis = Server.CreateObject("ADODB.Recordset")
rsAnalisis.ActiveConnection = cn_STRING
rsAnalisis.Source = "SELECT TOP 100 PERCENT Operacion, Analisis, Estado, Version, Concepto, OrdenAp, Orden, Apartado  FROM dbo.analisis_lineas_vista  WHERE idSubapartado= " + Replace(rsAnalisis__par1, "'", "''") + "  ORDER BY Operacion, Analisis, Estado, Version,OrdenAp, Orden"
rsAnalisis.CursorType = 0
rsAnalisis.CursorLocation = 2
rsAnalisis.LockType = 1
rsAnalisis.Open()

rsAnalisis_numRows = 0
%>
<%
Dim rspartidas__par1
rspartidas__par1 = "-1"
If (idSubapartado <> "") Then 
  rspartidas__par1 = idSubapartado
End If
%>
<%
Dim rspartidas
Dim rspartidas_numRows

Set rspartidas = Server.CreateObject("ADODB.Recordset")
rspartidas.ActiveConnection = cn_STRING
rspartidas.Source = "SELECT Partida, CodigoPartida  FROM dbo.PartidasPresupuestarias  WHERE idSubapartado = " + Replace(rspartidas__par1, "'", "''") + "  ORDER BY CodigoPartida ASC"
rspartidas.CursorType = 0
rspartidas.CursorLocation = 2
rspartidas.LockType = 1
rspartidas.Open()

rspartidas_numRows = 0
%>
<%
if rsAnalisis.eof and rsPlantillas.eof and rsPartidas.eof then 
	dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	
	
	'cambiar orden 
	cmdDelete.CommandText = "Update SubApartados set OrdenConsultaSub=OrdenConsultaSub-1 where OrdenConsultaSub> (select OrdenConsultaSub from SubApartados where idSubApartado=" & idSubApartado & ")  and idApartado =(select idApartado from SubApartados where idSubApartado=" & idSubApartado & ")"
	cmdDelete.Execute

	
	
	cmdDelete.CommandText = "Delete from dbo.SubApartados where idSubApartado=" & idSubApartado
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	response.Redirect("SubApartados.asp?index=" & Mindex)
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
<%Cabecera "../img/Maestros.gif", "Maestros: Subapartados ", "Borrar subapartado" %>
<%Cabecera2 "Cerrar"%>
<br>
<table width="100%">
	<tr>
  	<td colspan="6" class="TextoNegrita"><%= TraducirTexto("No puede borrarse el subapartado, porque se utiliza para las siguientes líneas:") %></td>
	</tr>
	<tr>
		<td colspan="4" class="TextoInverso"><%= TraducirTexto("Plantilla") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Plantilla") %>Apartado</td>
		<td class="TextoInverso"><%= TraducirTexto("Plantilla") %>Línea</td>
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
		<td class="TextoInverso"><%= TraducirTexto("Operación") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Análisis") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Estado") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Versión") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Apartado") %></td>
		<td class="TextoInverso"><%= TraducirTexto("Linea") %></td>
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
	<tr>
  		<td colspan="6" class="TextoNegrita"><br><br><%= TraducirTexto("Tiene asignadas las siguientes partidas:") %></td>
	</tr>
	<tr>
  		<td colspan="6" class="TextoInverso"><%= TraducirTexto("Partidas:") %></td>
	</tr>
	<%while not rspartidas.eof%>
	<tr>
		<td class="Texto" colspan="6"><%=(rspartidas.Fields.Item("CodigoPartida").Value)%>&nbsp;<%=(rspartidas.Fields.Item("Partida").Value)%></td>
	</tr>
	<%
		rspartidas.movenext
	wend
	%>
</table>
</body>
</html>
<%
end if
%>
<%
rsPlantillas.Close()
Set rsPlantillas = Nothing
%>
<%
rsAnalisis.Close()
Set rsAnalisis = Nothing
%>
<%
rspartidas.Close()
Set rspartidas = Nothing
%>
