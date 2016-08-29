<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idUnidad, Mindex
idUnidad=Trim(Request.Form("idUnidad"))
Mindex=Trim(Request.Form("index"))

%>
<%
Dim rsAnalisis__par1
rsAnalisis__par1 = "-1"
If (idUnidad <> "") Then 
  rsAnalisis__par1 = idUnidad
End If
%>
<%
Dim rsAnalisis
Dim rsAnalisis_numRows

Set rsAnalisis = Server.CreateObject("ADODB.Recordset")
rsAnalisis.ActiveConnection = cn_STRING
rsAnalisis.Source = "SELECT *  FROM dbo.Analisis_Apartados_vista  WHERE idUnidad1=" + Replace(rsAnalisis__par1, "'", "''") + "  or idUnidad2=" + Replace(rsAnalisis__par1, "'", "''") + "  or idUnidad3=" + Replace(rsAnalisis__par1, "'", "''") + "  ORDER BY analisis, orden"
rsAnalisis.CursorType = 0
rsAnalisis.CursorLocation = 2
rsAnalisis.LockType = 1
rsAnalisis.Open()

rsAnalisis_numRows = 0
%>
<%
Dim rsPlantillas__par1
rsPlantillas__par1 = "-1"
If (idUnidad <> "") Then 
  rsPlantillas__par1 = idUnidad
End If
%>
<%
Dim rsPlantillas
Dim rsPlantillas_numRows

Set rsPlantillas = Server.CreateObject("ADODB.Recordset")
rsPlantillas.ActiveConnection = cn_STRING
rsPlantillas.Source = "SELECT *  FROM dbo.Plantilla_Apartados_vista  WHERE idUnidad1=" + Replace(rsPlantillas__par1, "'", "''") + "  or idUnidad2=" + Replace(rsPlantillas__par1, "'", "''") + "  or idUnidad3=" + Replace(rsPlantillas__par1, "'", "''") + "  order by plantilla, orden"
rsPlantillas.CursorType = 0
rsPlantillas.CursorLocation = 2
rsPlantillas.LockType = 1
rsPlantillas.Open()

rsPlantillas_numRows = 0
%>
<%
if rsAnalisis.eof and rsPlantillas.eof then 
	dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	cmdDelete.CommandText = "Delete from dbo.Unidades where idUnidad=" & idUnidad
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	response.Redirect("Unidades.asp?index=" & Mindex)
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
<%Cabecera "../img/Maestros.gif", "Maestros: Unidades", "Borrar unidad" %>
<%Cabecera2 "Cerrar"%>
<br>
<table width="100%">
	<tr>
  	<td colspan="6" class="TextoNegrita"><%=TraducirTexto("No puede borrarse la unidad, porque se utiliza en los siguientes apartados:")%></td>
	</tr>
	<tr>
		<td colspan="4" class="TextoInverso"><%=TraducirTexto("Plantilla")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Apartado")%></td>
	</tr>
	<%while not rsPlantillas.eof  %>
		<tr>
			<td colspan="4" class="Texto"><%=(rsPlantillas.Fields.Item("Plantilla").Value)%></td>
			<td class="Texto"><%=(rsPlantillas.Fields.Item("Apartado").Value)%></td>
		</tr>
	<%
		rsPlantillas.movenext
	wend
	%>
	<tr><td><br>&nbsp;</td></tr>
	<tr>
		<td class="TextoInverso"><%=TraducirTexto("Operación")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Análisis")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Estado")%>")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Versión")%></td>
		<td class="TextoInverso"><%=TraducirTexto("Apartado")%></td>
	</tr>
	
  <%while not rsAnalisis.eof  %>
		<tr>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Operacion").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Analisis").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Estado").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Version").Value)%></td>
			<td class="Texto"><%=(rsAnalisis.Fields.Item("Apartado").Value)%></td>
		</tr>
	<%
		rsAnalisis.movenext
	wend
	%>
</table>
</body>
</html>
<%
end if
%>

<%
rsAnalisis.Close()
Set rsAnalisis = Nothing
%>
<%
rsPlantillas.Close()
Set rsPlantillas = Nothing
%>
