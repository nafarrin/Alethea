<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idTipoNegocio, Mindex
idTipoNegocio=Trim(Request.Form("idTipoNegocio"))
Mindex=Trim(Request.Form("index"))

%>
<%
Dim rsRelacion__par1
rsRelacion__par1 = "-1"
If (idTipoNegocio  <> "") Then 
  rsRelacion__par1 = idTipoNegocio 
End If
%>
<%
Dim rsRelacion
Dim rsRelacion_numRows

Set rsRelacion = Server.CreateObject("ADODB.Recordset")
rsRelacion.ActiveConnection = cn_STRING
rsRelacion.Source = "SELECT *  FROM dbo.OperacionesInmobiliarias  WHERE idTipoNegocio = " + Replace(rsRelacion__par1, "'", "''") + "  ORDER BY Operacion"
rsRelacion.CursorType = 0
rsRelacion.CursorLocation = 2
rsRelacion.LockType = 1
rsRelacion.Open()

rsRelacion_numRows = 0
%>
<%
if rsRelacion.eof then 
dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	cmdDelete.CommandText = "Delete from dbo.TiposNegocio where idTipoNegocio=" & idTipoNegocio
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	response.Redirect("TiposNegocio.asp?index=" & Mindex)
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
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos de negocio ", "Borrar tipo de negocio" %>
<%Cabecera2 "Cerrar" %>
<br>
<table width="100%">
	<tr>
  	<td colspan="6" class="TextoNegrita"><%= TraducirTexto("No puede borrarse el tipo de negocio, porque está asignado en las siguientes operaciones")%>:</td>
	</tr>
	<tr>
		<td colspan="4" class="TextoInverso"><%= TraducirTexto("Operación")%></td>
	</tr>
	<%while not rsRelacion.eof  %>
		<tr>
			<td colspan="4" class="Texto"><%=(rsRelacion.Fields.Item("Operacion").Value)%></td>
		</tr>
	<%
		rsRelacion.movenext
	wend
	%>
</table>
<%
end if
%>

<%
rsRelacion.Close()
Set rsRelacion = Nothing
%>
