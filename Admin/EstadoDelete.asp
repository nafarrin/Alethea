<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idEstado, Mindex
idEstado=Trim(Request.Form("idEstado"))
Mindex=Trim(Request.Form("index"))

dim sqlQuery, rsAnalisis

SqlQuery="Select top 1 * from AnalisisViabilidad where idEstado=" & idEstado

AbrirRecordSet rsAnalisis, sqlQuery, cn_STRING

if rsAnalisis.eof then 
	dim cmdDelete
	Set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = cn_STRING
	cmdDelete.CommandText = "Delete from Estados where idEstado=" & idEstado
	cmdDelete.Execute
	cmdDelete.ActiveConnection.Close
	
	set cmdDelete=nothing
	
	response.Redirect("Estados.asp?index=" & Mindex)
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
<%Cabecera "../img/Maestros.gif", "Maestros: Estados ", "Borrar estado" %>
<%Cabecera2 "Cerrar" %>

<table width="100%">
	<tr>
  	<td c class="TextoNegrita"><%=TraducirTexto("No puede borrarse el estado porque tiene asignados análisis.")%></td>
</table>
<%

end if
CerrarRecordSet rsAnalisis

%>