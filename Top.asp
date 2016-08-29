<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% option explicit %>
<!--#include file="Connections/cnWeb.asp" -->
<%
TraducirTexto("Resultado Antes Impuestos")
%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Desconectar(){
	document.frmNavegar.action="Desconectar.asp";
	document.frmNavegar.target="_parent";
	document.frmNavegar.submit();
}
function CambioIdioma(){
	document.frmNavegar.action="Usuarios/CambioIdioma.asp";
	document.frmNavegar.target="mainFrame";
	document.frmNavegar.submit();
}
function Conexiones(){
	document.frmNavegar.action="Conexion/gestionConexion.asp";
	document.frmNavegar.target="mainFrame";
	document.frmNavegar.submit();
}
</script>
</head>
<body topmargin="0" background="img/FondoTop2.gif" rightmargin="0">
<form name="frmNavegar" action="" ></form><table width="100%" border="0" cellpadding="0" cellspacing="0" height="20px">
<tr valign="baseline">
	<td class="TextoBlanco">&nbsp;&nbsp;<%= TraducirTexto("Usuario")%>:&nbsp;<%= request.Cookies("NombreUsuario" )%></td>
	<td width="10px" valign="middle" align="right" nowrap="nowrap"><%if BuscarValor("C_SPRINT_CONEXIONES", "count(*)", "idUsuario = " & request.Cookies("idUsuario"))>1 then%><img src="img/Alarma.gif" alt="<%=traducirTExto("Su usuario tiene más de una conexión abierta")%>" style="cursor:hand" onClick="Conexiones()"><%end if%></td>
	<td width="10px"></td>
	<td width="45px"  class="TextoBlanco" valign="middle" nowrap="nowrap"><img src="img/CambioIdioma.gif" alt="<%= TraducirTexto("Cambio idioma")%>" style="cursor:hand" onClick="CambioIdioma()">&nbsp;<img src="img/Desconectar.gif" onClick="Desconectar()" style="cursor:hand" alt="<%= TraducirTexto("Desconectar")%>"></td>
</tr>
</table>
</body>
</html>
