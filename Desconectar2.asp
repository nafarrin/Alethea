<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% option explicit %>
<!--#include file="Connections/cnWebINI.asp" -->
<!--#include file="includes/ConvertirFecha.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>SPRINT</title>
</head>
<body>
<img src="imgLogos/logo.jpg"><br>Cerrando<img src="img/clock_off.gif">
<%
if request.Cookies("idUsuario")<>"" then
	dim sqlQuery
	dim cmdAux
	Set cmdAux = Server.CreateObject("ADODB.Command")
	cmdAux.ActiveConnection = cn_STRING
	
	sqlQuery="Delete from C_SPRINT_CONEXIONES where idUsuario=" & request.Cookies("idUsuario") & " and Inicio='" & request.Cookies("CadenaActual") & "' "
	cmdAux.CommandText = sqlQuery
	cmdAux.Execute
	
	if request.Cookies("ErrorCadena")="" then
		cmdAux.CommandText = "Insert into C_SPRINT_LOG (Usuario, Accion,FechaAccion) Values ('" & request.Cookies("Usuario") & "' ,'Fin','" & ConvertirFecha( date() & " " & time )&"') "
		cmdAux.Execute
	end if
	set cmdAux=nothing
	
	Session.Abandon
	
	response.Cookies("Usuario")=""
	response.Cookies("NombreUsuario")= ""
	response.Cookies("idUsuario")=""
	response.Cookies("idDepartamentoCookie")=""
	response.Cookies("SeguridadProyectos")=""
	response.Cookies("CadenaActual")=""
	response.Cookies("ErrorCadena")=""
end if
'response.Redirect("default.asp")
%>
<script language="javascript">
window.close();
</script>
</body>
</html>

