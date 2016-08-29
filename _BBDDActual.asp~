<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<%

Dim cnVersion
'cnVersion = "Provider=sqloledb;Persist Security Info=False;Connect Timeout=60; Data Source=localhost;Initial Catalog=SPRINT;User Id=SPRINTUser;Password=SPRINTUser;"
cnVersion = "Provider=sqloledb;Persist Security Info=False;Connect Timeout=60; Data Source=DELPHIN-PC\SQL2K;Initial Catalog=SPRINT;User Id=SPRINTUser;Password=SPRINTUser;"

dim rsVersion
dim sqlQuery

sqlQuery="SELECT BBDDOLAP, Version, Servidor FROM Instalaciones where BBDD='" & request.Cookies("BBDDConexion") & "' "
AbrirRecordSet rsVersion, sqlQuery , cnVersion
response.Cookies("Servidor")=rsVersion.fields("Servidor").value
response.Cookies("BBDD_OLAP")=rsVersion.fields("BBDDOLAP").value
%>
<!--#include file="includesSQL/Funciones.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/Web.css" rel="stylesheet" type="text/css">
</head>

<body class="MenuSup" topmargin="0" bottommargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="CabeceraPrevision">Conectado a: <%=(rsVersion.Fields.Item("Version").Value)%></td>
		<td align="right" class="CabeceraPrevision"><a href="_default.asp" target="_parent"><img border="0" src="img/Desconectar.gif"></a></td>
	</tr>
</table>
</body>
</html>
<%
CerrarRecordSet rsVersion
%>
