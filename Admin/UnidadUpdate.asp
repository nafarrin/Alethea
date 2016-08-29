<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idUnidad
dim rsAux, sqlQuery, i

idUnidad=Trim(Request.Form("idUnidad"))

sqlQuery="Select * from Unidades where idUnidad=" & idUnidad

AbrirRecordSet rsAux, sqlQuery, cn_STRING

%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Documento sin t&iacute;tulo</title>
<script language="javascript">
function Enviar(){
	document.frmNavegar.submit();
}
</script>
</head>

<body onLoad="Enviar()">
<form name="frmNavegar" action="UnidadNuevo.asp" method="post">
<%
for i=0 to rsAux.fields.count-1
	%>
  <input type="hidden" name="<%= rsAux.fields(i).name %>" value="<%= rsAux.fields(i).value %>">
  <%
next
%>
</form>
</body>
</html>
<%
CerrarRecordSet rsAux
%>
