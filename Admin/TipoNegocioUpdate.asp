<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idTipoNegocio
idTipoNegocio=Trim(Request("idTipoNegocio"))

dim sqlQuery
Dim rsAux

sqlQuery="SELECT *  FROM TiposNegocio  WHERE idTipoNegocio =" & idTipoNegocio
AbrirRecordSet rsAux, sqlQuery, cn_STRING

%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<form name="form1" action="TipoNegocioDatos.asp" method="post">
<%
dim i
'campos de la línea
for i=0 to rsAux.fields.count-1
%>
  <input type="hidden" name="<%= rsAux.fields(i).name %>" value="<%= rsAux.fields(i).value %>">
  <%
next
%>
<input type="hidden" name="index" value="<%=Trim(Request("index"))%>">
</form>
</body>
<%
CerrarRecordSet rsAux
%>
<script language="JavaScript">
document.form1.submit();
</script>
</html>