<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoImpuesto
idTipoImpuesto=Trim(Request("idTipoImpuesto"))

dim sqlQuery
Dim rsAux

sqlQuery="SELECT *  FROM AUX_TipoLineaImpuestos  WHERE idTipoImpuesto =" & idTipoImpuesto & " "
AbrirRecordSet rsAux, sqlQuery, cn_STRING

%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<form name="form1" action="TipoLineaImpuestosDatos.asp" method="post">
<%
dim i
'campos de la línea
for i=0 to rsAux.fields.count-1
%>
  <input type="hidden" name="<%= rsAux.fields(i).name %>" value="<%= rsAux.fields(i).value %>">
  <%
next
%>
</form>
</body>
<%
rsAux.Close()
Set rsAux = Nothing
%>
<script language="JavaScript">
document.form1.submit();
</script>
</html>
