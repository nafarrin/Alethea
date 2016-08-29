<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idApartado
dim rsAux, sqlQuery
dim index

index=Trim(Request("index"))
idApartado=Trim(Request.Form("idApartado"))

sqlQuery="Select * from Lineas_Apartados where idApartado=" & idApartado

AbrirRecordSet rsAux, sqlQuery, cn_STRING

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>SPRINT</title>
</head>

<body>

<form name="form1" action="ApartadoDatos.asp" method="post">
<%
dim i
for i=0 to rsAux.fields.count-1
	%>
  <input type="hidden" name="<%= rsAux.fields(i).name %>" value="<%= rsAux.fields(i).value %>">
  <%
next
%>
<input type="hidden" name="index" value="<%=index%>">
</form>
</body>
<%
CerrarRecordSet rsAux
%>
<script language="JavaScript">
document.form1.submit();
</script>
</html>
