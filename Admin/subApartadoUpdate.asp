<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idSubApartado
dim rsAux, sqlQuery
dim index

index=Trim(Request("index"))
idSubApartado=Trim(Request.Form("idSubApartado"))

sqlQuery="Select * from SubApartados where idSubApartado=" & idSubApartado

AbrirRecordSet rsAux, sqlQuery, cn_STRING

%><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>SPRINT</title>
</head>
<form name="form1" action="SubApartadoDatos.asp" method="post">
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
