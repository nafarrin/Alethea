<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idLogo
idLogo=Trim(Request("idLogo"))

dim index
index=Trim(Request("index"))

dim sqlQuery
Dim rsAux

sqlQuery="SELECT *  FROM Logos  WHERE idLogo =" & idLogo & " "
AbrirRecordSet rsAux, sqlQuery, cn_STRING

%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<form name="form1" action="LogoDatos.asp" method="post">
<%
dim i
'campos de la línea
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


