<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim Logo,  sqlQuery
Logo=Trim(Request("Logo"))

if Logo<>"" then
	sqlQuery= sqlQuery &" and Logos.Logo like '%" & replace(Logo,"'","''") & "%' "
end if
if sqlQuery<>"" then
	sqlQuery= " where Logos.idLogo>-1 " &  sqlQuery
end if
%>
<%
Dim rsLogos
AbrirRecordSet rsLogos, "SELECT * FROM Logos "&sqlQuery&" order by Logo ", cn_STRING
%>
<%
Dim rsCuantos
AbrirRecordSet rsCuantos, "SELECT DISTINCT idLogo FROM Sociedades", cn_STRING
%>

<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Nueva(){
	document.frmNavegar.action="LogoDatos.asp";
	document.frmNavegar.submit();
}
function Filtrar(){
	document.frmNavegar.action="Logos.asp";
	document.frmNavegar.submit();
}
function DesFiltrar(){
	document.frmNavegar.Logo.value="";
	document.frmNavegar.action="Logos.asp";
	document.frmNavegar.submit();
}
function Borrar(Cual){
	if (confirm("<%= TraducirTexto("¿Desea eliminar el logo?")%>")){
		document.frmNavegar.action="LogoDelete.asp";
		document.frmNavegar.idLogo.value=Cual;
		document.frmNavegar.submit();
	}
}
function Modificar(Cual){
	document.frmNavegar.action="LogoUpdate.asp";
	document.frmNavegar.idLogo.value=Cual;
	document.frmNavegar.submit();
}
function Cerrar(){
	document.frmNavegar.action="../main.asp";
	document.frmNavegar.submit();
}
</script>
</head>
<body><form name="frmNavegar" action="" method="post">
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Logos", "Listado de logos" %>
<%Cabecera2 "Nuevo;Aplicar Filtros;Quitar Filtros;Cerrar"%>

<table width="100%" class="Casilla">
	<tr>
		
      <td class="TextoNegrita" colspan="4"><%= TraducirTexto("Filtros")%>
        <hr size="1" noshade class="Texto"></td>
	</tr>
	<tr>
		
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Logo")%>:</td>
		
      <td class="Texto" align="left" width="35%">
<input type="text" name="Logo" value="<%=Logo%>" style="width:100%" maxlength="100"></td>
		
      <td class="TextoNegrita" align="right" width="15%">&nbsp;</td>
		
      <td class="Texto" align="left" width="35%">&nbsp; </td>
	</tr>
</table>
<br>

  <table border="0" width="100%" cellspacing="0" cellpadding="2">
    <tr > 
      <td class="Cabecera"><%= TraducirTexto("Logo")%></td>
      <td class="Cabecera"><%= TraducirTexto("Nombre imagen")%></td>
      <td class="Cabecera" colspan="2" width="5%">&nbsp;</td>
    </tr>
    <% While (NOT rsLogos.EOF) %>
    <tr> 
      <td class="Linea"><%=(rsLogos.Fields.Item("Logo").Value)%></td>
	  <td class="Linea"><%=(rsLogos.Fields.Item("NombreImagen").Value)%></td>
	  <td class="Linea">
	  <img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsLogos.Fields.Item("idLogo").Value)%>')" src="../img/Editar.gif">
	  </td>
      <td class="Linea">
	  <%
	  rsCuantos.filter="idLogo=" & (rsLogos.Fields.Item("idLogo").Value)
	   If (rsLogos.Fields.Item("Borrar").Value) and rsCuantos.eof Then %>
	  <img alt="<%= TraducirTexto("Eliminar")%>" style="cursor:hand" onClick="Borrar('<%=(rsLogos.Fields.Item("idLogo").Value)%>')" src="../img/Borrar.gif">
   <% Else %>&nbsp;
<% End If %>
</td>
 	</tr>
	<% rsLogos.movenext 
	Wend %>
  </table>
	<input type="hidden" name="idLogo" value="">
</form>
</body>
</html>
<%
rsLogos.Close()
Set rsLogos = Nothing
%>
<%
rsCuantos.Close()
Set rsCuantos = Nothing
%>
