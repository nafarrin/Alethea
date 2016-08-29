<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery="SELECT t.*, i.TipoVersionIdioma AS Traduccion FROM Aux_tipoVersion t "_
& " LEFT JOIN IdiomasTipoVersion i ON t.idTipoVersion = i.idTipoVersion AND i.idIdioma=" &  Request.Cookies("idIdiomaCookie") _
& " order by t.TipoVersion "

'sqlQuery="SELECT *  FROM Aux_tipoVersion order by TipoVersion "

AbrirRecordSet rsAux, sqlQuery, cn_STRING
%>
<%
Dim rsIdioma
sqlQuery="SELECT Idioma From Idiomas Where idIdioma="& Request.Cookies("idIdiomaCookie")
AbrirRecordSet rsIdioma, sqlQuery, cn_STRING
%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Modificar(Cual){
	document.frmNavegar.action="TiposVersionesUpdate.asp";
	document.frmNavegar.idTipoVersion.value=Cual;
	document.frmNavegar.submit();
}

function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="TipoVersionIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos Versión", "Listado de Tipos de Versiones" %>
<%Cabecera2 "Cerrar" %>

  <table border="0" width="100%" cellspacing="0" cellpadding="2">
  	<tr>
		<td colspan="2"></td>
		<td class="Cabecera" colspan="2" nowrap="nowrap" align="center"><%= TraducirTexto("Exportar a OLAP")%></td>
	</tr>
    <tr > 
      <td class="Cabecera" width="35%"><%= TraducirTexto("Tipo Versión")%></td>
	  <td class="Cabecera" width="65%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
	  <td class="Cabecera" nowrap="nowrap" align="center"><%= TraducirTexto("Presupuesto")%></td>
	  <td class="Cabecera" nowrap="nowrap" align="center"><%= TraducirTexto("Previsión")%></td>
      <td class="Cabecera" colspan="3">&nbsp;</td>
    </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea" ><%=(rsAux.Fields.Item("TipoVersion").Value)%></td>
	  <td class="Linea" ><%=(rsAux.Fields.Item("Traduccion").Value)%>&nbsp;</td>
	  <td class="Linea" align="center" ><img src="../img/Estado<%=(rsAux.Fields.Item("ExportarPresupuestoOLAP").Value)%>.gif"></td>
	  <td class="Linea" align="center" ><img src="../img/Estado<%=(rsAux.Fields.Item("ExportarPrevisionOLAP").Value)%>.gif"></td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idTipoVersion").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idTipoVersion").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmNavegar" action="" method="post">
	<input type="hidden" name="idTipoVersion" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsIdioma
CerrarRecordSet rsAux
%>
