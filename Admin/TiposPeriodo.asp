<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery="SELECT Aux_TipoPeriodo.idTipoPeriodo, Aux_TipoPeriodo.TipoPeriodo, IdiomasTipoPeriodo.TipoPeriodoIdioma " _
	& " FROM Aux_TipoPeriodo "_
		& " LEFT JOIN IdiomasTipoPeriodo " _
			& " ON Aux_TipoPeriodo.idTipoPeriodo = IdiomasTipoPeriodo.idTipoPeriodo " _
			& " AND IdiomasTipoPeriodo.idIdioma=" &  Request.Cookies("idIdiomaCookie") _
	& " order by TipoPeriodo "

'sqlQuery="SELECT *  FROM Aux_TipoPeriodo order by TipoPeriodo "

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
	document.frmNavegar.action="TiposPeriodoUpdate.asp";
	document.frmNavegar.idTipoPeriodo.value=Cual;
	document.frmNavegar.submit();
}

function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="TipoPeriodoIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos Periodos", "Listado de Tipos de Periodos" %>
<%Cabecera2 "Cerrar" %>
<br>
  <table border="0" width="100%" cellspacing="0" cellpadding="2">
    <tr > 
      <td class="Cabecera" width="35%"><%= TraducirTexto("Tipo Periodo")%></td>
	  <td class="Cabecera" width="60%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
      <td class="Cabecera" colspan="3">&nbsp;</td>
    </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea" ><%=(rsAux.Fields.Item("TipoPeriodo").Value)%></td>
	  <td class="Linea" ><%=(rsAux.Fields.Item("TipoPeriodoIdioma").Value)%>&nbsp;</td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idTipoPeriodo").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idTipoPeriodo").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmNavegar" action="" method="post">
	<input type="hidden" name="idTipoPeriodo" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsAux
CerrarRecordSet rsIdioma
%>
