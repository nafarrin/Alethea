<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery="SELECT Aux_ModalidadCashFlow.idModalidadCashFlow, Aux_ModalidadCashFlow.ModalidadCashFlow, IdiomasModalidadCashFlow.ModalidadCashFlowIdioma AS Traduccion " _
	& " FROM Aux_ModalidadCashFlow "_
		& " LEFT JOIN IdiomasModalidadCashFlow " _ 
			& " ON Aux_ModalidadCashFlow.idModalidadCashFlow = IdiomasModalidadCashFlow.idModalidadCashFlow " _
			& " AND IdiomasModalidadCashFlow.idIdioma=" & idIdiomaCookieCombo _
		& " order by Aux_ModalidadCashFlow.ModalidadCashFlow "

AbrirRecordSet rsAux, sqlQuery, cn_STRING
%>
<%
Dim rsIdioma
sqlQuery="SELECT Idioma From Idiomas Where idIdioma="& idIdiomaCookieCombo
AbrirRecordSet rsIdioma, sqlQuery, cn_STRING
%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Modificar(Cual){
	document.frmNavegar.action="ModalidadCashFlowUpdate.asp";
	document.frmNavegar.idModalidadCashFlow.value=Cual;
	document.frmNavegar.submit();
}

function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="ModalidadCashFlowIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Modalidades CashFlow", "Listado de Modalidades CashFlow" %>
<%Cabecera2 "Cerrar" %>

<br>
  <table border="0" width="100%" cellspacing="0" cellpadding="2">
     <tr>   
 	 	<td class="Cabecera" width="30%"><%=TraducirTexto("Modalidad CashFlow")%></td>
	    <td class="Cabecera" width="65%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
        <td class="Cabecera" colspan="3">&nbsp;</td>
     </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea"><%=(rsAux.Fields.Item("ModalidadCashFlow").Value)%></td>
	   <td class="Linea" ><%=(rsAux.Fields.Item("Traduccion").Value)%>&nbsp;</td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idModalidadCashFlow").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idModalidadCashFlow").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmNavegar" action="" method="post">
	<input type="hidden" name="idModalidadCashFlow" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsIdioma
CerrarRecordSet rsAux
%>
