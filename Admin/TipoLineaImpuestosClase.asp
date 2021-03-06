<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery="SELECT Aux_TipoLineaImpuestosClase.idClaseImpuestos, Aux_TipoLineaImpuestosClase.ClaseImpuestos, IdiomasTipoLineaImpuestosClase.ClaseImpuestosIdioma AS Traduccion, isnull(PartidasPresupuestarias.Codigopartida + ' ','') + ISNULL(IdiomasPartidasPresupuestarias.PartidaIdioma, PartidasPresupuestarias.Partida) AS Partida " _
	& " FROM Aux_TipoLineaImpuestosClase "_
	& " LEFT JOIN IdiomasTipoLineaImpuestosClase " _
	& " 	ON Aux_TipoLineaImpuestosClase.idClaseImpuestos = IdiomasTipoLineaImpuestosClase.idClaseImpuestos " _
	& " 	AND IdiomasTipoLineaImpuestosClase.idIdioma=" &  idIdiomaCookieCombo _
	& " LEFT JOIN PartidasPresupuestarias " _
	& " 	ON Aux_TipoLineaImpuestosClase.idPartidaDefecto=PartidasPresupuestarias.idPartida " _
	& " LEFT JOIN IdiomasPartidasPresupuestarias " _
	& " 	ON PartidasPresupuestarias.idPartida = IdiomasPartidasPresupuestarias.idPartida " _
	& " 	AND IdiomasPartidasPresupuestarias.idIdioma=" &  idIdiomaCookieCombo _
	& " order by Aux_TipoLineaImpuestosClase.ClaseImpuestos "


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
	document.frmNavegar.action="TipoLineaImpuestosClaseUpdate.asp";
	document.frmNavegar.idClaseImpuestos.value=Cual;
	document.frmNavegar.submit();
}

function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="ClaseImpuestosIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Clase Impuestos", "Listado de Clases Impuestos" %>
<%Cabecera2 "Cerrar" %>

<br>
  <table border="0" width="100%" cellspacing="0" cellpadding="2">
     <tr>   
 	 	<td class="Cabecera" width="30%"><%=TraducirTexto("Clase Impuestos")%></td>
	    <td class="Cabecera" width="35%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
 	 	<td class="Cabecera" width="30%"><%=TraducirTexto("Partida por defecto")%></td>
        <td class="Cabecera" colspan="3">&nbsp;</td>
     </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea"><%=(rsAux.Fields.Item("ClaseImpuestos").Value)%></td>
	   <td class="Linea" >&nbsp;<%=(rsAux.Fields.Item("Traduccion").Value)%></td>
	   <td class="Linea" >&nbsp;<%=(rsAux.Fields.Item("Partida").Value)%></td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idClaseImpuestos").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idClaseImpuestos").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmNavegar" action="" method="post">
	<input type="hidden" name="idClaseImpuestos" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsIdioma
CerrarRecordSet rsAux
%>
