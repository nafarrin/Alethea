<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery= " SELECT Aux_Visualizacion.idVisualizacion, Aux_Visualizacion.Visualizacion, IdiomasVisualizacion.VisualizacionIdioma AS Traduccion " _
	& " FROM Aux_Visualizacion " _
		& " LEFT JOIN IdiomasVisualizacion " _
			& " ON Aux_Visualizacion.idVisualizacion = IdiomasVisualizacion.idVisualizacion " _
			& " AND IdiomasVisualizacion.idIdioma=" &  idIdiomaCookieCombo _
	& " order by Aux_Visualizacion.Visualizacion "

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
	document.frmDatos.action="VisualizacionUpdate.asp";
	document.frmDatos.idVisualizacion.value=Cual;
	document.frmDatos.submit();
}

function Traducir(Cual){
	document.frmDatos.action="../Idiomas/Traducciones.asp";
	document.frmDatos.idCampoIdioma.value=Cual;
	document.frmDatos.CampoIdioma.value="VisualizacionIdioma";
	document.frmDatos.submit();
}

function Cerrar(){
	document.frmDatos.action="../main.asp";
	document.frmDatos.submit();
}
</script>
</head>
<body><form name="frmDatos" action="" method="post">
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Visualizaciones", "Listado de Visualizaciones" %>
<%Cabecera2 "Cerrar" %>

<br>
  <table border="0" width="100%" cellspacing="0" cellpadding="2">
     <tr>   
 	 	<td class="Cabecera" width="30%"><%=TraducirTexto("Visualización")%></td>
	    <td class="Cabecera" width="65%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
        <td class="Cabecera" colspan="3">&nbsp;</td>
     </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea"><%=(rsAux.Fields.Item("Visualizacion").Value)%></td>
	   <td class="Linea" ><%=(rsAux.Fields.Item("Traduccion").Value)%>&nbsp;</td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idVisualizacion").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idVisualizacion").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmDatos" action="" method="post">
	<input type="hidden" name="idVisualizacion" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsIdioma
CerrarRecordSet rsAux
%>
