<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsAux, sqlQuery

sqlQuery="SELECT t.*, i.ClaseNegocioIdioma AS Traduccion FROM Aux_ClaseNegocio t "_
& " LEFT JOIN IdiomasClaseNegocio i ON t.idClaseNegocio = i.idClaseNegocio AND i.idIdioma=" &  Request.Cookies("idIdiomaCookie") _
& " order by t.ClaseNegocio "
'sqlQuery="SELECT *  FROM Aux_ClaseNegocio order by clasenegocio "
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
	document.frmNavegar.action="NegociosClaseUpdate.asp";
	document.frmNavegar.idClaseNegocio.value=Cual;
	document.frmNavegar.submit();
}

function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="ClaseNegocioIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Clase Negocios", "Listado de Clase Negocios" %>
<%Cabecera2 "Cerrar" %>

<br>
  <table border="0" width="100%" cellspacing="0" cellpadding="2">
     <tr>   
 	 	<td class="Cabecera" width="30%"><%=TraducirTexto("Clase Negocio")%></td>
		 <td class="Cabecera" width="65%"><%=(rsIdioma.Fields.Item("Idioma").Value)%>&nbsp;</td>
		<td class="Cabecera" colspan="2">&nbsp;</td>
     </tr>
    <% While (NOT rsAux.EOF) %>
    <tr> 
      <td class="Linea"><%=(rsAux.Fields.Item("ClaseNegocio").Value)%></td>
	   <td class="Linea" ><%=(rsAux.Fields.Item("Traduccion").Value)%>&nbsp;</td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsAux.Fields.Item("idClaseNegocio").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsAux.Fields.Item("idClaseNegocio").Value)%>')" src="../img/Editar.gif"></td>
    </tr>
    <% 
  rsAux.MoveNext()
Wend
%>
  </table>
	<form name="frmNavegar" action="" method="post">
	<input type="hidden" name="idClaseNegocio" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
CerrarRecordSet rsAux
CerrarRecordSet rsIdioma
%>
