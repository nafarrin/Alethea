<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoFTP, TipoFTP,  index
TipoFTP=Trim(Request.Form("TipoFTP"))
idTipoFTP=Trim(Request.Form("idTipoFTP"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoFTP="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo FTP")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoFTP SET " _
			& " TipoFTP='" & replace(TipoFTP, "'","''") & "' "  _
			& " where idTipoFTP=" & idTipoFTP
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TipoFTP.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoFTP__MMColParam
rsTipoFTP__MMColParam = "1"
If (Request.Form("idTipoFTP") <> "") Then 
  rsTipoFTP__MMColParam = Request.Form("idTipoFTP")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoFTP
	
	sqlQuery="SELECT * FROM Aux_TipoFTP WHERE idTipoFTP = "& rsTipoFTP__MMColParam &  ""
	
	AbrirRecordSet rsTipoFTP, sqlQuery, cn_STRING

	 TipoFTP=(rsTipoFTP.Fields.Item("TipoFTP").Value)
	 
	 CerrarRecordset rsTipoFTP
end if

%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function Guardar(){
	document.form1.submit();
}
function Cerrar(){
	window.location.href="TipoFTP.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo FTP", "Modificar Tipo de FTP" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoFTPUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de FTP")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("TipoFTP")%> *:</td>
      <td> <input type="text" name="TipoFTP" value="<%=TipoFTP%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoFTP" value="<%= idTipoFTP %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
