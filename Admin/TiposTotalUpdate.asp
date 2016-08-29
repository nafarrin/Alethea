<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoTotal, TipoTotal,  index
TipoTotal=Trim(Request.Form("TipoTotal"))
idTipoTotal=Trim(Request.Form("idTipoTotal"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoTotal="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Total")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoTotal SET " _
			& " TipoTotal='" & replace(TipoTotal, "'","''") & "' "  _
			& " where idTipoTotal=" & idTipoTotal
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TiposTotal.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoTotal__MMColParam
rsTipoTotal__MMColParam = "1"
If (Request.Form("idTipoTotal") <> "") Then 
  rsTipoTotal__MMColParam = Request.Form("idTipoTotal")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoTotal
	
	sqlQuery="SELECT * FROM Aux_TipoTotal WHERE idTipoTotal = "& rsTipoTotal__MMColParam &  ""
	
	AbrirRecordSet rsTipoTotal, sqlQuery, cn_STRING

	 TipoTotal=(rsTipoTotal.Fields.Item("TipoTotal").Value)
	 
	 CerrarRecordset rsTipoTotal
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
	window.location.href="TiposTotal.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos Totales ", "Modificar Tipo de Total" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TiposTotalUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de Total")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita">Tipo Total *:</td>
      <td> <input type="text" name="TipoTotal" value="<%=TipoTotal%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoTotal" value="<%= idTipoTotal %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
