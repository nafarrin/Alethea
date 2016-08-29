<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoApartado, TipoApartado,  index
TipoApartado=Trim(Request.Form("TipoApartado"))
idTipoApartado=Trim(Request.Form("idTipoApartado"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoApartado="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Apartado")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoApartado SET " _
			& " TipoApartado='" & replace(TipoApartado, "'","''") & "' "  _
			& " where idTipoApartado=" & idTipoApartado
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TipoApartados.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoApartado__MMColParam
rsTipoApartado__MMColParam = "1"
If (Request.Form("idTipoApartado") <> "") Then 
  rsTipoApartado__MMColParam = Request.Form("idTipoApartado")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoApartado
	
	sqlQuery="SELECT * FROM Aux_TipoApartado WHERE idTipoApartado = "& rsTipoApartado__MMColParam &  ""
	
	AbrirRecordSet rsTipoApartado, sqlQuery, cn_STRING

	 TipoApartado=(rsTipoApartado.Fields.Item("TipoApartado").Value)
	 
	 CerrarRecordset rsTipoApartado
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
	window.location.href="TipoApartados.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Apartados ", "Modificar Tipo de Apartado" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoApartadosUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de Apartado")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita">TipoApartado *:</td>
      <td> <input type="text" name="TipoApartado" value="<%=TipoApartado%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoApartado" value="<%= idTipoApartado %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
