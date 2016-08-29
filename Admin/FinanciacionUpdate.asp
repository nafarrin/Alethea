<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idFinanciacion, Financiacion,  index
Financiacion=Trim(Request.Form("Financiacion"))
idFinanciacion=Trim(Request.Form("idFinanciacion"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Financiacion="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Financiación")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TiposFinanciacion SET " _
			& " Financiacion='" & replace(Financiacion, "'","''") & "' "  _
			& " where idFinanciacion=" & idFinanciacion
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("Financiacion.asp")
	
	end if
	
end if
%>
<%
Dim rsFinanciacion__MMColParam
rsFinanciacion__MMColParam = "1"
If (Request.Form("idFinanciacion") <> "") Then 
  rsFinanciacion__MMColParam = Request.Form("idFinanciacion")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsFinanciacion
	
	sqlQuery="SELECT * FROM Aux_TiposFinanciacion WHERE idFinanciacion = "& rsFinanciacion__MMColParam &  ""
	
	AbrirRecordSet rsFinanciacion, sqlQuery, cn_STRING

	 Financiacion=(rsFinanciacion.Fields.Item("Financiacion").Value)
	 
	 CerrarRecordset rsFinanciacion
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
	window.location.href="Financiacion.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Apartados ", "Modificar Financiación" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="FinanciacionUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Financiación")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Financiación")%> *:</td>
      <td> <input type="text" name="Financiacion" value="<%=Financiacion%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idFinanciacion" value="<%= idFinanciacion %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
