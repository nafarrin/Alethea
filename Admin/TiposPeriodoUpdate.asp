<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoPeriodo, TipoPeriodo,  index
TipoPeriodo=Trim(Request.Form("TipoPeriodo"))
idTipoPeriodo=Trim(Request.Form("idTipoPeriodo"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoPeriodo="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Periodo")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoPeriodo SET " _
			& " TipoPeriodo='" & replace(TipoPeriodo, "'","''") & "' "  _
			& " where idTipoPeriodo=" & idTipoPeriodo
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TiposPeriodo.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoPeriodo__MMColParam
rsTipoPeriodo__MMColParam = "1"
If (Request.Form("idTipoPeriodo") <> "") Then 
  rsTipoPeriodo__MMColParam = Request.Form("idTipoPeriodo")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoPeriodo
	
	sqlQuery="SELECT * FROM Aux_TipoPeriodo WHERE idTipoPeriodo = "& rsTipoPeriodo__MMColParam &  ""
	
	AbrirRecordSet rsTipoPeriodo, sqlQuery, cn_STRING

	 TipoPeriodo=(rsTipoPeriodo.Fields.Item("TipoPeriodo").Value)
	 
	 CerrarRecordset rsTipoPeriodo
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
	window.location.href="TiposPeriodo.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos Periodos ", "Modificar Tipo de Periodo" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TiposPeriodoUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de Periodo")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita">Tipo Periodo *:</td>
      <td> <input type="text" name="TipoPeriodo" value="<%=TipoPeriodo%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoPeriodo" value="<%= idTipoPeriodo %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
