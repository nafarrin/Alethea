<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idAgrupador, Agrupador,  index
Agrupador=Trim(Request.Form("Agrupador"))
idAgrupador=Trim(Request.Form("idAgrupador"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Agrupador="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Tipo Partida Agrupador")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoPartidaAgrupador SET " _
			& " Agrupador='" & replace(Agrupador, "'","''") & "' "  _
			& " where idAgrupador=" & idAgrupador
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TipoPartidaAgrupador.asp")
	
	end if
	
end if
%>
<%
Dim rsAgrupador__MMColParam
rsAgrupador__MMColParam = "1"
If (Request.Form("idAgrupador") <> "") Then 
  rsAgrupador__MMColParam = Request.Form("idAgrupador")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsAgrupador
	
	sqlQuery="SELECT * FROM Aux_TipoPartidaAgrupador WHERE idAgrupador = "& rsAgrupador__MMColParam &  ""
	
	AbrirRecordSet rsAgrupador, sqlQuery, cn_STRING

	 Agrupador=(rsAgrupador.Fields.Item("Agrupador").Value)
	 
	 CerrarRecordset rsAgrupador
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
	window.location.href="TipoPartidaAgrupador.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Partida Agrupador ", "Modificar Tipo Partida Agrupador" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="AgrupadorUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Agrupador ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo Partida Agrupador")%> *:</td>
      <td> <input type="text" name="Agrupador" value="<%=Agrupador%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idAgrupador" value="<%= Trim(Request.Form("idAgrupador")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
