<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idVisualizacion, Visualizacion,  index
Visualizacion=Trim(Request.Form("Visualizacion"))
idVisualizacion=Trim(Request.Form("idVisualizacion"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Visualizacion="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Visualización")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_Visualizacion SET " _
			& " Visualizacion='" & replace(Visualizacion, "'","''") & "' "  _
			& " where idVisualizacion=" & idVisualizacion
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("Visualizacion.asp")
	
	end if
	
end if
%>
<%
Dim rsVisualizacion__MMColParam
rsVisualizacion__MMColParam = "1"
If (Request.Form("idVisualizacion") <> "") Then 
  rsVisualizacion__MMColParam = Request.Form("idVisualizacion")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsVisualizacion
	
	sqlQuery="SELECT * FROM Aux_Visualizacion WHERE idVisualizacion = "& rsVisualizacion__MMColParam &  ""
	
	AbrirRecordSet rsVisualizacion, sqlQuery, cn_STRING

	 Visualizacion=(rsVisualizacion.Fields.Item("Visualizacion").Value)
	 
	 CerrarRecordset rsVisualizacion
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
	window.location.href="Visualizacion.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Visualización ", "Modificar Visualización" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="VisualizacionUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Visualización ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Visualización")%> *:</td>
      <td> <input type="text" name="Visualizacion" value="<%=Visualizacion%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idVisualizacion" value="<%= Trim(Request.Form("idVisualizacion")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
