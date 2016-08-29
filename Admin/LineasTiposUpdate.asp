<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoLinea, TipoLinea,  index
TipoLinea=Trim(Request.Form("TipoLinea"))
idTipoLinea=Trim(Request.Form("idTipoLinea"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoLinea="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Linea")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Lineas_Tipos SET " _
			& " TipoLinea='" & replace(TipoLinea, "'","''") & "' "  _
			& " where idTipoLinea=" & idTipoLinea
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("LineasTipos.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoLinea__MMColParam
rsTipoLinea__MMColParam = "1"
If (Request.Form("idTipoLinea") <> "") Then 
  rsTipoLinea__MMColParam = Request.Form("idTipoLinea")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoLinea
	
	sqlQuery="SELECT * FROM Lineas_Tipos WHERE idTipoLinea = "& rsTipoLinea__MMColParam &  ""
	
	AbrirRecordSet rsTipoLinea, sqlQuery, cn_STRING

	 TipoLinea=(rsTipoLinea.Fields.Item("TipoLinea").Value)
	 
	 CerrarRecordset rsTipoLinea
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
	window.location.href="LineasTipos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos Líneas ", "Modificar Tipo de Línea" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="LineasTiposUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de Línea")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita">Tipo Línea *:</td>
      <td> <input type="text" name="TipoLinea" value="<%=TipoLinea%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoLinea" value="<%= idTipoLinea %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
