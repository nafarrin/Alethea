<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idCampoLineaPlantilla, CampoLineaPlantilla,  index
CampoLineaPlantilla=Trim(Request.Form("CampoLineaPlantilla"))
idCampoLineaPlantilla=Trim(Request.Form("idCampoLineaPlantilla"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if CampoLineaPlantilla="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Campo Linea Plantilla")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_CampoLineaPlantilla SET " _
			& " CampoLineaPlantilla='" & replace(CampoLineaPlantilla, "'","''") & "' "  _
			& " where idCampoLineaPlantilla=" & idCampoLineaPlantilla
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("CampoLineaPlantilla.asp")
	
	end if
	
end if
%>
<%
Dim rsCampoLineaPlantilla__MMColParam
rsCampoLineaPlantilla__MMColParam = "1"
If (Request.Form("idCampoLineaPlantilla") <> "") Then 
  rsCampoLineaPlantilla__MMColParam = Request.Form("idCampoLineaPlantilla")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsCampoLineaPlantilla
	
	sqlQuery="SELECT * FROM Aux_CampoLineaPlantilla WHERE idCampoLineaPlantilla = "& rsCampoLineaPlantilla__MMColParam &  ""
	
	AbrirRecordSet rsCampoLineaPlantilla, sqlQuery, cn_STRING

	 CampoLineaPlantilla=(rsCampoLineaPlantilla.Fields.Item("CampoLineaPlantilla").Value)
	 
	 CerrarRecordset rsCampoLineaPlantilla
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
	window.location.href="CampoLineaPlantilla.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Campos Lineas Plantilla", "Modificar Campo Linea Plantilla" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="CampoLineaPlantillaUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Campo Linea Plantilla ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Campo Linea Plantilla")%> *:</td>
      <td> <input type="text" name="CampoLineaPlantilla" value="<%=CampoLineaPlantilla%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idCampoLineaPlantilla" value="<%= Trim(Request.Form("idCampoLineaPlantilla")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
