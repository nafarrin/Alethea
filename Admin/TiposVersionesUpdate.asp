<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoVersion, TipoVersion,  index, ExportarPresupuestoOLAP, ExportarPrevisionOLAP
TipoVersion=Trim(Request.Form("TipoVersion"))
idTipoVersion=Trim(Request.Form("idTipoVersion"))
ExportarPresupuestoOLAP=Trim(Request.Form("ExportarPresupuestoOLAP"))
ExportarPrevisionOLAP=Trim(Request.Form("ExportarPrevisionOLAP"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoVersion="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Versión")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_tipoVersion SET " _
			& " TipoVersion='" & replace(TipoVersion, "'","''") & "', "  _
			& " ExportarPresupuestoOLAP=" & ExportarPresupuestoOLAP & ", " _
			& " ExportarPrevisionOLAP=" & ExportarPrevisionOLAP & " " _
			& " where idTipoVersion=" & idTipoVersion
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TiposVersiones.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoVersion__MMColParam
rsTipoVersion__MMColParam = "1"
If (Request.Form("idTipoVersion") <> "") Then 
  rsTipoVersion__MMColParam = Request.Form("idTipoVersion")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoVersion
	
	sqlQuery="SELECT * FROM Aux_tipoVersion WHERE idTipoVersion = "& rsTipoVersion__MMColParam &  ""
	
	AbrirRecordSet rsTipoVersion, sqlQuery, cn_STRING


	 TipoVersion=(rsTipoVersion.Fields.Item("TipoVersion").Value)
	 ExportarPresupuestoOLAP=(rsTipoVersion.Fields.Item("ExportarPresupuestoOLAP").Value)
	 ExportarPrevisionOLAP=(rsTipoVersion.Fields.Item("ExportarPrevisionOLAP").Value)
	 
	 CerrarRecordset rsTipoVersion
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
	window.location.href="TiposVersiones.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Versiones ", "Modificar Tipo de Versión" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TiposVersionesUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo de Version")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("TipoVersion")%> *:</td>
      <td> <input type="text" name="TipoVersion" value="<%=TipoVersion%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
	<tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Exportar Presupuesto OLAP")%> *:</td>
      <td>
	  <select name="ExportarPresupuestoOLAP">
	  	<option value="1" <%if ExportarPresupuestoOLAP="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
		<option value="0" <%if ExportarPresupuestoOLAP="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
	  </select>
      </td>
    </tr>
	   <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Exportar Prevision OLAP")%> *:</td>
      <td>
	  <select name="ExportarPrevisionOLAP">
	  	<option value="1" <%if ExportarPrevisionOLAP="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
		<option value="0" <%if ExportarPrevisionOLAP="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
	  </select>
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoVersion" value="<%= idTipoVersion %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
