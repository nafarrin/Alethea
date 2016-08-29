<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idClaseNegocio, ClaseNegocio,  index
ClaseNegocio=Trim(Request.Form("ClaseNegocio"))
idClaseNegocio=Trim(Request.Form("idClaseNegocio"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if ClaseNegocio="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Clase de Negocio")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_ClaseNegocio SET " _
			& " ClaseNegocio='" & replace(ClaseNegocio, "'","''") & "' "  _
			& " where idClaseNegocio=" & idClaseNegocio
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("NegociosClase.asp")
	
	end if
	
end if
%>
<%
Dim rsClaseNegocio__MMColParam
rsClaseNegocio__MMColParam = "1"
If (Request.Form("idClaseNegocio") <> "") Then 
  rsClaseNegocio__MMColParam = Request.Form("idClaseNegocio")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsClaseNegocio
	
	sqlQuery="SELECT * FROM Aux_ClaseNegocio WHERE idClaseNegocio = "& rsClaseNegocio__MMColParam &  ""
	
	AbrirRecordSet rsClaseNegocio, sqlQuery, cn_STRING

	 ClaseNegocio=(rsClaseNegocio.Fields.Item("ClaseNegocio").Value)
	 
	 CerrarRecordset rsClaseNegocio
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
	window.location.href="NegociosClase.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Clase Negocio ", "Modificar Clase Negocio" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="NegociosClaseUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Clase Negocio")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Clase de Negocio")%> *:</td>
      <td> <input type="text" name="ClaseNegocio" value="<%=ClaseNegocio%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idClaseNegocio" value="<%= Trim(Request.Form("idClaseNegocio")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
