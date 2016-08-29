<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idClaseImpuestos, ClaseImpuestos,  index, idPartidaDefecto
ClaseImpuestos=Trim(Request.Form("ClaseImpuestos"))
idClaseImpuestos=Trim(Request.Form("idClaseImpuestos"))
idPartidaDefecto=Trim(Request.Form("idPartidaDefecto"))
if not isnumeric(idPartidaDefecto) then idPartidaDefecto=""
index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if ClaseImpuestos="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Clase de Impuestos")
	end if
	
	if not HayError then 
		dim  cn
		
		if idPartidaDefecto="" then idPartidaDefecto="null"
		
		sqlQuery="Update Aux_TipoLineaImpuestosClase SET " _
			& " ClaseImpuestos='" & replace(ClaseImpuestos, "'","''") & "', "  _
			& " idPartidaDefecto=" & idPartidaDefecto _
			& " where idClaseImpuestos=" & idClaseImpuestos
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TipoLineaImpuestosClase.asp")
	
	end if
	
end if
%>
<%
Dim rsClaseImpuestos__MMColParam
rsClaseImpuestos__MMColParam = "1"
If (Request.Form("idClaseImpuestos") <> "") Then 
  rsClaseImpuestos__MMColParam = Request.Form("idClaseImpuestos")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsClaseImpuestos
	
	sqlQuery="SELECT * FROM Aux_TipoLineaImpuestosClase WHERE idClaseImpuestos = "& rsClaseImpuestos__MMColParam &  ""
	
	AbrirRecordSet rsClaseImpuestos, sqlQuery, cn_STRING

	 ClaseImpuestos=(rsClaseImpuestos.Fields.Item("ClaseImpuestos").Value)
	 
	 idPartidaDefecto=(rsClaseImpuestos.Fields.Item("idPartidaDefecto").Value)
	 
	 
	 CerrarRecordset rsClaseImpuestos
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
	window.location.href="TipoLineaImpuestosClase.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Clase Impuestos ", "Modificar Clase Impuestos" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoLineaImpuestosClaseUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar ClaseImpuestos ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Clase de Impuestos")%> *:</td>
      <td> <input type="text" name="ClaseImpuestos" value="<%=ClaseImpuestos%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Partida por defecto")%>:</td>
      <td><%PintarCombo "idPartidaDefecto", "PartidasPresupuestarias where idApartado in (Select idApartado from Lineas_Apartados where ApartadoImpuestos=1) ", "idPartida", "isnull(Codigopartida + ' ','') + Partida", idPartidaDefecto, CTE_MostrarVacio %></td>
    </tr>
  </table>
  <input type="hidden" name="idClaseImpuestos" value="<%= Trim(Request.Form("idClaseImpuestos")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
