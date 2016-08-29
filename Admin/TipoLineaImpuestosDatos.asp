<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->


<%
dim TipoImpuesto, index, idTipoImpuesto, Liquidacion
dim idClaseImpuestos, idPeriodoLiquidacion, AdmiteDevolucion, idPartidaLiquidacion
TipoImpuesto=Trim(Request.Form("TipoImpuesto"))
idTipoImpuesto=Trim(Request.Form("idTipoImpuesto"))
Liquidacion=Trim(Request.Form("Liquidacion"))
idClaseImpuestos=Trim(Request.Form("idClaseImpuestos"))
idPeriodoLiquidacion=Trim(Request.Form("idPeriodoLiquidacion"))
AdmiteDevolucion=Trim(Request.Form("AdmiteDevolucion"))
idPartidaLiquidacion=Trim(Request.Form("idPartidaLiquidacion"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoImpuesto="" then
		HAyError=true
		msgError=msgError &  MostrarError(CTE_NULO, TraducirTexto("Tipo Impuesto")) 
	end if
	
	if idPartidaLiquidacion="" and Liquidacion="1" then
		HAyError=true
		msgError=msgError &   MostrarError(CTE_NULO, TraducirTexto("Partida liquidación")) 
	end if
	
	if not hayError then 
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200
		
		if idPartidaLiquidacion="" then idPartidaLiquidacion="null"
	
		if idTipoImpuesto<>"" then 
			sqlQuery="UPDATE AUX_TipoLineaImpuestos set " _
				& " TipoImpuesto='" & replace(TipoImpuesto, "'","''") & "', " _
				& " Liquidacion=" & Liquidacion & ", " _
				& " idClaseImpuestos=" & idClaseImpuestos & ", " _
				& " idPeriodoLiquidacion=" & idPeriodoLiquidacion & ", " _
				& " AdmiteDevolucion=" & AdmiteDevolucion & ", " _
				& " idPartidaLiquidacion=" & idPartidaLiquidacion & " " _
				& " where idTipoImpuesto=" & idTipoImpuesto
		else
			sqlquery="INSERT INTO AUX_TipoLineaImpuestos " _
				& " (TipoImpuesto, Liquidacion, idClaseImpuestos, idPeriodoLiquidacion, AdmiteDevolucion, idPartidaLiquidacion) " _
				& " VALUES " _
				& " ('" & replace(TipoImpuesto, "'","''") & "'," & Liquidacion & "," & idClaseImpuestos & "," & idPeriodoLiquidacion & "," & AdmiteDevolucion & "," & idPartidaLiquidacion  & ") "
			
		end if
	
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("TipoLineaImpuestos.asp")
	end if
	
	
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
	window.location.href="TipoLineaImpuestos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: TipoLineaImpuestos ", "Modificar TipoImpuesto" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoLineaImpuestosDatos.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Tipo Impuesto")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo Impuesto")%> *:</td>
      <td> <input type="text" name="TipoImpuesto" value="<%=TipoImpuesto%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Clase de Impuestos")%> *:</td>
      <td><%PintarComboTraducido "idClaseImpuestos", "AUX_TipoLineaImpuestosClase", "idClaseImpuestos", "ClaseImpuestos", "ClaseImpuestosIdioma", "IdiomasTipoLineaImpuestosClase", "", idClaseImpuestos, CTE_OcultarVacio%></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Liquidación Automática")%> *:</td>
      <td>
	  <select name="Liquidacion">
	  	<option value="1" <%if Liquidacion="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
		<option value="0" <%if Liquidacion="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
	  </select>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Partida liquidación")%> *:</td>
      <td><%PintarCombo "idPartidaLiquidacion", "PartidasPresupuestarias LEFT JOIN IdiomasPartidasPresupuestarias ON PartidasPresupuestarias.idPartida = IdiomasPartidasPresupuestarias.idPartida AND IdiomasPartidasPresupuestarias.idIdioma="&idIdiomaCookieCombo&" where idApartado in (SELECT idApartado FROM Lineas_Apartados where ApartadoImpuestos=1)", "PartidasPresupuestarias.idPartida", "CodigoPartida + ' ' + ISNULL(IdiomasPartidasPresupuestarias.PartidaIdioma, PartidasPresupuestarias.Partida)", idPartidaLiquidacion, CTE_MostrarVacio%></td>
    </tr>
     <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Periodo liquidación")%> *:</td>
      <td><%PintarComboTOTALTraducido "idPeriodoLiquidacion", "Periodos", "idPeriodo", "Periodo", "PeriodoIdioma", "IdiomasPeriodos", "", idPeriodoLiquidacion, CTE_OcultarVacio, "", "NumMeses" %></td>
    </tr>
   <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Admite devolución")%> *:</td>
      <td>
	  <select name="AdmiteDevolucion">
	  	<option value="1" <%if AdmiteDevolucion="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
		<option value="0" <%if AdmiteDevolucion="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
	  </select>
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoImpuesto" value="<%=idTipoImpuesto %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
