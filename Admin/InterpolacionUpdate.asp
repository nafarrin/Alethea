<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idInterpolacion, Interpolacion,  index
Interpolacion=Trim(Request.Form("Interpolacion"))
idInterpolacion=Trim(Request.Form("idInterpolacion"))


index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Interpolacion="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Tipo Apartado")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_Interpolacion SET " _
			& " Interpolacion='" & replace(Interpolacion, "'","''") & "' "  _
			& " where idInterpolacion=" & idInterpolacion
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("Interpolacion.asp")
	
	end if
	
end if
%>
<%
Dim rsInterpolacion__MMColParam
rsInterpolacion__MMColParam = "1"
If (Request.Form("idInterpolacion") <> "") Then 
  rsInterpolacion__MMColParam = Request.Form("idInterpolacion")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsInterpolacion
	
	sqlQuery="SELECT * FROM Aux_Interpolacion WHERE idInterpolacion = "& rsInterpolacion__MMColParam &  ""
	
	AbrirRecordSet rsInterpolacion, sqlQuery, cn_STRING

	 Interpolacion=(rsInterpolacion.Fields.Item("Interpolacion").Value)
	 
	 CerrarRecordset rsInterpolacion
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
	window.location.href="Interpolacion.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Apartados ", "Modificar Interpolación" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="InterpolacionUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Interpolación")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Interpolación")%> *:</td>
      <td> <input type="text" name="Interpolacion" value="<%=Interpolacion%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idInterpolacion" value="<%= idInterpolacion %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
