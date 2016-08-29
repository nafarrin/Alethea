<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim Unidad,  index, Decimales, Desglosable, idUnidad
idUnidad=Trim(Request.Form("idUnidad"))
Unidad=Trim(Request.Form("Unidad"))
Decimales=Trim(Request.Form("Decimales"))
if not isnumeric(Decimales) then Decimales=2
Desglosable=Trim(Request.Form("Desglosable"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("MM_Insert"))<>"" then
	HayError=false
	if Unidad="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Unidad")
	end if
	
	if not HayError then 
		dim cn, sqlQuery
		CrearConexion cn
		
		
		if idUnidad="" then
			sqlQuery="INSERT INTO Unidades " _
				& "( Unidad, Decimales, Desglosable) " _
				& " VALUES " _
				& "('" &  replace(Unidad,"'","''") &  "', " & Decimales & ", " & Desglosable & ") " 
		else
			sqlQuery="UPDATE Unidades SET " _
				& " Unidad='" & replace(Unidad,"'","''") &"'," _
				& " Decimales=" & Decimales & ", " _
				& " Desglosable=" & Desglosable & " " _
				& " WHERE idUnidad=" & idUnidad
		end if
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("Unidades.asp?index=" &  index )
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
	document.frmDatos.submit();
}
function Cerrar(){
	window.location.href="Unidades.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Unidades", "Datos unidad" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="UnidadNuevo.asp" name="frmDatos">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos unidad")%>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Unidad")%> *:</td>
      <td> <input type="text" name="Unidad" value="<%=Unidad%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Decimales")%> *:</td>
      <td> <select name="Decimales">
          <%
			dim i
			for i=0 to 4%>
          <option value="<%=i%>" <%if cstr(Decimales)=cstr(i) then response.Write(" Selected ")%>><%= i %></option>
          <%next%>
        </select> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Desglosable")%> *:</td>
      <td> <select name="Desglosable">
          <option value="1" <%if cstr(Desglosable)="1" then response.Write(" Selected ")%>><%= TraducirTexto("Sí")%></option>
          <option value="0" <%if cstr(Desglosable)="0" then response.Write(" Selected ")%>><%= TraducirTexto("No")%></option>
        </select> 
	</td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="1">
  <input type="hidden" name="idUnidad" value="<%= idUnidad %>">
  <input type="hidden" name="Index" value="<%= Index %>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
