<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idLogo
idLogo= Request.Form("idLogo")

dim Logo,  index,i, NombreImagen
Logo=Trim(Request.Form("Logo"))

index=Trim(Request("index"))
NombreImagen=Trim(Request.Form("NombreImagen"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Logo="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Logo")
	end if
	if NombreImagen="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_SIMPLE,"Debe especificar un nombre para la imagen")
	end if
	
	if not HayError then
		
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200	
			
		if idLogo="" then 'insertar nuevo
			
			sqlQuery="INSERT INTO Logos(Logo, NombreImagen) VALUES " _
				& "('" & replace(Logo,"'","''") & "','" & replace(NombreImagen,"'","''") & "') " 				
		
		else 'modificar 
			sqlQuery="UPDATE Logos  SET " _
			& " Logo='" & replace(Logo,"'","''") &"'," _
			& " NombreImagen='" & replace(NombreImagen,"'","''") &"'" _
			& " WHERE idLogo=" & idLogo
						
		end if
		
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("Logos.asp?index="&index)
		
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
	window.location.href="Logos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Logos ", "Datos Logo" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="" name="form1">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos Logo")%></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Logo")%> *:</td>
      <td> <input type="text" name="Logo" value="<%=Logo%>" style="width:250px" maxlength="50"></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Nombre Imagen")%> *:</td>
      <td><input type="text" name="NombreImagen" value="<%=NombreImagen%>" style="width:250px" maxlength="50"></td>
    </tr>
  </table>
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="idLogo" value="<%=idLogo%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
