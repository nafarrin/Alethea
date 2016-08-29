<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idUso
idUso= Request.Form("idUso")

dim Uso,  index
Uso=Trim(Request.Form("Uso"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Uso="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Uso")
	end if
	
	if not HayError then
		
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200	
			
		if idUso="" then 'insertar nuevo
			
			sqlQuery="INSERT INTO Lineas_Usos VALUES " _
				& "('" &  replace(Uso,"'","''") & "') " 				
		
		else 'modificar 
			sqlQuery="UPDATE Lineas_Usos SET " _
			& " Uso='" & replace(Uso,"'","''") &"'" _
			& " WHERE idUso=" & idUso
						
		end if
		
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("Usos.asp?index="&index)
		
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
	window.location.href="Usos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Usos", "Datos uso" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>
<form method="post" action="" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		<td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos uso")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Uso")%> *:</td>
      <td> <input type="text" name="Uso" value="<%=Uso%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="idUso" value="<%=idUso %>" size="32">
  
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
