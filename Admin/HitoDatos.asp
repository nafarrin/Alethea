<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idHito
idHito= Request.Form("idHito")

dim Hito,  index, Orden,i, Color
Hito=Trim(Request.Form("Hito"))

index=Trim(Request("index"))
Orden=Trim(Request.Form("Orden"))
Color=Trim(Request.Form("Color"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Hito="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"hito")
	end if
	if Color="" then
		HAyError=true
		msgError=msgError & "<br>"& TraducirTexto("Debe especificar un color para el planning") & "."
	end if
	
	if not HayError then
		
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200	
			
		if idHito="" then 'insertar nuevo
			
			sqlQuery="INSERT INTO Hitos VALUES " _
				& "('" & replace(Hito,"'","''") & "','" &  replace(Orden,"'","''") & "','" & Color & "') " 				
		
		else 'modificar 
			sqlQuery="UPDATE Hitos  SET " _
			& " Hito='" & replace(Hito,"'","''") &"'," _
			& " Orden='" & replace(Orden,"'","''") &"'," _
			& " Color='" & replace(Color,"'","''") &"'" _
			& " WHERE idHito=" & idHito
						
		end if
		
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("Hitos.asp?index="&index)
		
	end if	
	
	
end if
%>


<%
Dim rsNumHitos
Dim rsNumHitos_numRows

Set rsNumHitos = Server.CreateObject("ADODB.Recordset")
rsNumHitos.ActiveConnection = cn_STRING
rsNumHitos.Source = "SELECT isnull(max(orden),0) as OrdenMAX FROM Hitos"
rsNumHitos.CursorType = 0
rsNumHitos.CursorLocation = 2
rsNumHitos.LockType = 1
rsNumHitos.Open()

rsNumHitos_numRows = 0
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
	window.location.href="Hitos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Hitos ", "Datos hito" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="" name="form1">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos hito")%></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Hito")%> *:</td>
      <td> <input type="text" name="Hito" value="<%=Hito%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Orden")%> *:</td>
      <td> <select name="Orden">
          <% for i = (rsNumHitos.Fields.Item("OrdenMax").Value)+1 to 1 step-1%>
          <option <%if cstr(Orden)=cstr(i) then response.Write(" selected ")%> value="<%= i %>"><%= i %></option>
          <%next%>
        </select> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"  valign="top"><%= TraducirTexto("Color planning")%> *:</td>
      <td> 
		  <table>
		  <%
		  dim j 
		  for i =1 to  4
		  %><tr><%
		  	for j=1 to 5		
		  	%><td><input class="checkbox" <%if cstr(Color)=cstr(5*(i-1)+j) then response.Write(" checked ")%> name="Color" type="radio" value="<%=5*(i-1)+j%>"><img src="../img/Hitos/Hito<%=5*(i-1)+j%>.gif"></td><%
			next
			%></tr><%
		  next
		  %>
		  </table>
	  </td>
    </tr>
  </table>
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="idHito" value="<%=idHito%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
<%
rsNumHitos.Close()
Set rsNumHitos = Nothing
%>
