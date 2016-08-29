<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoNegocio
idTipoNegocio=Trim(Request.Form("idTipoNegocio"))

dim TipoNegocio,  index, idClaseNegocio
TipoNegocio=Trim(Request.Form("TipoNegocio"))
idClaseNegocio=Trim(Request.Form("idClaseNegocio"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("MM_Insert"))<>"" then
	HayError=false
	if TipoNegocio="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"tipo de negocio")
	end if
	
	if not HayError then
		
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200	
			
		if idTipoNegocio="" then 'insertar nuevo
			
			sqlQuery="INSERT INTO TiposNegocio " _
				& " (TipoNegocio, idClaseNegocio) " _
				& " VALUES " _
				& "('" & replace(TipoNegocio,"'","''") & "','" & idClaseNegocio & "') " 				
		
		else 'modificar 
			sqlQuery="UPDATE TiposNegocio  SET " _
			& " TipoNegocio='" & replace(TipoNegocio,"'","''") &"'," _
			& " idClaseNegocio='" & idClaseNegocio &"'" _
			& " WHERE idTipoNegocio=" & idTipoNegocio
						
		end if
		
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("TiposNegocio.asp?index="&index)
		
		
	end if
	
end if
%>

<%
Dim rsClaseNegocio
Dim rsClaseNegocio_numRows

Set rsClaseNegocio = Server.CreateObject("ADODB.Recordset")
rsClaseNegocio.ActiveConnection = cn_STRING
rsClaseNegocio.Source = "SELECT * FROM dbo.Aux_ClaseNegocio"
rsClaseNegocio.CursorType = 0
rsClaseNegocio.CursorLocation = 2
rsClaseNegocio.LockType = 1
rsClaseNegocio.Open()

rsClaseNegocio_numRows = 0
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
	window.location.href="TiposNegocio.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos de negocio ", "Datos tipo de negocio" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoNegocioDatos.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos tipo de negocio")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo de negocio")%> *:</td>
      <td> <input type="text" name="TipoNegocio" value="<%=TipoNegocio%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
	<tr>
	   <td class="TextoNegrita" align="right" ><%= TraducirTexto("Clase de Negocio")%>:</td>
      <td>
	  	<select name="idClaseNegocio">
	  	  <%While (NOT rsClaseNegocio.EOF)%>
	  	  <option value="<%=(rsClaseNegocio.Fields.Item("idClaseNegocio").Value)%>" <%If (Not isNull(idClaseNegocio)) Then If (CStr(rsClaseNegocio.Fields.Item("idClaseNegocio").Value) = CStr(idClaseNegocio)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=TraducirTexto(rsClaseNegocio.Fields.Item("Clasenegocio").Value)%></option>
	  	  <%
			  rsClaseNegocio.MoveNext()
			Wend
			%>
		</select>
	  </td>	
	</tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
  <input type="hidden" name="idTipoNegocio" value="<%=idTipoNegocio%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
<%
rsClaseNegocio.Close()
Set rsClaseNegocio = Nothing
%>
