<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->


<%
dim idApartado
idApartado=Trim(Request.Form("idApartado"))

dim Apartado,  index, Soportado, OrdenConsulta, idVisualizacion, idClaseNegocio, ExportarOLAP, ApartadoImpuestos, SignoOLAP ', ApartadoInteres
dim idTipoApartadoOLAP
Apartado=Trim(Request.Form("Apartado"))
idClaseNegocio=Trim(Request.Form("idClaseNegocio"))
Soportado=Trim(Request.Form("Soportado"))
OrdenConsulta=Trim(Request.Form("OrdenConsulta"))
idVisualizacion=Trim(Request.Form("idVisualizacion"))
ApartadoImpuestos=Trim(Request.Form("ApartadoImpuestos"))
SignoOLAP=Trim(Request.Form("SignoOLAP"))
idTipoApartadoOLAP=Trim(Request.Form("idTipoApartadoOLAP"))
'ApartadoInteres=Trim(Request.Form("ApartadoInteres"))

if Request.Form("ExportarOLAP") then
	ExportarOLAP = "1"
else
	ExportarOLAP = "0"
end if

index=Trim(Request("index"))


dim HayError, msgError, sqlQuery

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Apartado="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Apartado")
	end if	
	
	if not hayError then 
		dim cn
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200
		
		'sqlQuery = "Update Lineas_Apartados set OrdenConsulta=OrdenConsulta+1 where OrdenConsulta>=" & OrdenConsulta
		'cn.execute sqlQuery
		
		if idTipoApartadoOLAP="" then idTipoApartadoOLAP="null"
	
		if idApartado<>"" then 
			sqlQuery="UPDATE Lineas_Apartados set " _
				& " Apartado='" & replace(Apartado, "'","''") & "', " _
				& " Soportado=" & Soportado & ", " _
				& " OrdenConsulta=" & OrdenConsulta & ", " _
				& " idVisualizacion=" & idVisualizacion & ", " _
				& " idClaseNegocio=" & idClaseNegocio & ", " _
				& " ApartadoImpuestos=" & ApartadoImpuestos & ", " _
				& " idTipoApartadoOLAP=" & idTipoApartadoOLAP & ", " _
				& " SignoOLAP=" & SignoOLAP & ", " _
				& " ExportarOLAP=" & ExportarOLAP & " " _	
				& " where idApartado=" & idApartado
'				& " ApartadoInteres=" & ApartadoInteres & ", " _
		else
			sqlquery="INSERT INTO Lineas_Apartados " _
				& " (Apartado, Soportado, OrdenConsulta, idVisualizacion, idClaseNegocio, SignoOLAP, ExportarOLAP, ApartadoImpuestos, idTipoApartadoOLAP) " _
				& " VALUES " _
				& " ('" & replace(Apartado, "'","''") & "', " & Soportado & ", " & OrdenConsulta & ", " & idVisualizacion & ", " & idClaseNegocio & ", " & SignoOLAP & ", " & ExportarOLAP & "," & ApartadoImpuestos & "," & idTipoApartadoOLAP & ") "
			
		end if
	
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("Apartados.asp")
	end if
	
	
end if

Dim rsOrden, MaxValor
if idApartado<>"" then
	sqlQuery="SELECT max(OrdenConsulta) as MaxOrden, Count(*) as NumApartados  FROM Lineas_Apartados"
else
	sqlQuery="SELECT isnull(max(OrdenConsulta),0) +1 as MaxOrden, isnull(Count(*),0) +1 as NumApartados  FROM Lineas_Apartados"
end if
AbrirRecordSet rsOrden, sqlQuery, cn_STRING

if (rsOrden.Fields.Item("MaxOrden").Value)>(rsOrden.Fields.Item("NumApartados").Value) then
	MaxValor=(rsOrden.Fields.Item("MaxOrden").Value)
else
	MaxValor=(rsOrden.Fields.Item("NumApartados").Value)
end if
%>
<%
Dim rsVisualizar
Dim rsVisualizar_numRows

Set rsVisualizar = Server.CreateObject("ADODB.Recordset")
rsVisualizar.ActiveConnection = cn_STRING
rsVisualizar.Source = "SELECT * FROM dbo.Aux_visualizacion"
rsVisualizar.CursorType = 0
rsVisualizar.CursorLocation = 2
rsVisualizar.LockType = 1
rsVisualizar.Open()

rsVisualizar_numRows = 0
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
	window.location.href="Apartados.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Apartados ", "Datos apartado" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>


<form method="POST" action="" name="form1">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos apartado") %></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Apartado")%> *:</td>
      <td> <input type="text" name="Apartado" value="<%=Apartado%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
	<tr>
	   <td class="TextoNegrita" align="right" ><%= TraducirTexto("Clase de negocio")%>:</td>
      <td><%PintarComboTraducido "idClaseNegocio", "Aux_ClaseNegocio", "idClaseNegocio", "ClaseNegocio", "ClaseNegocioIdioma", "IdiomasClaseNegocio", "", idClaseNegocio, CTE_OcultarVacio %></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo apartado")%> *:</td>
      <td><%PintarComboTraducido "Soportado", "Aux_TipoApartado", "idTipoApartado", "TipoApartado", "TipoApartadoIdioma", "IdiomasTipoApartado", "", Soportado, CTE_OcultarVacio%></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Visualizar")%> *:</td>
      <td><select name="idVisualizacion">
        <%While (NOT rsVisualizar.EOF)%>
        <option value="<%=(rsVisualizar.Fields.Item("idVisualizacion").Value)%>" <%If (Not isNull(idVisualizacion)) Then If (CStr(rsVisualizar.Fields.Item("idVisualizacion").Value) = CStr(idVisualizacion)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=TraducirTexto(rsVisualizar.Fields.Item("Visualizacion").Value)%></option>
        <%
		  rsVisualizar.MoveNext()
		Wend
		%>
        </select> </td>
    </tr>    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Orden consulta")%> *:</td>
      <td>
	  	<select name="OrdenConsulta">
		<% 
		dim i
		for i=MaxValor to 1 step-1
		%>
			<option <% if cstr(OrdenConsulta)=cstr(i) then response.Write(" selected ")%> value="<%= i %>"><%= i %></option>
		<%
			next
		%>
		</select>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo apartado OLAP")%> *:</td>
      <td><%PintarComboTraducido "idTipoApartadoOLAP", "Aux_TipoApartado", "idTipoApartado", "TipoApartado", "TipoApartadoIdioma", "IdiomasTipoApartado", "", idTipoApartadoOLAP, CTE_MostrarVacio%></td>
    </tr>
	<tr valign="baseline"> 
		<td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Exportar OLAP")%> *:</td>
		<td> 
			<select name="ExportarOLAP">
				<option value="1" <%If ExportarOLAP="1" then response.Write(" SELECTED ") %> ><%= TraducirTexto("Sí")%></option>
				<option value="0" <%If ExportarOLAP="0" then response.Write(" SELECTED ") %> ><%= TraducirTexto("No")%></option>
			</select>
		</td>
	</tr> 
	<tr valign="baseline"> 
		<td nowrap align="right"  class="TextoNegrita" valign="top"><%= TraducirTexto("Signo exportación OLAP")%> *:
		<br><span class="TextoCash"><%= TraducirTexto("El valor que se muestra en las consultas se multiplicará por este valor")%></span>
		</td>
		<td valign="middle"> 
			<select name="SignoOLAP">
				<option value="1" <%If SignoOLAP="1" then response.Write(" SELECTED ") %> >+1</option>
				<option value="0" <%If SignoOLAP="0" then response.Write(" SELECTED ") %> >0</option>
				<option value="-1" <%If SignoOLAP="-1" then response.Write(" SELECTED ") %> >-1</option>
			</select>
		</td>
	</tr> 
	<tr valign="baseline"> 
		<td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Apartado Impuestos")%>*:</td>
		<td> 
			<select name="ApartadoImpuestos">
				<option value="0" <%If ApartadoImpuestos="0" then response.Write(" SELECTED ") %> ><%= TraducirTexto("No")%></option>
				<option value="1" <%If ApartadoImpuestos="1" then response.Write(" SELECTED ") %> ><%= TraducirTexto("Sí")%></option>
			</select>
		</td>
	</tr> 
	<!--tr valign="baseline"> 
		<td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Apartado Financieros")%>*:</td>
		<td> 
			<select name="ApartadoInteres">
				<option value="0" <%'If ApartadoInteres="0" then response.Write(" SELECTED ") %> ><%'= TraducirTexto("No")%></option>
				<option value="1" <%'If ApartadoInteres="1" then response.Write(" SELECTED ") %> ><%'= TraducirTexto("Sí")%></option>
			</select>
		</td>
	</tr--> 
</table>
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="idApartado" value="<%=idApartado%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
<%
rsOrden.Close()
Set rsOrden = Nothing
%>
<%
rsVisualizar.Close()
Set rsVisualizar = Nothing
%>