<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idSubApartado
idSubApartado=Trim(Request.Form("idSubApartado"))

dim SubApartado,  index, idApartado, OrdenConsultaSub
SubApartado=Trim(Request.Form("SubApartado"))

dim Tesoreria
Tesoreria=Trim(Request.Form("Tesoreria"))

index=Trim(Request("index"))
idApartado=Trim(Request.Form("idApartado"))
OrdenConsultaSub=Trim(Request.Form("OrdenConsultaSub"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if SubApartado="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"subapartado")
	end if
	if idApartado="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Apartado") 
	end if
	
	if not hayError then 
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200
	
		if idSubApartado<>"" then 
			sqlQuery="UPDATE SubApartados set " _
				& " SubApartado='" & replace(SubApartado, "'","''") & "', " _
				& " idApartado=" & idApartado & ", " _
				& " OrdenConsultaSub=" & OrdenConsultaSub & ", " _
				& " Tesoreria=" & Tesoreria & " " _
				& " where idSubApartado=" & idSubApartado
		else
			sqlquery="INSERT INTO SubApartados " _
				& " VALUES " _
				& " ('" & replace(SubApartado, "'","''") & "','" & idApartado & " ','" & OrdenConsultaSub & "',' "&Tesoreria& "')"
		end if

		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("SubApartados.asp?index="& index)
	end if
	
	
	
end if
%>

<%
Dim rsApartados

sqlQuery = "SELECT lineas_apartados.idApartado, ISNULL(IdiomasApartados.ApartadoIdioma, lineas_apartados.Apartado) AS Apartado " _
			& " FROM Lineas_Apartados " _ 
				& " LEFT JOIN IdiomasApartados " _
					& " ON lineas_apartados.idApartado = IdiomasApartados.idApartado " _
					& " AND idiomasApartados.idIdioma=" & idIdiomaCookieCombo _
				& " ORDER BY Apartado ASC"

AbrirRecordSet rsApartados, sqlQuery, cn_STRING
%>
<%
Dim rsOrden__par1
rsOrden__par1 = "-1"
If (idApartado <> "") Then 
  rsOrden__par1 = idApartado
End If
%>
<%
Dim rsOrden
Dim rsOrden_numRows
dim sqlAux
if idSubApartado="" then sqlAux= "+ 1" 

Set rsOrden = Server.CreateObject("ADODB.Recordset")
rsOrden.ActiveConnection = cn_STRING
rsOrden.Source = "SELECT max(OrdenConsultaSub)" & sqlAux & " as MaxOrden  FROM dbo.SubApartados  where idApartado=" + Replace(rsOrden__par1, "'", "''") + ""
rsOrden.CursorType = 0
rsOrden.CursorLocation = 2
rsOrden.LockType = 1
rsOrden.Open()

rsOrden_numRows = 0
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
function Recargar(){
	document.form1.Insertar.value="";
	document.form1.OrdenConsultaSub.value="";
	
	document.form1.submit();
}

function Cerrar(){
	window.location.href="SubApartados.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Subapartados ", "Datos Subapartado" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="SubApartadoDatos.asp" name="form1">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos subapartado")%>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("subapartado")%> *:</td>
      <td> <input type="text" name="SubApartado" value="<%=SubApartado%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("apartado")%> *:</td>
      <td> <select name="idApartado" onChange="Recargar()">
          <option value="" <%If (Not isNull(idApartado)) Then If ("" = CStr(idApartado)) Then Response.Write("SELECTED") : Response.Write("")%>></option>
          <%
While (NOT rsApartados.EOF)
%>
          <option value="<%=(rsApartados.Fields.Item("idApartado").Value)%>" <%If (Not isNull(idApartado)) Then If (CStr(rsApartados.Fields.Item("idApartado").Value) = CStr(idApartado)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=TraducirTexto(rsApartados.Fields.Item("Apartado").Value)%></option>
          <%
  rsApartados.MoveNext()
Wend
If (rsApartados.CursorType > 0) Then
  rsApartados.MoveFirst
Else
  rsApartados.Requery
End If
%>
        </select> </td>
    </tr>
	<tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Orden consulta")%> *:</td>
      <td>
	  	<select name="OrdenConsultaSub">
		<% 
		dim i, j
		if not isnull(rsOrden.Fields.Item("MaxOrden").Value) then 
			j=rsOrden.Fields.Item("MaxOrden").Value
		else
			j=1
		end if
		for i=j to 1 step-1
		%>
			<option <% if cstr(OrdenConsultaSub)=cstr(i) then response.Write(" selected ")%> value="<%= i %>"><%= i %></option>
		<%
			next
		%>
		</select>
	  </td>
    </tr>
	<tr valign="baseline"> 
      <td nowrap align="right" class="TextoNegrita"><%= TraducirTexto("Tesorería")%> *:</td>
	  <td>
		<select name="Tesoreria">
		 <option value="0" <%if cstr(Tesoreria)="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
         <option value="1" <%if cstr(Tesoreria)="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>		
		</select>
      </td>
    </tr>
  </table>
  <input type="hidden" name="Insertar" value="form1">
  <input type="hidden" name="idSubApartado" value="<%=idSubApartado%>">
  <input type="hidden" name="index" value="<%=index%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
<%
rsApartados.Close()
Set rsApartados = Nothing
%>
<%
rsOrden.Close()
Set rsOrden = Nothing
%>
