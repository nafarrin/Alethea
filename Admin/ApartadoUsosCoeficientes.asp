<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idApartado
idApartado=Trim(Request.Form("idApartado"))


dim HayError, msgError
HayError=false

if Trim(Request.Form("Insertar"))<>"" then
	'comprobar valores
	dim x
	for each x in request.Form
		if ucase(mid(x,1,len("Uso_")))=ucase("Uso_") then
			if not isnumeric(request.Form(x)) then
				HayError=true
				msgError="Debe poner un valor numérico para todos los valores."
			else
				if cdbl(request.Form(x))<1 then
					HayError=true
					msgError="Los coeficientes tienen que ser mayores o iguale que 1."
				end if
			end if
		end if
	next
	
	if not HayError then
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200

		cn.BeginTrans
		
		'1 Borrar las anteriores
		sqlQuery="Delete  from Apartados_Usos_Coeficientes where idApartado=" & idApartado
		cn.execute sqlQuery

		dim idUso
		'2 poner los nuevos valores
		for each x in request.Form
			if ucase(mid(x,1,len("Uso_")))=ucase("Uso_") then
				idUso=mid(x,len("Uso_")+1)
				sqlQuery ="Insert into Apartados_Usos_Coeficientes (idApartado, idUso, Coeficiente) " _
					& " VALUES " _
					& "(" & idApartado & ", " & idUso & ", " & replace(Trim(Request.Form(x)),",",".") & ") " 
				cn.execute sqlQuery
			end if
		next

		
		
		if err=0  then
			cn.CommitTrans
		Else
			HayError=true
			cn.RollbackTrans
		end if
		
		set cn=nothing

		if Not HayError then response.Redirect("Apartados.asp?index=" & Trim(Request.Form("index")))
	
	end if
end if
%>
<%
Dim rsApartado__MMColParam
rsApartado__MMColParam = "1"
If (Request.Form("idApartado") <> "") Then 
  rsApartado__MMColParam = Request.Form("idApartado")
End If
%>
<%
Dim rsApartado
Dim rsApartado_numRows

Set rsApartado = Server.CreateObject("ADODB.Recordset")
rsApartado.ActiveConnection = cn_STRING
rsApartado.Source = "SELECT * FROM dbo.Lineas_Apartados WHERE idApartado = " + Replace(rsApartado__MMColParam, "'", "''") + ""
rsApartado.CursorType = 0
rsApartado.CursorLocation = 2
rsApartado.LockType = 1
rsApartado.Open()

rsApartado_numRows = 0
%>
<%
Dim rsUsos
Dim rsUsos_numRows

Set rsUsos = Server.CreateObject("ADODB.Recordset")
rsUsos.ActiveConnection = cn_STRING
rsUsos.Source = "SELECT * FROM dbo.Lineas_Usos ORDER BY Uso ASC"
rsUsos.CursorType = 0
rsUsos.CursorLocation = 2
rsUsos.LockType = 1
rsUsos.Open()

rsUsos_numRows = 0
%>
<%
Dim rsValores__par1
rsValores__par1 = "-1"
If (idApartado <> "") Then 
  rsValores__par1 = idApartado
End If
%>
<%
Dim rsValores
Dim rsValores_numRows

Set rsValores = Server.CreateObject("ADODB.Recordset")
rsValores.ActiveConnection = cn_STRING
rsValores.Source = "SELECT *  FROM dbo.Apartados_Usos_Coeficientes  WHERE idApartado = " + Replace(rsValores__par1, "'", "''") + ""
rsValores.CursorType = 0
rsValores.CursorLocation = 2
rsValores.LockType = 1
rsValores.Open()

rsValores_numRows = 0
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
	window.location.href="Apartados.asp?index=<%= Trim(Request.Form("index")) %>"
}
</script>

<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Apartados ", "Coeficientes apartado/ usos" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<table width="100%" class="Casilla">
	<tr>
		<td class="TextoNegrita" width="10%"><%=TraducirTexto("Apartado")%>:</td>
		<td class="Texto"><%=(rsApartado.Fields.Item("Apartado").Value)%></td>
	</tr>
</table>
<br>
<%
dim  NumCols, MaxCols
MaxCols=3

dim Valores()
redim Valores(0)
while not rsUsos.eof 
	if clng((rsUsos.Fields.Item("idUso").Value))>ubound(Valores) then
		redim preserve Valores(clng((rsUsos.Fields.Item("idUso").Value)))
	end if
	if Trim(Request.Form("Insertar"))="" then 
		rsValores.filter="idUso=" & (rsUsos.Fields.Item("idUso").Value)
		if not rsValores.eof then
			Valores(clng((rsUsos.Fields.Item("idUso").Value)))=(rsValores.Fields.Item("Coeficiente").Value)
		else
			Valores(clng((rsUsos.Fields.Item("idUso").Value)))=1
		end if
	else
		Valores(clng((rsUsos.Fields.Item("idUso").Value)))=Trim(Request.Form("Uso_" & (rsUsos.Fields.Item("idUso").Value)))
	end if
	rsUsos.movenext
wend
rsUsos.filter=""
%>
<form method="POST" action="ApartadoUsosCoeficientes.asp" name="frmDatos">
<table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="<%= MaxCols*2 %>" class="Cabecera" ><img src="../img/Vacio.gif">Coeficientes por Uso</td>
    </tr>
	<%
	NumCols=0
	while not rsUsos.eof
		NumCols=Numcols+1
		if NumCols=1 then
		%>
		<tr>
		<%
		end if
		%>
			<td class="TextoNegrita"><%=(rsUsos.Fields.Item("Uso").Value)%>:</td>
			<td><input type="text" name="Uso_<%=(rsUsos.Fields.Item("idUso").Value)%>" value="<%= Valores(rsUsos.Fields.Item("idUso").Value) %>" maxlength="4" size="5"></td>
		<%
		if Numcols=MaxCols then
			Numcols=0
		%>
		</tr>
		<%
		end if
		rsUsos.movenext
	wend
	if Numcols<>0 and Numcols<>MaxCols then
		%>
		</tr>
		<%
	end if
	%>
</table>
  <span class="TextoRojo"><%= msgError %></span> 
  <input type="hidden" name="idApartado" value="<%=idApartado%>" size="32">
	<input type="hidden" name="Index" value="<%= Trim(Request.Form("Index")) %>">
	<input type="hidden" name="Insertar" value="1">
</form>

</body>
</html>
<%
rsApartado.Close()
Set rsApartado = Nothing
%>
<%
rsUsos.Close()
Set rsUsos = Nothing
%>
<%
rsValores.Close()
Set rsValores = Nothing
%>
