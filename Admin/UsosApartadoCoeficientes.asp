<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idUso
idUso=Trim(Request.Form("idUso"))


dim HayError, msgError
HayError=false

if Trim(Request.Form("Insertar"))<>"" then
	'comprobar valores
	dim x
	for each x in request.Form
		if ucase(mid(x,1,len("Apartado_")))=ucase("Apartado_") then
			if not isnumeric(request.Form(x)) then
				HayError=true
				msgError=TraducirTexto("Debe poner un valor numérico para todos los coeficientes") & "."
			else
				if cdbl(request.Form(x))<1 then
					HayError=true
					msgError=TraducirTexto("Los coeficientes tienen que ser mayores o iguales que 1") & "."
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
		sqlQuery="Delete  from Apartados_Usos_Coeficientes where idUso=" & idUso
		cn.execute sqlQuery

		dim idApartado
		'2 poner los nuevos valores
		for each x in request.Form
			if ucase(mid(x,1,len("Apartado_")))=ucase("Apartado_") then
				idApartado=mid(x,len("Apartado_")+1)
				sqlQuery ="Insert into Apartados_Usos_Coeficientes (idUso, idApartado, Coeficiente) " _
					& " VALUES " _
					& "(" & idUso & ", " & idApartado & ", " & replace(Trim(Request.Form(x)),",",".") & ") " 
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

		if Not HayError then response.Redirect("Usos.asp?index=" & Trim(Request.Form("index")))
	
	end if
end if
%>
<%
Dim rsUso__par1
rsUso__par1 = "-1"
If (idUso <> "") Then 
  rsUso__par1 = idUso
End If
%>
<%
Dim rsUso
Dim rsUso_numRows

sqlQuery = "SELECT Lineas_Usos.idUso, ISNULL(IdiomasUsos.UsoIdioma, Lineas_Usos.Uso) AS Uso FROM Lineas_Usos "_
	& " LEFT JOIN IdiomasUsos " _
		& " ON Lineas_Usos.idUso = IdiomasUsos.idUso " _
		& " AND IdiomasUsos.idIdioma=" &  idIdiomaCookieCombo _
	& " WHERE Lineas_Usos.idUso = " + Replace(rsUso__par1, "'", "''") _
	& " order by Uso "
	
AbrirRecordSet rsUso, sqlQuery, cn_STRING

rsUso_numRows = 0
%>
<%
Dim rsApartados
Dim rsApartados_numRows

sqlQuery = "SELECT Lineas_Apartados.idApartado, ISNULL(IdiomasApartados.ApartadoIdioma, Lineas_Apartados.Apartado) AS Apartado " _
	& " FROM Lineas_Apartados " _
		& " LEFT JOIN IdiomasApartados " _
			& " ON Lineas_Apartados.idApartado = IdiomasApartados.idApartado " _
			& " AND IdiomasApartados.idIdioma=" & idIdiomaCookieCombo _
		& " ORDER BY Apartado "

AbrirRecordSet rsApartados, sqlQuery, cn_STRING

%>
<%
Dim rsValores__par1
rsValores__par1 = "-1"
If (idUso <> "") Then 
  rsValores__par1 = idUso
End If
%>
<%
Dim rsValores
Dim rsValores_numRows

Set rsValores = Server.CreateObject("ADODB.Recordset")
rsValores.ActiveConnection = cn_STRING
rsValores.Source = "SELECT *  FROM dbo.Apartados_Usos_Coeficientes  WHERE idUso = " + Replace(rsValores__par1, "'", "''") + ""
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
	window.location.href="Usos.asp?index=<%= Trim(Request.Form("index")) %>"
}
</script>

<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Usos", "Coeficientes Uso/ Apartados" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<table width="100%" class="Casilla">
	<tr>
		<td class="TextoNegrita" width="10%"><%= TraducirTexto("Uso")%>:</td>
		<td class="Texto"><%=(rsUso.Fields.Item("Uso").Value)%></td>
	</tr>
</table>
<br>
<%
dim  NumCols, MaxCols
MaxCols=3

dim Valores()
redim Valores(0)
while not rsApartados.eof 
	if clng((rsApartados.Fields.Item("idApartado").Value))>ubound(Valores) then
		redim preserve Valores(clng((rsApartados.Fields.Item("idApartado").Value)))
	end if
	if Trim(Request.Form("Insertar"))="" then 
		rsValores.filter="idApartado=" & (rsApartados.Fields.Item("idApartado").Value)
		if not rsValores.eof then
			Valores(clng((rsApartados.Fields.Item("idApartado").Value)))=(rsValores.Fields.Item("Coeficiente").Value)
		else
			Valores(clng((rsApartados.Fields.Item("idApartado").Value)))=1
		end if
	else
		Valores(clng((rsApartados.Fields.Item("idApartado").Value)))=Trim(Request.Form("Apartado_" & (rsApartados.Fields.Item("idApartado").Value)))
	end if
	rsApartados.movenext
wend
rsApartados.filter=""
%>
<form method="POST" action="UsosApartadoCoeficientes.asp" name="frmDatos">
<table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="<%= MaxCols*2 %>" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Coeficientes por Apartado")%></td>
    </tr>
	<%
	NumCols=0
	while not rsApartados.eof
		NumCols=Numcols+1
		if NumCols=1 then
		%>
		<tr>
		<%
		end if
		%>
			<td class="TextoNegrita"><%=(rsApartados.Fields.Item("Apartado").Value)%>:</td>
			<td><input type="text" name="Apartado_<%=(rsApartados.Fields.Item("idApartado").Value)%>" value="<%= Valores(rsApartados.Fields.Item("idApartado").Value) %>" maxlength="4" size="5"></td>
		<%
		if Numcols=MaxCols then
			Numcols=0
		%>
		</tr>
		<%
		end if
		rsApartados.movenext
	wend
	if Numcols<>0 and Numcols<>MaxCols then
		%>
		</tr>
		<%
	end if
	%>
</table>
  <span class="TextoRojo"><%= msgError %></span> 
  <input type="hidden" name="idUso" value="<%=idUso%>" size="32">
	<input type="hidden" name="Index" value="<%= Trim(Request.Form("Index")) %>">
	<input type="hidden" name="Insertar" value="1">
</form>

</body>
</html>
<%
rsUso.Close()
Set rsUso = Nothing
%>
<%
rsApartados.Close()
Set rsApartados = Nothing
%>
<%
rsValores.Close()
Set rsValores = Nothing
%>
