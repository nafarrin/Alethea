<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim Apartado,  sqlQuery, idClaseNegocio, Soportado
Apartado=Trim(Request("Apartado"))
idClaseNegocio=Trim(Request("idClaseNegocio"))
Soportado=Trim(Request.Form("Soportado"))

if Apartado<>"" then
	sqlQuery= sqlQuery &" and ISNULL(IdiomasApartados.ApartadoIdioma, Lineas_Apartados.Apartado) like '%" & replace(Apartado,"'","''") & "%' "
end if
if idClaseNegocio<>"" then
	sqlQuery= sqlQuery &" and Lineas_Apartados.idClaseNegocio=" & idClaseNegocio
end if
if Soportado<>"" then
	sqlQuery= sqlQuery &" and Lineas_Apartados.Soportado=" & Soportado
end if
if sqlQuery<>"" then
	sqlQuery= " where Lineas_Apartados.idApartado>-1 " &  sqlQuery
end if
%>
<%
Dim rsApartados
Dim rsApartados_numRows

sqlQuery = "SELECT Lineas_Apartados.idApartado, Lineas_Apartados.ordenconsulta, " _
	& " ISNULL(IdiomasApartados.ApartadoIdioma, Lineas_Apartados.Apartado) AS Apartado, " _
	& " ISNULL(IdiomasClaseNegocio.ClasenegocioIdioma, Aux_ClaseNegocio.Clasenegocio) AS Clasenegocio, " _
	& " Lineas_Apartados.apartadoimpuestos, Lineas_Apartados.apartadointeres,  " _
	& " Lineas_Apartados.exportarolap, Lineas_Apartados.signoOLAP, " _
	& " ISNULL(IdiomasTipoApartado.TipoApartadoIdioma, Aux_tipoApartado.TipoApartado) AS TipoApartado,  " _
	& " ISNULL(ISNULL(IdiomasTipoApartado2.TipoApartadoIdioma, Aux_tipoApartadoOLAP.TipoApartado),'-') AS TipoApartadoOLAP,  " _
	& " ISNULL(IdiomasVisualizacion.visualizacionIdioma, aux_visualizacion.visualizacion) AS visualizacion, Lineas_Apartados.ApartadoInteres " _
	& " FROM Lineas_Apartados " _
		& " INNER JOIN Aux_tipoApartado " _
			& " on Lineas_Apartados.soportado=Aux_tipoApartado.idTipoApartado " _
		& " LEFT JOIN IdiomasTipoApartado " _
			& " ON Aux_tipoApartado.idTipoApartado = IdiomasTipoApartado.idTipoApartado " _
			& " AND IdiomasTipoApartado.idIdioma=" & idIdiomaCookieCombo _ 
		& " INNER JOIN aux_visualizacion " _
			& " on Lineas_Apartados.idvisualizacion=aux_visualizacion.idvisualizacion " _
		& " LEFT JOIN IdiomasVisualizacion " _
			& " ON aux_visualizacion.idvisualizacion=IdiomasVisualizacion.idvisualizacion " _
			& " AND IdiomasVisualizacion.idIdioma=" & idIdiomaCookieCombo _ 
		& " INNER JOIN aux_claseNegocio " _
			& " on Lineas_Apartados.idClaseNegocio=Aux_ClaseNegocio.idClaseNegocio " _
		& " LEFT JOIN IdiomasClaseNegocio " _
			& " ON Aux_ClaseNegocio.idClaseNegocio = IdiomasClaseNegocio.idClaseNegocio " _
			& " AND IdiomasClaseNegocio.idIdioma=" & idIdiomaCookieCombo _
		& " LEFT JOIN IdiomasApartados " _
			& " ON Lineas_Apartados.idApartado = IdiomasApartados.idApartado " _
			& " AND IdiomasApartados.idIdioma=" & idIdiomaCookieCombo _
		& " LEFT JOIN Aux_tipoApartado Aux_tipoApartadoOLAP" _
			& " on Lineas_Apartados.idTipoApartadoOLAP=Aux_tipoApartadoOLAP.idTipoApartado " _
		& " LEFT JOIN IdiomasTipoApartado IdiomasTipoApartado2" _
			& " ON Aux_tipoApartadoOLAP.idTipoApartado = IdiomasTipoApartado2.idTipoApartado " _
			& " AND IdiomasTipoApartado2.idIdioma=" & idIdiomaCookieCombo _ 
		& sqlQuery & " ORDER BY Apartado "

AbrirRecordSet rsApartados, sqlQuery, cn_STRING

rsApartados_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
rsApartados_numRows = rsApartados_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsApartados_total
Dim rsApartados_first
Dim rsApartados_last

' set the record count
rsApartados_total = rsApartados.RecordCount

' set the number of rows displayed on this page
If (rsApartados_numRows < 0) Then
  rsApartados_numRows = rsApartados_total
Elseif (rsApartados_numRows = 0) Then
  rsApartados_numRows = 1
End If

' set the first and last displayed record
rsApartados_first = 1
rsApartados_last  = rsApartados_first + rsApartados_numRows - 1

' if we have the correct record count, check the other stats
If (rsApartados_total <> -1) Then
  If (rsApartados_first > rsApartados_total) Then
    rsApartados_first = rsApartados_total
  End If
  If (rsApartados_last > rsApartados_total) Then
    rsApartados_last = rsApartados_total
  End If
  If (rsApartados_numRows > rsApartados_total) Then
    rsApartados_numRows = rsApartados_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsApartados_total = -1) Then

  ' count the total records by iterating through the recordset
  rsApartados_total=0
  While (Not rsApartados.EOF)
    rsApartados_total = rsApartados_total + 1
    rsApartados.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsApartados.CursorType > 0) Then
    rsApartados.MoveFirst
  Else
    rsApartados.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsApartados_numRows < 0 Or rsApartados_numRows > rsApartados_total) Then
    rsApartados_numRows = rsApartados_total
  End If

  ' set the first and last displayed record
  rsApartados_first = 1
  rsApartados_last = rsApartados_first + rsApartados_numRows - 1
  
  If (rsApartados_first > rsApartados_total) Then
    rsApartados_first = rsApartados_total
  End If
  If (rsApartados_last > rsApartados_total) Then
    rsApartados_last = rsApartados_total
  End If

End If
%>
<%
Dim MM_paramName 
%>

<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsApartados
MM_rsCount   = rsApartados_total
MM_size      = rsApartados_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsApartados_first = MM_offset + 1
rsApartados_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsApartados_first > MM_rsCount) Then
    rsApartados_first = MM_rsCount
  End If
  If (rsApartados_last > MM_rsCount) Then
    rsApartados_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Web.css" rel="stylesheet" type="text/css">
<script language="javascript">
var idClaseNegocio="<%=idClaseNegocio%>";
</script>
<script language="javascript" src="../Combos/js/FuncionesCombos.js"></script>
<script language="javascript" src="../Combos/js/RecargarCombos.js"></script>
<script language="JavaScript">
function Nueva(){
	document.frmNavegar.action="ApartadoDatos.asp";
	document.frmNavegar.submit();
}
function Filtrar(){
	document.frmNavegar.action="Apartados.asp";
	document.frmNavegar.submit();
}
function DesFiltrar(){
	document.frmNavegar.Apartado.value="";
	document.frmNavegar.idClaseNegocio.value="";
	document.frmNavegar.Soportado.value="";
	idClaseNegocio="";
	document.frmNavegar.action="Apartados.asp";
	document.frmNavegar.submit();
}
function Borrar(Cual){
	if (confirm("<%=TraducirTexto("�Desea eliminar el apartado?")%>")){
		document.frmNavegar.action="ApartadoDelete.asp";
		document.frmNavegar.idApartado.value=Cual;
		document.frmNavegar.submit();
	}
}
function Modificar(Cual){
	document.frmNavegar.action="ApartadoUpdate.asp";
	document.frmNavegar.idApartado.value=Cual;
	document.frmNavegar.submit();
}
function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="ApartadoIdioma";
	document.frmNavegar.submit();
}
function Coeficientes(Cual){
	document.frmNavegar.action="ApartadoUsosCoeficientes.asp";
	document.frmNavegar.idApartado.value=Cual;
	document.frmNavegar.submit();
}

function Cerrar(){
	document.frmNavegar.action="../main.asp";
	document.frmNavegar.submit();
}
</script>
</head>
<body onLoad="RecargarTodosCombos()"><form name="frmNavegar" action="" method="post">
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Apartados ", "Listado de apartados" %>
<%Cabecera2 "Nuevo;Aplicar Filtros;Quitar Filtros;Cerrar"%>
<table width="100%" class="Casilla">
	<tr>
		
      <td class="TextoNegrita" colspan="4"><%= TraducirTexto("Filtros")%>
        <hr size="1" noshade class="Texto"></td>
	</tr>
	<tr> 
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Apartado")%>:</td>
      <td class="Texto" align="left" width="35%"><input type="text" name="Apartado" value="<%=Apartado%>" style="width:100%" maxlength="100"></td>
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Clase de negocio")%>:</td>
      <td class="Texto" align="left" width="35%" id="DIVidClaseNegocioD"></td>
    </tr>
	<tr>
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Tipo apartado")%>:</td>
      <td class="Texto" align="left" width="35%"><%PintarCombo "Soportado", "Aux_TipoApartado LEFT JOIN IdiomasTipoApartado ON Aux_TipoApartado.idTipoApartado=IdiomasTipoApartado.idTipoApartado AND IdiomasTipoApartado.idIdioma="&idIdiomaCookieCombo, "Aux_TipoApartado.idTipoApartado", "ISNULL(IdiomasTipoApartado.TipoApartadoIdioma, Aux_TipoApartado.TipoApartado)", Soportado, CTE_MostrarVacio%></td>
	</tr>
</table>
<br>

  <table border="0" width="100%" cellspacing="0" cellpadding="2">
    <tr > 
      <td class="Cabecera" ><%= TraducirTexto("Apartado")%></td>
      <td class="Cabecera"  nowrap><%= TraducirTexto("Clase de negocio")%></td>
      <td class="Cabecera"nowrap><%= TraducirTexto("Tipo apartado")%></td>
      <td class="Cabecera" nowrap><%= TraducirTexto("Visualizar")%></td>
      <td class="Cabecera"  align="center"><%= TraducirTexto("Orden consulta")%></td>
      <td class="Cabecera"  align="center"><%= TraducirTexto("Apartado Impuestos")%></td>
      <!--td class="Cabecera"  align="center"><%= TraducirTexto("Apartado Financieros")%></td-->
      <td class="Cabecera"><%= TraducirTexto("Tipo apartado OLAP")%></td>
      <td class="Cabecera"  align="center"><%= TraducirTexto("Exportar OLAP")%></td>
      <td class="Cabecera"  align="center"><%= TraducirTexto("Signo OLAP")%></td>
      <td class="Cabecera" colspan="4" width="1px">&nbsp;</td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsApartados.EOF)) %>
    <tr> 
      <td class="Linea"><%=(rsApartados.Fields.Item("Apartado").Value)%></td>
      <td class="Linea"><%=(rsApartados.Fields.Item("ClaseNegocio").Value)%></td>
	  <td class="Linea"><%=(rsApartados.Fields.Item("TipoApartado").Value)%></td>
	  <td class="Linea" nowrap><%=(rsApartados.Fields.Item("visualizacion").Value)%></td>
	  <td class="Linea" align="center"><%=(rsApartados.Fields.Item("OrdenCOnsulta").Value)%></td>
	  <td class="Linea" align="center"><img src="../img/Estado<%=(rsApartados.Fields.Item("ApartadoImpuestos").Value)%>.gif"></td>
	  <!--td class="Linea" align="center"><img src="../img/Estado<%=(rsApartados.Fields.Item("ApartadoInteres").Value)%>.gif"></td-->
	  <td class="Linea"><%=(rsApartados.Fields.Item("TipoApartadoOLAP").Value)%></td>
	  <td class="Linea" align="center"><img src="../img/Estado<%=(rsApartados.Fields.Item("ExportarOLAP").Value)%>.gif"></td>
	  <td class="Linea" align="center"><%=(rsApartados.Fields.Item("SignoOLAP").Value)%></td>
      <td class="Linea">&nbsp;<!--img alt="Coeficientes usos" style="cursor:hand" onClick="Coeficientes('<%'=(rsApartados.Fields.Item("idApartado").Value)%>')" src="../img/Porcentaje.jpg"--></td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsApartados.Fields.Item("idApartado").Value)%>')" src="../img/Idiomas.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsApartados.Fields.Item("idApartado").Value)%>')" src="../img/Editar.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Eliminar")%>" style="cursor:hand" onClick="Borrar('<%=(rsApartados.Fields.Item("idApartado").Value)%>')" src="../img/Borrar.gif"></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsApartados.MoveNext()
Wend
%>
  </table>



<span class="Texto2"><%=TraducirTexto("Registros")%>&nbsp; <%=(rsApartados_first)%>&nbsp; <%=TraducirTexto("a")%> &nbsp;<%=(rsApartados_last)%>&nbsp; <%=TraducirTexto("de")%> &nbsp;<%=(rsApartados_total)%></span> 
<table border="0" width="50%" align="center">
  <tr> 
    <td width="23%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_moveFirst%>"><img src="../img/First.gif" border=0></a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="31%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_movePrev%>"><img src="../img/Previous.gif" border=0></a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveNext%>"><img src="../img/Next.gif" border=0></a> 
      <% End If ' end Not MM_atTotal %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveLast%>"><img src="../img/Last.gif" border=0></a> 
      <% End If ' end Not MM_atTotal %> </td>
  </tr>
</table>
	<input type="hidden" name="Index" value="<%= MM_index %>">
	<input type="hidden" name="idApartado" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">	
</form>
</body>
</html>
<%
rsApartados.Close()
Set rsApartados = Nothing
%>
