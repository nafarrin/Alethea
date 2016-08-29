<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<!--#include file="../includes/Funciones.asp" -->

<%
dim rsTipos

dim TipoImpuesto,  sqlQuery
TipoImpuesto=Trim(Request("TipoImpuesto"))


sqlQuery="SELECT AUX_TipoLineaImpuestos.idTipoImpuesto, ISNULL(IdiomasTipoLineaImpuestos.TipoImpuestoIdioma, AUX_TipoLineaImpuestos.TipoImpuesto) AS TipoImpuesto, AUX_TipoLineaImpuestos.Borrar, " _
	& " AUX_TipoLineaImpuestos.Liquidacion, ISNULL(IdiomasTipoLineaImpuestosClase.ClaseImpuestosIdioma, AUX_TipoLineaImpuestosClase.ClaseImpuestos) AS ClaseImpuestos, ISNULL(IdiomasPeriodos.PeriodoIdioma, Periodos.Periodo) AS Periodo,  " _
	& " AUX_TipoLineaImpuestos.AdmiteDevolucion, PartidasPresupuestarias.CodigoPartida + ' ' + ISNULL(IdiomasPartidasPresupuestarias.PartidaIdioma, PartidasPresupuestarias.Partida) AS Partida " _
	& " FROM AUX_TipoLineaImpuestos  " _
	& " INNER JOIN AUX_TipoLineaImpuestosClase  " _
	& " 	ON AUX_TipoLineaImpuestos.idClaseImpuestos=AUX_TipoLineaImpuestosClase.idClaseImpuestos  " _
	& " INNER JOIN Periodos  " _
	& " 	ON AUX_TipoLineaImpuestos.idPeriodoLiquidacion=Periodos.idPeriodo " _
	& " LEFT JOIN PartidasPresupuestarias " _
	& "		ON AUX_TipoLineaImpuestos.idPartidaLiquidacion=PartidasPresupuestarias.idPartida " _
	& " LEFT JOIN IdiomasTipoLineaImpuestos " _
	& " 	ON AUX_TipoLineaImpuestos.idTipoImpuesto = IdiomasTipoLineaImpuestos.idTipoImpuesto " _
	& " 	AND IdiomasTipoLineaImpuestos.idIdioma=" & idIdiomaCookieCombo _
	& " LEFT JOIN IdiomasPeriodos " _
	& " 	ON Periodos.idPeriodo = IdiomasPeriodos.idPeriodo " _
	& " 	AND IdiomasPeriodos.idIdioma=" & idIdiomaCookieCombo _
	& " LEFT JOIN IdiomasPartidasPresupuestarias " _
	& " 	ON PartidasPresupuestarias.idPartida = IdiomasPartidasPresupuestarias.idPartida " _
	& " 	AND IdiomasPartidasPresupuestarias.idIdioma=" &  idIdiomaCookieCombo _
	& " LEFT JOIN IdiomasTipoLineaImpuestosClase " _
	& " 	ON AUX_TipoLineaImpuestos.idClaseImpuestos = IdiomasTipoLineaImpuestosClase.idClaseImpuestos " _
	& " 	AND IdiomasTipoLineaImpuestosClase.idIdioma=" & idIdiomaCookieCombo
if TipoImpuesto<>"" then
	sqlQuery= sqlQuery & " where ISNULL(IdiomasTipoLineaImpuestos.TipoImpuestoIdioma, AUX_TipoLineaImpuestos.TipoImpuesto) like '%" &replace(TipoImpuesto,"'","''") & "%' "
end if
sqlQuery= sqlQuery &  " order by TipoImpuesto "

AbrirRecordset rsTipos, sqlQuery, cn_STRING
%>
<%
dim rsTipos_numRows
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
rsTipos_numRows = rsTipos_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsTipos_total
Dim rsTipos_first
Dim rsTipos_last

' set the record count
rsTipos_total = rsTipos.RecordCount

' set the number of rows displayed on this page
If (rsTipos_numRows < 0) Then
  rsTipos_numRows = rsTipos_total
Elseif (rsTipos_numRows = 0) Then
  rsTipos_numRows = 1
End If

' set the first and last displayed record
rsTipos_first = 1
rsTipos_last  = rsTipos_first + rsTipos_numRows - 1

' if we have the correct record count, check the other stats
If (rsTipos_total <> -1) Then
  If (rsTipos_first > rsTipos_total) Then
    rsTipos_first = rsTipos_total
  End If
  If (rsTipos_last > rsTipos_total) Then
    rsTipos_last = rsTipos_total
  End If
  If (rsTipos_numRows > rsTipos_total) Then
    rsTipos_numRows = rsTipos_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsTipos_total = -1) Then

  ' count the total records by iterating through the recordset
  rsTipos_total=0
  While (Not rsTipos.EOF)
    rsTipos_total = rsTipos_total + 1
    rsTipos.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsTipos.CursorType > 0) Then
    rsTipos.MoveFirst
  Else
    rsTipos.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsTipos_numRows < 0 Or rsTipos_numRows > rsTipos_total) Then
    rsTipos_numRows = rsTipos_total
  End If

  ' set the first and last displayed record
  rsTipos_first = 1
  rsTipos_last = rsTipos_first + rsTipos_numRows - 1
  
  If (rsTipos_first > rsTipos_total) Then
    rsTipos_first = rsTipos_total
  End If
  If (rsTipos_last > rsTipos_total) Then
    rsTipos_last = rsTipos_total
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

Set MM_rs    = rsTipos
MM_rsCount   = rsTipos_total
MM_size      = rsTipos_numRows
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
rsTipos_first = MM_offset + 1
rsTipos_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsTipos_first > MM_rsCount) Then
    rsTipos_first = MM_rsCount
  End If
  If (rsTipos_last > MM_rsCount) Then
    rsTipos_last = MM_rsCount
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
<script language="JavaScript">
function Nueva(){
	document.frmNavegar.action="TipoLineaImpuestosDatos.asp";
	document.frmNavegar.submit();
}
function Filtrar(){
	document.frmNavegar.action="TipoLineaImpuestos.asp";
	document.frmNavegar.submit();
}
function DesFiltrar(){
	document.frmNavegar.TipoImpuesto.value="";
	document.frmNavegar.action="TipoLineaImpuestos.asp";
	document.frmNavegar.submit();
}
function Borrar(Cual){
	if (confirm("<%= TraducirTexto("¿Desea eliminar el tipo de impuesto?")%>")){
		document.frmNavegar.action="TipoLineaImpuestosDelete.asp";
		document.frmNavegar.idTipoImpuesto.value=Cual;
		document.frmNavegar.submit();
	}
}
function Modificar(Cual){
	document.frmNavegar.action="TipoLineaImpuestosUpdate.asp";
	document.frmNavegar.idTipoImpuesto.value=Cual;
	document.frmNavegar.submit();
}
function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="TipoImpuestoIdioma";
	document.frmNavegar.submit();
}
function Cerrar(){
	document.frmNavegar.action="../main.asp";
	document.frmNavegar.submit();
}
</script>
</head>
<body><form name="frmNavegar" action="" method="post">
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Línea Impuestos ", "Listado de Tipos de Líneas de Impuestos" %>
<%Cabecera2 "Nuevo;Aplicar Filtros;Quitar Filtros;Cerrar"%>

<table width="100%" class="Casilla">
	<tr>
		
      <td class="TextoNegrita" colspan="4"><%= TraducirTexto("Filtros")%>
        <hr size="1" noshade class="Texto"></td>
	</tr>
	<tr>
		
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Tipo")%>:</td>
		
      <td class="Texto" align="left" width="35%">
<input type="text" name="TipoImpuesto" value="<%=TipoImpuesto%>" style="width:100%" maxlength="100"></td>
		
      <td class="TextoNegrita" align="right" width="15%">&nbsp;</td>
		
      <td class="Texto" align="left" width="35%">&nbsp; </td>
	</tr>
</table>
<br>

  <table border="0" width="100%" cellspacing="0" cellpadding="2">
    <tr > 
      <td class="Cabecera" ><%= TraducirTexto("Tipo Impuesto")%></td>
      <td class="Cabecera" ><%= TraducirTexto("Clase")%></td>
      <td class="Cabecera" align="center"><%= TraducirTexto("Liquidación Automática")%></td>
      <td class="Cabecera"><%= TraducirTexto("Partida liquidación")%></td>
      <td class="Cabecera" ><%= TraducirTexto("Periodo liquidación")%></td>
      <td class="Cabecera" align="center"><%= TraducirTexto("Admite devolución")%></td>
      <td class="Cabecera" colspan="3" width="1px">&nbsp;</td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsTipos.EOF)) %>
    <tr> 
      <td class="Linea"><%=(rsTipos.Fields.Item("TipoImpuesto").Value)%></td>
      <td class="Linea"><%=(rsTipos.Fields.Item("ClaseImpuestos").Value)%></td>
	  <td class="Linea" align="center"><img src="../img/Estado<%= rsTipos.Fields.Item("Liquidacion").Value %>.gif"></td>
      <td class="Linea"><%=(rsTipos.Fields.Item("Partida").Value)%>&nbsp;</td>
      <td class="Linea"><%=(rsTipos.Fields.Item("Periodo").Value)%></td>
	  <td class="Linea" align="center"><img src="../img/Estado<%= rsTipos.Fields.Item("AdmiteDevolucion").Value %>.gif"></td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsTipos.Fields.Item("idTipoImpuesto").Value)%>')" src="../img/Idiomas.gif"></td>
     <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsTipos.Fields.Item("idTipoImpuesto").Value)%>')" src="../img/Editar.gif"></td>
      <td class="Linea"><%if cstr(rsTipos.Fields.Item("Borrar").Value)="1" then %><img alt="<%= TraducirTexto("Eliminar")%>" style="cursor:hand" onClick="Borrar('<%=(rsTipos.Fields.Item("idTipoImpuesto").Value)%>')" src="../img/Borrar.gif"><%else%>&nbsp;<%end if%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsTipos.MoveNext()
Wend
%>
  </table>



<span class="Texto2"><%=TraducirTexto("Registros")%>&nbsp; <%=(rsTipos_first)%>&nbsp; <%=TraducirTexto("a")%> &nbsp;<%=(rsTipos_last)%>&nbsp; <%=TraducirTexto("de")%> &nbsp;<%=(rsTipos_total)%></span> 
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
	<input type="hidden" name="idTipoImpuesto" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
rsTipos.Close()
Set rsTipos = Nothing
%>
