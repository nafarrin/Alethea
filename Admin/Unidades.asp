<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim rsUnidades
dim unidad,  sqlQuery
unidad=Trim(Request("unidad"))

sqlQuery= "SELECT Unidades.idUnidad, ISNULL(IdiomasUnidades.UnidadIdioma, Unidades.Unidad) AS Unidad, Unidades.Decimales, Unidades.Desglosable  FROM Unidades " _
	& " LEFT JOIN IdiomasUnidades " _
			& " ON Unidades.idUnidad = IdiomasUnidades.idUnidad " _
			& " AND IdiomasUnidades.idIdioma=" & idIdiomaCookieCombo

if unidad<>"" then
	sqlQuery= sqlQuery &" where ISNULL(IdiomasUnidades.UnidadIdioma, Unidades.Unidad) like '%" & replace(unidad,"'","''") & "%' "
end if

sqlQuery =  sqlQuery &  " order by unidad "

AbrirRecordSet rsUnidades, sqlQuery, cn_STRING
%>
<%
dim rsUnidades_numRows
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
rsUnidades_numRows = rsUnidades_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsUnidades_total
Dim rsUnidades_first
Dim rsUnidades_last

' set the record count
rsUnidades_total = rsUnidades.RecordCount

' set the number of rows displayed on this page
If (rsUnidades_numRows < 0) Then
  rsUnidades_numRows = rsUnidades_total
Elseif (rsUnidades_numRows = 0) Then
  rsUnidades_numRows = 1
End If

' set the first and last displayed record
rsUnidades_first = 1
rsUnidades_last  = rsUnidades_first + rsUnidades_numRows - 1

' if we have the correct record count, check the other stats
If (rsUnidades_total <> -1) Then
  If (rsUnidades_first > rsUnidades_total) Then
    rsUnidades_first = rsUnidades_total
  End If
  If (rsUnidades_last > rsUnidades_total) Then
    rsUnidades_last = rsUnidades_total
  End If
  If (rsUnidades_numRows > rsUnidades_total) Then
    rsUnidades_numRows = rsUnidades_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsUnidades_total = -1) Then

  ' count the total records by iterating through the recordset
  rsUnidades_total=0
  While (Not rsUnidades.EOF)
    rsUnidades_total = rsUnidades_total + 1
    rsUnidades.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsUnidades.CursorType > 0) Then
    rsUnidades.MoveFirst
  Else
    rsUnidades.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsUnidades_numRows < 0 Or rsUnidades_numRows > rsUnidades_total) Then
    rsUnidades_numRows = rsUnidades_total
  End If

  ' set the first and last displayed record
  rsUnidades_first = 1
  rsUnidades_last = rsUnidades_first + rsUnidades_numRows - 1
  
  If (rsUnidades_first > rsUnidades_total) Then
    rsUnidades_first = rsUnidades_total
  End If
  If (rsUnidades_last > rsUnidades_total) Then
    rsUnidades_last = rsUnidades_total
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

Set MM_rs    = rsUnidades
MM_rsCount   = rsUnidades_total
MM_size      = rsUnidades_numRows
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
rsUnidades_first = MM_offset + 1
rsUnidades_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsUnidades_first > MM_rsCount) Then
    rsUnidades_first = MM_rsCount
  End If
  If (rsUnidades_last > MM_rsCount) Then
    rsUnidades_last = MM_rsCount
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
	document.frmNavegar.action="unidadNuevo.asp";
	document.frmNavegar.submit();
}
function Filtrar(){
	document.frmNavegar.action="Unidades.asp";
	document.frmNavegar.submit();
}
function DesFiltrar(){
	document.frmNavegar.unidad.value="";
	document.frmNavegar.action="Unidades.asp";
	document.frmNavegar.submit();
}
function Borrar(Cual){
	if (confirm("<%= TraducirTexto("�Desea eliminar la unidad?")%>")){
		document.frmNavegar.action="unidadDelete.asp";
		document.frmNavegar.idUnidad.value=Cual;
		document.frmNavegar.submit();
	}
}
function Modificar(Cual){
	document.frmNavegar.action="unidadUpdate.asp";
	document.frmNavegar.idUnidad.value=Cual;
	document.frmNavegar.submit();
}
function Traducir(Cual){
	document.frmNavegar.action="../Idiomas/Traducciones.asp";
	document.frmNavegar.idCampoIdioma.value=Cual;
	document.frmNavegar.CampoIdioma.value="UnidadIdioma";
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
<%Cabecera "../img/Maestros.gif", "Maestros: Unidades", "Listado de unidades" %>
<%Cabecera2 "Nuevo;Aplicar Filtros;Quitar Filtros;Cerrar"%>

<table width="100%" class="Casilla">
	<tr>
		
      <td class="TextoNegrita" colspan="4"><%= TraducirTexto("Filtros")%>
        <hr size="1" noshade class="Texto"></td>
	</tr>
	<tr>
		
      <td class="TextoNegrita" align="right" width="15%"><%= TraducirTexto("Unidad")%>:</td>
		
      <td class="Texto" align="left" width="35%">
<input type="text" name="unidad" value="<%=unidad%>" style="width:100%" maxlength="100"></td>
		
      <td class="TextoNegrita" align="right" width="15%">&nbsp;</td>
		
      <td class="Texto" align="left" width="35%">&nbsp;</td>
	</tr>
</table>
<br>

  <table border="0" width="100%" cellspacing="0" cellpadding="2">
    <tr > 
      <td class="Cabecera" width="45%"><%= TraducirTexto("Unidad")%></td>
       <td class="Cabecera" width="25%" align="center"><%= TraducirTexto("Decimales")%></td>
      <td class="Cabecera" width="25%" align="center"><%= TraducirTexto("Desglosable")%></td>
     <td class="Cabecera" colspan="2" width="5%">&nbsp;</td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsUnidades.EOF)) %>
    <tr> 
      <td class="Linea"><%=(rsUnidades.Fields.Item("unidad").Value)%></td>
      <td class="Linea" align="center"><%=(rsUnidades.Fields.Item("Decimales").Value)%></td>
      <td class="Linea" align="center"><img src="../img/Estado<%=(rsUnidades.Fields.Item("Desglosable").Value)%>.gif"></td>
	  <td class="Linea"><img alt="<%= TraducirTexto("Traducir")%>" style="cursor:hand" onClick="Traducir('<%=(rsUnidades.Fields.Item("idUnidad").Value)%>')" src="../img/Idiomas.gif"></td>
     <td class="Linea"><img alt="<%= TraducirTexto("Modificar")%>" style="cursor:hand" onClick="Modificar('<%=(rsUnidades.Fields.Item("idUnidad").Value)%>')" src="../img/Editar.gif"></td>
      <td class="Linea"><img alt="<%= TraducirTexto("Eliminar")%>" style="cursor:hand" onClick="Borrar('<%=(rsUnidades.Fields.Item("idUnidad").Value)%>')" src="../img/Borrar.gif"></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsUnidades.MoveNext()
Wend
%>
  </table>



<span class="Texto2"><%=TraducirTexto("Registros")%>&nbsp; <%=(rsUnidades_first)%>&nbsp; <%=TraducirTexto("a")%> &nbsp;<%=(rsUnidades_last)%>&nbsp; <%=TraducirTexto("de")%> &nbsp;<%=(rsUnidades_total)%></span> 
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
	<input type="hidden" name="idUnidad" value="">
	<input type="hidden" name="idCampoIdioma" value="">
	<input type="hidden" name="CampoIdioma" value="">
</form>
</body>
</html>
<%
rsUnidades.Close()
Set rsUnidades = Nothing
%>
