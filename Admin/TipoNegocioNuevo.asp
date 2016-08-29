<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim TipoNegocio,  index, idClaseNegocio
TipoNegocio=Trim(Request.Form("TipoNegocio"))
idClaseNegocio=Trim(Request.Form("idClaseNegocio"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("MM_Insert"))<>"" then
	HayError=false
	if TipoNegocio="" then
		HAyError=true
		msgError=msgError & "<br>Debe especificar un valor para tipo de negocio"
	end if
end if
%>
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = cn_STRING
  MM_editTable = "dbo.TiposNegocio"
  MM_editRedirectUrl = "TiposNegocio.asp?index=" & index
  MM_fieldsStr  = "TipoNegocio|value|idClaseNegocio|value"
  MM_columnsStr = "TipoNegocio|',none,''|idClaseNegocio|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

	if HAyError then MM_abortEdit=true

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
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
<%Cabecera "../img/Maestros.gif", "Maestros: Tipos de negocio ", "Nuevo tipo de negocio" %>
<%Cabecera2 "Guardar y cerrar;Volver al listado" %>
<br>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif">Nuevo tipo 
        de negocio </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita">Tipo de negocio *:</td>
      <td> <input type="text" name="TipoNegocio" value="<%=TipoNegocio%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
	<tr>
	   <td class="TextoNegrita" align="right" >Clase de negocio:</td>
      <td>
	  	<select name="idClaseNegocio">
	  	  <%While (NOT rsClaseNegocio.EOF)%>
	  	  <option value="<%=(rsClaseNegocio.Fields.Item("idClaseNegocio").Value)%>" <%If (Not isNull(idClaseNegocio)) Then If (CStr(rsClaseNegocio.Fields.Item("idClaseNegocio").Value) = CStr(idClaseNegocio)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsClaseNegocio.Fields.Item("Clasenegocio").Value)%></option>
	  	  <%
			  rsClaseNegocio.MoveNext()
			Wend
			%>
		</select>
	  </td>
	
	</tr>

  </table>
  <input type="hidden" name="MM_insert" value="form1">
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
