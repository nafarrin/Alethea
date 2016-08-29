<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim Campo,  index
Campo=Trim(Request.Form("Campo"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("MM_Update"))<>"" then
	HayError=false
	'if TipoVersion="" then
	'	HAyError=true
	'	msgError=msgError & "<br>Debe especificar un valor para Tipo de Versión"
	'end if
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = cn_STRING
  MM_editTable = "dbo.Aux_CamposAnalisis"
  MM_editColumn = "idCampo"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "CamposAuxiliares.asp?index=" & index
  MM_fieldsStr  = "Campo|value|idCampo|value"
  MM_columnsStr = "Campo|',none,''|idCampo|none,none,NULL"

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

	if HAyError then MM_abortEdit=true

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim rsCampo__MMColParam
rsCampo__MMColParam = "1"
If (Request.Form("idCampo") <> "") Then 
  rsCampo__MMColParam = Request.Form("idCampo")
End If
%>
<%
Dim rsCampo
Dim rsCampo_numRows

Set rsCampo = Server.CreateObject("ADODB.Recordset")
rsCampo.ActiveConnection = cn_STRING
rsCampo.Source = "SELECT * FROM Aux_CamposAnalisis WHERE idCampo = " + Replace(rsCampo__MMColParam, "'", "''") + ""
rsCampo.CursorType = 0
rsCampo.CursorLocation = 2
rsCampo.LockType = 1
rsCampo.Open()

rsCampo_numRows = 0
%>
<%
if Trim(Request.Form("MM_Update"))="" then
	 Campo=(rsCampo.Fields.Item("Campo").Value)
end if

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
	window.location.href="CamposAuxiliares.asp?index=<%= index %>"
}
</script>
</head>
<body>

<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Campos Auxiliares", "Modificar Campo Auxiliar del AV." %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Campo") %></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><% =TraducirTexto("Campo")%>:</td>
      <td> <input type="text" name="Campo" value="<%=Campo%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idCampo" value="<%= Trim(Request.Form("idCampo")) %>" size="32">
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rsCampo.Fields.Item("idCampo").Value %>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo Obligatorio")%><br>
<%= TraducirTexto(msgError) %></span>
</body>
</html>
<%
rsCampo.Close()
Set rsCampo = Nothing
%>
