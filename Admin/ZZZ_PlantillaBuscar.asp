<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
Dim Plantilla
Plantilla = Request.Form("Plantilla")
Dim idEstado
idEstado = request.form("idEstado")

Dim rsPlantilla
dim sqlQuery

Set rsPlantilla = Server.CreateObject("ADODB.Recordset")
rsPlantilla.ActiveConnection = cn_STRING
sqlQuery = "SELECT max( idPlantilla) as idPlantilla  FROM dbo.Plantillas " _
	& " WHERE Plantilla = '" & Replace(Plantilla, "'", "''") & "' " 
	
if idEstado<>"" then
	sqlQuery = sqlQuery & " and idEstado=" & idEstado
else
	sqlQuery = sqlQuery & " and idEstado is null"
end if
rsPlantilla.CursorType = 0
rsPlantilla.CursorLocation = 2
rsPlantilla.LockType = 1
rsPlantilla.Open sqlQuery

dim idPlantilla
idPlantilla=(rsPlantilla.Fields.Item("idPlantilla").Value)

rsPlantilla.Close()
Set rsPlantilla = Nothing

response.Redirect("PlantillaGestion.asp?idPlantilla=" & idplantilla & "&index=" & Trim(Request.Form("index")) & "&iddesglose=" & Trim(Request.querystring("iddesglose")))
%>
