<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idUso, Mindex
idUso=Trim(Request.Form("idUso"))
Mindex=Trim(Request.Form("index"))

dim cmdDelete
Set cmdDelete = Server.CreateObject("ADODB.Command")
cmdDelete.ActiveConnection = cn_STRING
cmdDelete.CommandText = "Delete from dbo.Lineas_Usos where idUso=" & idUso
cmdDelete.Execute
cmdDelete.ActiveConnection.Close

set cmdDelete=nothing

response.Redirect("Usos.asp?index=" & Mindex)
%>