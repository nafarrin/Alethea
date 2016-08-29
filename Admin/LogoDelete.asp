<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idLogo, Mindex
idLogo=Trim(Request.Form("idLogo"))
Mindex=Trim(Request.Form("index"))

dim cmdDelete
Set cmdDelete = Server.CreateObject("ADODB.Command")
cmdDelete.ActiveConnection = cn_STRING
cmdDelete.CommandText = "Delete from Logos where idLogo=" & idLogo
cmdDelete.Execute
cmdDelete.ActiveConnection.Close

set cmdDelete=nothing

response.Redirect("Logos.asp?index=" & Mindex)
%>