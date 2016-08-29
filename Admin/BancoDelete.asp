<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit  %>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idBanco, Mindex
idBanco=Trim(Request.Form("idBanco"))
Mindex=Trim(Request.Form("index"))

dim cmdDelete
Set cmdDelete = Server.CreateObject("ADODB.Command")
cmdDelete.ActiveConnection = cn_STRING
cmdDelete.CommandText = "Delete from dbo.Lineas_Banco where idBanco=" & idBanco
cmdDelete.Execute
cmdDelete.ActiveConnection.Close

set cmdDelete=nothing

response.Redirect("Bancos.asp?index=" & Mindex)
%>