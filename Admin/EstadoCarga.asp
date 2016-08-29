<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->
<%
dim idEstado
idEstado= Trim(Request.Form("idEstado"))
dim cn, sqlQuery

Set cn = Server.CreateObject("ADODB.Connection")
cn.open cn_STRING
cn.CommandTimeout = 200

cn.BeginTrans

sqlQuery="Update Estados set EstadoCarga=0"
cn.execute sqlQuery

sqlQuery="Update Estados set EstadoCarga=1 where idEstado=" & idestado
cn.execute sqlQuery


if err=0  then
	cn.CommitTrans
Else
	cn.RollbackTrans
end if

set cn=nothing

response.Redirect("Estados.asp?index=" & Trim(Request.Form("index")))

%>