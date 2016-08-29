<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
response.Cookies("BBDDConexion")=Trim(Request.Form("idVersion"))
response.Cookies("BBDDName_OLAP")=Trim(Request.Form("BBDDOLAP"))
response.Cookies("DTSExportacionOLAP")=Trim(Request.Form("DTSExportacionOLAP"))
response.Redirect("_MarcoDesarrollo.html")
%>