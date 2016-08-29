<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idTipoCadencia, TipoCadencia,  index
TipoCadencia=Trim(Request.Form("TipoCadencia"))
idTipoCadencia=Trim(Request.Form("idTipoCadencia"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if TipoCadencia="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Tipo cadencia")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_TipoCadencia SET " _
			& " TipoCadencia='" & replace(TipoCadencia, "'","''") & "' "  _
			& " where idTipoCadencia=" & idTipoCadencia
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("TipoCadencia.asp")
	
	end if
	
end if
%>
<%
Dim rsTipoCadencia__MMColParam
rsTipoCadencia__MMColParam = "1"
If (Request.Form("idTipoCadencia") <> "") Then 
  rsTipoCadencia__MMColParam = Request.Form("idTipoCadencia")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsTipoCadencia
	
	sqlQuery="SELECT * FROM Aux_TipoCadencia WHERE idTipoCadencia = "& rsTipoCadencia__MMColParam &  ""
	
	AbrirRecordSet rsTipoCadencia, sqlQuery, cn_STRING

	 TipoCadencia=(rsTipoCadencia.Fields.Item("TipoCadencia").Value)
	 
	 CerrarRecordset rsTipoCadencia
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
	window.location.href="TipoCadencia.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Tipo Cadencias ", "Modificar Tipo Cadencias" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="TipoCadenciaUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar TipoCadencia ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo de Cadencia")%> *:</td>
      <td> <input type="text" name="TipoCadencia" value="<%=TipoCadencia%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idTipoCadencia" value="<%= Trim(Request.Form("idTipoCadencia")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
