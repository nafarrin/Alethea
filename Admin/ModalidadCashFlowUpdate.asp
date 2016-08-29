<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim idModalidadCashFlow, ModalidadCashFlow,  index
ModalidadCashFlow=Trim(Request.Form("ModalidadCashFlow"))
idModalidadCashFlow=Trim(Request.Form("idModalidadCashFlow"))

index=Trim(Request("index"))

dim sqlQuery, HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if ModalidadCashFlow="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO, "Modalidad CashFlow")
	end if
	
	if not HayError then 
		dim  cn
		
		sqlQuery="Update Aux_ModalidadCashFlow SET " _
			& " ModalidadCashFlow='" & replace(ModalidadCashFlow, "'","''") & "' "  _
			& " where idModalidadCashFlow=" & idModalidadCashFlow
	
		CrearConexion cn
		
		cn.execute sqlQuery
		
		CerrarConnection cn
		
		response.Redirect("ModalidadCashFlow.asp")
	
	end if
	
end if
%>
<%
Dim rsModalidadCashFlow__MMColParam
rsModalidadCashFlow__MMColParam = "1"
If (Request.Form("idModalidadCashFlow") <> "") Then 
  rsModalidadCashFlow__MMColParam = Request.Form("idModalidadCashFlow")
End If
%>
<%
if Trim(Request.Form("Insertar"))="" then
	dim rsModalidadCashFlow
	
	sqlQuery="SELECT * FROM Aux_ModalidadCashFlow WHERE idModalidadCashFlow = "& rsModalidadCashFlow__MMColParam &  ""
	
	AbrirRecordSet rsModalidadCashFlow, sqlQuery, cn_STRING

	 ModalidadCashFlow=(rsModalidadCashFlow.Fields.Item("ModalidadCashFlow").Value)
	 
	 CerrarRecordset rsModalidadCashFlow
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
	window.location.href="ModalidadCashFlow.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Modalidad CashFlow ", "Modificar Modalidad CashFlow" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="ModalidadCashFlowUpdate.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar ModalidadCashFlow ")%></td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Modalidad de CashFlow")%> *:</td>
      <td> <input type="text" name="ModalidadCashFlow" value="<%=ModalidadCashFlow%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idModalidadCashFlow" value="<%= Trim(Request.Form("idModalidadCashFlow")) %>" size="32">
  <input type="hidden" name="Insertar" value="1">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
