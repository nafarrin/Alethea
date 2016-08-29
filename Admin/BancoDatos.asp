<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->



<%
dim Banco, idPartida, TipoInteres, Techo,  index, idBanco
Banco=Trim(Request.Form("Banco"))
idPartida=Trim(Request.Form("idPartida"))
TipoInteres=Trim(Request.Form("TipoInteres"))
Techo=Trim(Request.Form("Techo"))
idBanco=Trim(Request.Form("idBanco"))

index=Trim(Request("index"))

dim HayError, msgError

if Trim(Request.Form("Insertar"))<>"" then
	HayError=false
	if Banco="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"banco") 
	end if
	if idPartida<>"" then
		if not isnumeric(TEcho) then
			HAyError=true
			msgError=msgError & "<br>"& TraducirTexto("Si especifica una partida el techo tiene que ser numérico") & "."
		end if
		if not isnumeric(TipoInteres) then
			HAyError=true
			msgError=msgError & "<br>"& TraducirTexto("Si especifica una partida el tipo de interés tiene que ser numérico") & "."
		end if
	end if
	
	
	if not hayError then 
		dim cn, sqlQuery
		
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.open cn_STRING
		cn.CommandTimeout = 200
	
		if idpartida="" then 
			idpartida="null"
			Tipointeres="null"
			Techo="null"
		else
			TipoInteres=replace(TipoInteres,",",".")
			Techo=replace(Techo,",",".")
		end if
		
	
		if idBanco<>"" then 
			sqlQuery="UPDATE Lineas_Banco set " _
				& " Banco='" & replace(Banco, "'","''") & "', " _
				& " idPartida=" & idPartida & ", " _
				& " Techo=" & Techo & ", " _
				& " TipoInteres=" & TipoInteres & " " _
				& " where idBanco=" & idBanco
		else
			sqlquery="INSERT INTO Lineas_Banco " _
				& " (Banco, idPartida, Techo, TipoInteres) " _
				& " VALUES " _
				& " ('" & replace(Banco, "'","''") & "', " & idPartida & ", " & Techo & ", " & TipoInteres & ") "
			
		end if
	
		cn.execute sqlQuery
		
		set cn=nothing
		
		response.Redirect("Bancos.asp?index="&index)
	end if
	
	
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
	window.location.href="Bancos.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Bancos ", "Modificar banco" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="BancoDatos.asp" name="form1">
  <table align="center" class="Casilla">
  	<tr class="CabeceraNaranja">
		
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Modificar Banco")%> </td>
	</tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Banco")%> *:</td>
      <td> <input type="text" name="Banco" value="<%=Banco%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Partida Asociada")%>:</td>
      <td><%PintarCombo "idPartida" , "PartidasPresupuestarias LEFT JOIN IdiomasPartidasPresupuestarias ON PartidasPresupuestarias.idPartida=IdiomasPartidasPresupuestarias.idPartida AND IdiomasPartidasPresupuestarias.IdIdioma=" &idIdiomaCookieCombo,  "PartidasPresupuestarias.idPartida", "isnull(CodigoPartida,'') +' ' + ISNULL(IdiomasPartidasPresupuestarias.PartidaIdioma, PartidasPresupuestarias.partida)", idPartida, CTE_MostrarVacio%>
</td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Techo")%>:</td>
      <td> <input type="text" name="Techo" value="<%=Techo%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Tipo Interés")%>:</td>
      <td> <input type="text" name="TipoInteres" value="<%=TipoInteres%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
  </table>
  <input type="hidden" name="idBanco" value="<%=idBanco %>" size="32">
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="index" value="<%=index%>">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
