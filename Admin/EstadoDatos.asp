<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="../Connections/cnWeb.asp" -->

<%
dim Estado, idEstado, index, OperacionObligatoria, EstadoReal, EstadoCarga

idEstado=Trim(Request.Form("idEstado"))

Estado=Trim(Request("Estado"))

if Request.Form("OperacionObligatoria") then 
	OperacionObligatoria = "1"
else
	OperacionObligatoria = "0"
end if


if Request.Form("EstadoReal") then 
	EstadoReal = "1"
else
	EstadoReal = "0"
end if

if Request.Form("EstadoCarga") then 
	EstadoCarga = "1"
else
	EstadoCarga = "0"
end if

index=Trim(Request("index"))

dim HayError, msgError, Insertar

Insertar=Trim(Request.Form("Insertar"))

if Insertar<>"" then
	HayError=false
	
	if Estado="" then
		HAyError=true
		msgError=msgError & MostrarError(CTE_NULO,"Estado")
	end if
	
	if not HayError then
		
		dim cmdAux, sqlQuery
		Set cmdAux = Server.CreateObject("ADODB.Command")
		cmdAux.ActiveConnection = cn_STRING		
			
		if idEstado="" then 'insertar nuevo
			
			sqlQuery="INSERT INTO Estados (Estado, OperacionObligatoria, EstadoReal) VALUES " _
				& "('" &  replace(Estado,"'","''") & "','" & OperacionObligatoria & "','" & EstadoReal & "') " 				
		
		else 'modificar 
			sqlQuery="UPDATE Estados SET " _
			& " Estado='" & replace(Estado,"'","''") &"'," _
			& " OperacionObligatoria='" & OperacionObligatoria &"'," _
			& " EstadoReal='" & EstadoReal &"'" _
			& " WHERE idEstado=" & idEstado
			
		end if
		
		'Response.Write(sqlQuery)
		cmdAux.CommandText = sqlQuery
		cmdAux.Execute
		
		if EstadoCarga="1" then
		
			if idEstado="" then 
				dim rsEstados, sqlQueryEstados
	
				sqlQueryEstados = "SELECT MAX(idEstado) AS MaxIdEstado FROM Estados WHERE Estado='" & replace(Estado,"'","''")  & "'"
	
				AbrirRecordSet rsEstados, sqlQueryEstados, cn_STRING
				idEstado = rsEstados.Fields.Item("MaxIdEstado").value
				
				cerrarrecordset rsEstados			 
			end if
			
			dim cn
			Set cn = Server.CreateObject("ADODB.Connection")
			cn.open cn_STRING
			cn.CommandTimeout = 200
	
			cn.BeginTrans
			
			sqlQuery="Update Estados set EstadoCarga=0"
			cn.execute sqlQuery
			
			sqlQuery="Update Estados set EstadoCarga=1 where idEstado=" & idEstado
			cn.execute sqlQuery			
			
			if err=0  then
				cn.CommitTrans
			Else
				cn.RollbackTrans
			end if
			
			set cn=nothing
		end if
		
		cmdAux.ActiveConnection.Close	
		response.Redirect("Estados.asp?Insertar=1")
		
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
	window.location.href="Estados.asp?index=<%= index %>"
}
</script>
</head>
<body>
<!--#include file="../includes/Cabecera.asp" -->
<%Cabecera "../img/Maestros.gif", "Maestros: Estados ", "Datos estado" %>
<%Cabecera2 "Guardar y Cerrar;Volver al listado" %>

<form method="POST" action="" name="form1">
  <table align="center" class="Casilla">
    <tr class="CabeceraNaranja"> 
      <td colspan="2" class="Cabecera" ><img src="../img/Vacio.gif"><%= TraducirTexto("Datos estado")%></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Estado")%> *:</td>
      <td> <input type="text" name="Estado" value="<%=Estado%>" style="width:250px" maxlength="50"> 
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Operación obligatoria")%> *:</td>
      <td> <select name="OperacionObligatoria">
          <option value="0" <%if cstr(OperacionObligatoria)="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
          <option value="1" <%if cstr(OperacionObligatoria)="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
        </select> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Estado real")%> *:</td>
      <td> <select name="EstadoReal">
          <option value="0" <%if cstr(EstadoReal)="0" then response.Write(" selected ")%>><%= TraducirTexto("No")%></option>
          <option value="1" <%if cstr(EstadoReal)="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
        </select> </td>
    </tr>
	<tr valign="baseline"> 
      <td nowrap align="right"  class="TextoNegrita"><%= TraducirTexto("Estado Predeterminado")%> *:</td>
      <td> <select name="EstadoCarga">
          <%if cstr(EstadoCarga)="0" then %><option value="0" selected><%= TraducirTexto("No")%></option><%end if%>
          <option value="1" <%if cstr(EstadoCarga)="1" then response.Write(" selected ")%>><%= TraducirTexto("Sí")%></option>
        </select> </td>
    </tr>
  </table>
  <input type="hidden" name="Insertar" value="1">
  <input type="hidden" name="idEstado" value="<%=idEstado %>" size="32">
</form>
<span class="TextoNegrita">
* <%=TraducirTexto("Campo obligatorio")%><br>
<%= msgError %></span>
</body>
</html>
