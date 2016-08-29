<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% option explicit %>
<!--#include file="Connections/cnWebINI.asp" -->
<!--#include file="includes/ConvertirFecha.asp" -->
<!--#include file="Includes/Funciones.asp" -->
<!--#include file="Includes/Encriptacion.asp" -->
<%
'declaraciones
dim Usuario
dim NumMax, NumActual, CadenaActual, Conectar

dim sqlQuery
Dim rsUsuario

Conectar=false
Usuario=Request.Form("Usuario")

'Usuario
sqlQuery="SELECT Usuarios.idUsuario, Usuarios.Usuario, Usuarios.pwd, Usuarios.Nombre, Usuarios.Apellidos, " _
	& " Usuarios.idDepartamento, Usuarios.idDelegacion, Usuarios.idPlan," _
	& " Usuarios.NoOcupa, Usuarios.Romper, Usuarios.idIdioma, Usuarios.pwdCode, Usuarios.idTipoUsuario," _
	& " UsuariosTipos.SeguridadDepartamento, UsuariosTipos.SeguridadDelegacion, " _
	& " UsuariosTipos.SeguridadPlan, UsuariosTipos.SeguridadEstado, UsuariosTipos.SeguridadProyectos" _
	& " FROM Usuarios " _
	& " INNER JOIN UsuariosTipos " _
	& " 	ON Usuarios.idTipoUsuario = UsuariosTipos.idTipoUsuario "  _
	& " WHERE Usuario='" & replace(Usuario ,"'","''") & "'" 
AbrirRecordSet rsUsuario, sqlQuery, cn_STRING



if not rsUsuario.eof then
		dim cn
		CrearConexion cn
		
		'comprobar contraseñas
		if Encrypt(Trim(Request.Form("pwd"))) = TRIM((rsUsuario.Fields.Item("pwdCode").Value) ) then
			Conectar=true
		else
			if Trim(Request.Form("pwd")) = TRIM((rsUsuario.Fields.Item("pwd").Value) ) and Trim(Request.Form("pwd"))<>"" then
				Conectar=true
				sqlQuery="UPDATE Usuarios SET pwdCode='" & Encrypt(Trim(Request.Form("pwd")))  & "', pwd=''  WHERE Usuario ='" & replace(Usuario ,"'","''") & "'" 
				cn.Execute sqlQuery
			end if
		end if

		'si se conecta
		if Conectar then
			'borrar contraseña inicial
			if Trim(Request.Form("pwd")) <>"" then
				sqlQuery="UPDATE Usuarios SET  pwd=''  WHERE Usuario ='" & replace(Usuario ,"'","''") & "'" 
				cn.Execute sqlQuery
			end if

			'Guardar datos de conexuión
			CadenaActual= ConvertirFecha(date() & " " & time )
			
			'borrar conexiones que llevan más de 18 horas
			sqlQuery="DELETE C_SPRINT_CONEXIONES WHERE GETDATE()-CAST(Inicio AS datetime) > 0.75" 
			cn.Execute sqlQuery
			
			'comprobar el numero de conexiones
			NumMax=clng(BuscarValor("C_SPRINT_SPRINT", "Sprint", "1=1"))
			NumActual=cLng(BuscarValor("C_SPRINT_CONEXIONES", "COUNT(DISTINCT idUsuario) ", "idUsuario not in (Select idUsuario from Usuarios where NoOcupa=1) and idUsuario<>" &(rsUsuario.Fields.Item("idUsuario").Value)& " " ))
			
			'si el usuario no ocupa no hace falta comprobarlo
			if (rsUsuario.Fields.Item("NoOcupa").Value) ="1" then NUmActual=0
			
			if NumActual>=numMax then
				'hay demasiados usuarios
				response.Redirect("NumMaximo.asp")
			else
				'guardamos los valores del usuario
				response.Cookies("Usuario")=Trim(Request.Form("Usuario"))
				response.Cookies("NombreUsuario")= (rsUsuario.Fields.Item("Nombre").Value) & " " & (rsUsuario.Fields.Item("Apellidos").Value)
				response.Cookies("idUsuario")=(rsUsuario.Fields.Item("idUsuario").Value)
				response.Cookies("NoOcupaCookie")=(rsUsuario.Fields.Item("NoOcupa").Value)
				if not isnull (rsUsuario.Fields.Item("idDepartamento").Value) then
					response.Cookies("idDepartamentoCookie")=(rsUsuario.Fields.Item("idDepartamento").Value)
				else
					response.Cookies("idDepartamentoCookie")="-1"
				end if
				if not isnull (rsUsuario.Fields.Item("idPlan").Value) then
					response.Cookies("idPlanCookie")=(rsUsuario.Fields.Item("idPlan").Value)
				else
					response.Cookies("idPlanCookie")="-1"
				end if
				response.Cookies("SeguridadProyectos")=(rsUsuario.Fields.Item("SeguridadProyectos").Value)
				response.Cookies("SeguridadDepartamento")=(rsUsuario.Fields.Item("SeguridadDepartamento").Value)
				response.Cookies("SeguridadDelegacion")=(rsUsuario.Fields.Item("SeguridadDelegacion").Value)
				response.Cookies("SeguridadPlan")=(rsUsuario.Fields.Item("SeguridadPlan").Value)
				response.Cookies("SeguridadEstado")=(rsUsuario.Fields.Item("SeguridadEstado").Value)
				response.Cookies("CadenaActual")=CadenaActual			
				
				dim idEstado,  rsAux
				sqlQuery="Select idEstado from UsuariosTiposEstados where idTipoUsuario=" & (rsUsuario.Fields.Item("idTipoUsuario").Value)
	
				AbrirRecordSet rsAux, sqlQuery, cn_STRING
				while not rsAux.eof
					idEstado=idEstado & rsAux.fields("idEstado").value & ","
					rsAux.movenext
				wend
				cerrarrecordset rsAux
				
				if idEstado="" then 
					idEstado="-1"
				elseif right(idEstado,1)="," then 
					idEstado = Left( idEstado, Len( idEstado ) - 1 )
				end if

				
				response.Cookies("idEstadoCookie")=idEstado			
				
				if not isnull (rsUsuario.Fields.Item("idDelegacion").Value) then
					response.Cookies("idDelegacionCookie")=(rsUsuario.Fields.Item("idDelegacion").Value)
				else
					response.Cookies("idDelegacionCookie")="-1"
				end if
				
				if not isnull (rsUsuario.Fields.Item("idIdioma").Value) then
					response.Cookies("idIdiomaCookie")=(rsUsuario.Fields.Item("idIdioma").Value)
				else
					response.Cookies("idIdiomaCookie")="1"
				end if
				
				GuardarDiccionario
				
				sqlQuery = "Insert into C_SPRINT_CONEXIONES (idUsuario, Inicio) Values (" & (rsUsuario.Fields.Item("idUsuario").Value) & " ,'" & CadenaActual & "') "
				cn.Execute sqlQuery
				
				sqlQuery = "Insert into C_SPRINT_LOG (Usuario, Accion,FechaAccion) Values ('" & (rsUsuario.Fields.Item("Usuario").Value) & "' ,'Inicio','" &  CadenaActual &"') "
				cn.Execute sqlQuery
				CerrarRecordSet rsUsuario
				
				CerrarConexion cn
				
				'AbrirConectar=true
				response.Redirect("index.asp")
			end if
			CerrarConexion cn
		'end if
	end if
else
	
end if

CerrarRecordSet rsUsuario

response.Redirect("Default.asp?UNT=24rf&Error=13dsf987d")
%>
