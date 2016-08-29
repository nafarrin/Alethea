<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option explicit%>
<!--#include file="Connections/cnWeb.asp" -->
<%
dim idUsuario, idIdioma
idUsuario=request.cookies("idUsuario")
idIdioma=request.cookies("idIdiomaCookie")
if idUsuario="" then idUsuario=-1

dim sqlQuery
dim rsMenus

sqlQuery= "SELECT 	Menus_Grupos.idGrupo, isnull(IdiomasMenusGrupos.GrupoIdioma, Menus_Grupos.Grupo) Grupo, " _
	& " Menus.idMenu, isnull(IdiomasMenus.MenuIdioma,Menus.Menu) Menu, Menus.Pagina, Menus_Grupos.MenuEspecial, Menus_Grupos.Orden, Menus.Orden " _
	& " FROM Menus  " _
	& " INNER JOIN UsuariosMenus  " _
	& " 	ON Menus.idMenu=UsuariosMenus.idMenu  " _
	& " INNER JOIN Menus_Grupos  " _
	& " 	ON Menus.idGrupo=Menus_Grupos.idGrupo  " _
	& " LEFT OUTER JOIN IdiomasMenusGrupos  " _
	& " 	ON Menus_Grupos.idGrupo=IdiomasMenusGrupos.idGrupo  " _
	& " 	AND IdiomasMenusGrupos.idIdioma="  & idIdioma  _
	& " LEFT OUTER JOIN IdiomasMenus  " _
	& " 	ON Menus.idMenu=IdiomasMenus.idMenu  " _
	& " 	AND IdiomasMenus.idIdioma=" & idIdioma  _
	& " WHERE Menus.Activo=1  " _
	& " AND UsuariosMenus.idUsuario=" & idUsuario  
	
if BuscarValor("OLAP_Usuarios", "count(*)", "idUsuario = " & idUsuario)>0 and  BuscarValor("C_SPRINT_SPRINT", "SprintO", "1=1")=0 then
	'añadir menu olap	
	sqlQuery= sqlQuery & " UNION " _
		&  "SELECT 	Menus_Grupos.idGrupo, isnull(IdiomasMenusGrupos.GrupoIdioma, Menus_Grupos.Grupo) Grupo, " _
		& " -1, '' Menu,  ''  Pagina,  Menus_Grupos.MenuEspecial, Menus_Grupos.Orden, Menus.Orden " _
		& " FROM Menus_Grupos  " _
		& " LEFT OUTER JOIN IdiomasMenusGrupos  " _
		& " 	ON Menus_Grupos.idGrupo=IdiomasMenusGrupos.idGrupo  " _
		& " 	AND IdiomasMenusGrupos.idIdioma="  & idIdioma  _
		& " LEFT JOIN Menus " _
		& "  ON Menus_Grupos.idGrupo=Menus.idGrupo " _
		& " WHERE Menus_Grupos.MenuEspecial=1  " 
end if 
	
sqlQuery= sqlQuery & " ORDER BY Menus_Grupos.Orden, Menus_Grupos.idGrupo, Menus.Orden, Menus.idMenu" 
AbrirRecordSet rsMenus, sqlQuery, cn_STRING

%>
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/Web.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
var idSubMenuActual=0;
function Mostrar(Cual) {
	if (document.all("Menu_"+Cual).style.display==""){
		document.all("Menu_"+Cual).style.display="none";
	}
	else {
		document.all("Menu_"+Cual).style.display="";
	}
	if (idSubMenuActual!=Cual) {
		if (idSubMenuActual!=0) {
			document.all("Menu_"+idSubMenuActual).style.display="none";
		}
		idSubMenuActual=Cual;
	}

}
function Navegar(Donde){

	document.frmNavegar.action=Donde;
	document.frmNavegar.submit();
}
</script>
</head>
<body topmargin="0" leftmargin="0"><form name="frmNavegar" target="mainFrame" method="post" action=""></form>
<%
dim idGrupoActual, Cerrar
idGrupoActual=-1
%>
<table width="100%" border="0" cellpadding="0" cellspacing="1">
<%
while not rsMenus.eof
	if idGrupoActual<>(rsMenus.Fields.Item("idGrupo").Value) then
	%>
		<tr>
			<td class="Menu" id="TDMenu_<%=(rsMenus.Fields.Item("idGrupo").Value)%>" onClick="Mostrar('<%=(rsMenus.Fields.Item("idGrupo").Value)%>')"  height="20px" style="cursor:hand;" onMouseOver="this.className='MenuAlt'" onMouseOut="this.className='Menu'">&nbsp;&nbsp;<%=(rsMenus.Fields.Item("Grupo").Value)%></td>
		</tr>
		<tr>
			<td>
				<div id="Menu_<%=(rsMenus.Fields.Item("idGrupo").Value)%>" style="width:100%;display:none;">
				<table width="97%" border="0" cellspacing="0" align="right">

	<%
		idGrupoActual=(rsMenus.Fields.Item("idGrupo").Value)
	end if

	if rsMenus.Fields.Item("MenuEspecial").Value="0" then
	%>
		<tr onMouseOver="this.className='SubMenuAlt'" onMouseOut="this.className='SubMenu'">
			<td class="SubMenu" width="1%" valign="top"><img src="img/BolaSub.gif" ></td>
			<td class="SubMenu" width="99%" valign="top" onClick="Navegar('<%=(rsMenus.Fields.Item("Pagina").Value)%>')"><%=(rsMenus.Fields.Item("Menu").Value)%></td>
		</tr>
	<%
	elseif rsMenus.Fields.Item("MenuEspecial").Value="1" then
	%>
		<!--#include file="OLAP/Principal/Menu.asp" -->
	<%
	end if
	rsMenus.movenext	
	
	Cerrar=false
	if rsMenus.eof then
		Cerrar=true
	elseif idGrupoActual<>(rsMenus.Fields.Item("idGrupo").Value) then
		Cerrar=true
	end if
	
	if Cerrar then
	%>
					<tr>
						<td class="SubMenu" colspan="2">&nbsp;</td>
					</tr>
				</table>
				</div>
			</td>
		</tr>
	<%
	end if
wend
%>
</table>
</body>
</html>
<%
CerrarRecordSet rsMenus
%>
