<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<title>SPRINT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript">
window.status="Equalid Solutions, s.l.";
function Control(){
	window.open("Desconectar2.asp",'SPRINT','top=0,left=0,height=10,width=10,scrollbars=yes,resizable=yes, status=no');
	//alert("para salir correctamente de la aplicación debe pulsarse el botón desconectar");
}
</script>
</head>
<frameset id="FrameTOTAL" rows="*" cols="185,8,*" frameborder="NO" border="0" framespacing="0" onUnload="Control()">
	<frameset id="FrameMENU" rows="84,*"  frameborder="NO" border="0" framespacing="0">
		<frame src="Logo.asp" name="MenuFrameLOGO" scrolling="NO" noresize >
		<frame src="Menu.asp" name="MenuFrameDATOS" scrolling="yes" >
	</frameset>
	<frame src="Cambio.asp" name="CambioFrame" scrolling="NO" noresize  >
	<frameset rows="25,*" frameborder="NO" border="0" framespacing="0">
		<frame src="Top.asp" name="topFrame" scrolling="NO" noresize frameborder="0">
		<frame src="Main.asp" name="mainFrame">
	</frameset>
</frameset>
<noframes><body>

</body></noframes>
</html>