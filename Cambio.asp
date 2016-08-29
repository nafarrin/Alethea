<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
function CambiarAncho(){
	if(parent.FrameTOTAL.cols=="0,8,*"){
		parent.FrameTOTAL.cols="185,8,*";
		document.all("ImagenCambio").src="img/Cambio2.gif"
	}
	else {
		parent.FrameTOTAL.cols="0,8,*";
		document.all("ImagenCambio").src="img/Cambio.gif"
	}
	
}

</script>
</head>
<body leftmargin="0" rightmargin="0"  topmargin="0" background="img/FondoCambio.gif" style="background-repeat:repeat-y">
<table height="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td height="48%" valign="top">
		<table border="0" cellpadding="0" cellspacing="0"  width="100%" background="img/FondoTop2.gif">
			<tr>
				<td><img src="img/FondoTop2.gif"></td>
			</tr>
			<tr>
				<td><img src="img/CambioGiro.gif"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td height="2%"><img style="cursor:hand" name="ImagenCambio" src="img/Cambio2.gif" onClick="CambiarAncho()">
</td>
</tr>
<tr>
	<td height="48%">&nbsp;</td>
</tr>
</table>
</body>
</html>
