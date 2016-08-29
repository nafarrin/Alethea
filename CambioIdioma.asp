<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>SPRINT</title>
<script language="javascript">
	function Cambiar(){
		document.frmDatos.action="Menu.asp";
		document.frmDatos.target="MenuFrameDATOS";
		document.frmDatos.submit();
		document.frmDatos.action="Top.asp";
		document.frmDatos.target="topFrame";
		document.frmDatos.submit();
		document.frmDatos.action="Main.asp";
		document.frmDatos.target="mainFrame";
		document.frmDatos.submit();
	}
</script>
</head>
<body onload="Cambiar()">
</body>
<form name="frmDatos" method="post"></form>
</html>
