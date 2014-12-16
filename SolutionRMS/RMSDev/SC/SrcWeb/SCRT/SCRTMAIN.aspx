<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTMAIN.aspx.vb" Inherits="SC.SCRTMAIN" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SCRTMAIN</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<script id="clientEventHandlersVBS" language="vbscript">
<!--
Sub window_onload            
    Dim goToPath : goToPath="SCRTMAIN.asp" & frmThis.txtURL.value  & "&DSN="& frmThis.txtDSN.value 
	'------------------------------------------------------------------
	'현재창에서 URL 이동
	'------------------------------------------------------------------
	location.href= goToPath    

End Sub
-->
		</script>
	</HEAD>
	<body leftmargin="0" topmargin="0">
		<form id="frmThis">
			<table border="0" width="100%" height="100%">
				<tr bgcolor="#2B588E">
					<td align="center" valign="middle"><img src="./images/Report.jpg"></td>
				</tr>
				<TR>
					<TD><INPUT type="hidden" id="txtDSN" runat="server" NAME="txtDSN" style="WIDTH: 20px; HEIGHT: 22px"
							size="1"> <INPUT type="hidden" id="txtURL" runat="server" name="txtURL" style="WIDTH: 16px; HEIGHT: 22px"
							size="1">
					</TD>
				</TR>
			</table>
		</form>
	</body>
</HTML>
