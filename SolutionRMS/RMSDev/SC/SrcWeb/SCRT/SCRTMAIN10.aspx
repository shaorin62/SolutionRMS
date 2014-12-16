<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTMAIN10.aspx.vb" Inherits="SC.SCRTMAIN10" %>
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
    ''''Crystal Report 10 버젼을 사용한다.
    Dim goToPath : goToPath="SCRTMAIN10.asp" & frmThis.txtURL.value  & "&DSN="& frmThis.txtDSN.value 
	'------------------------------------------------------------------
	'현재창에서 URL 이동
	'------------------------------------------------------------------
	location.href= goToPath    
	'------------------------------------------------------------------
	'현재창을 닫고 새창을 열기   - SCRTCLEAN.asp에서 에러발생하지 않음.
	'------------------------------------------------------------------
    'self.opener = self
    'window.close    
	'Dim myRndNum  : Randomize : myRndNum=Int((10000 * Rnd) + 1)  ''' 난수발생 초기화 1에서 10000까지 무작위 값 추출
	'gShowWindow goToPath, "window"&myRndNum, "1024", "768", ""  
End Sub
-->
		</script>
	</HEAD>
	<body leftmargin="0" topmargin="0">
		<form id="frmThis">
			<table border="0" width="100%" height="100%">
				<tr bgcolor="#2B588E">
					<td align="center" valign="center"><img src="./images/Report.jpg"></td>
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
