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
    ''''Crystal Report 10 ������ ����Ѵ�.
    Dim goToPath : goToPath="SCRTMAIN10.asp" & frmThis.txtURL.value  & "&DSN="& frmThis.txtDSN.value 
	'------------------------------------------------------------------
	'����â���� URL �̵�
	'------------------------------------------------------------------
	location.href= goToPath    
	'------------------------------------------------------------------
	'����â�� �ݰ� ��â�� ����   - SCRTCLEAN.asp���� �����߻����� ����.
	'------------------------------------------------------------------
    'self.opener = self
    'window.close    
	'Dim myRndNum  : Randomize : myRndNum=Int((10000 * Rnd) + 1)  ''' �����߻� �ʱ�ȭ 1���� 10000���� ������ �� ����
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
