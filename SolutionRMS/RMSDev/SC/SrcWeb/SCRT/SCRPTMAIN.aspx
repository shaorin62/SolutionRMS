<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRPTMAIN.aspx.vb" Inherits="SC.SC_RPT_MAIN" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>CrystalReport Print</TITLE>
		<!--
		'****************************************************************************************
		'�ý��۱��� : SFAR/ǥ�ػ���/ũ����Ż����Ʈ ������������
		'����  ȯ�� : ASP.NET, VB.NET, COM+ 
		'���α׷��� : SC_RPT_MAIN.aspx
		'��      �� : RPT ������ ����Ѵ�.
		'�Ķ�  ���� : ????/SC_RPT_MAIN.aspx?rpt=mm&Param=Y&Opt=A
		'Ư��  ���� : ������������(ó�������� CodeBehind����  ó��)
		'----------------------------------------------------------------------------------------
		'HISTORY    :1) 2003/10/13 By esShin
		'****************************************************************************************
		-->
		<META content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<META content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<META content="VBScript" name="vs_defaultClientScript">
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
		 Sub Window_OnLoad

		  	Dim vntInParam, intNo, path
			vntInParam = window.dialogArguments

			intNo = ubound(vntInParam)
			'Modal �Ǵ� Modeless â�� ����� �� ��
			path = "./SCRPTMAIN.asp?" & "DSN=" & frmThis.txtDBParams.value & "&ModuleDir=" & vntInParam(0) & "&ReportName=" & vntInParam(1) & "&Params=" & vntInParam(2) & "&Opt=" & vntInParam(3)
			
			'Open Window�� ����� ��(�ѱ� ���� �ذ� ����)
			'path = "./SCRPTMAIN.asp?" & "DSN=" & frmThis.txtDBParams.value & "&ModuleDir=" & frmThis.txtModuleDir.value & "&ReportName=" & frmThis.txtReportName.value & "&Params=" & frmThis.txtParams.value & "&Opt=" & frmThis.txtOpt.value
	
			Dim MyRndReportNum
			Randomize								  ' ���� �߻��⸦ �ʱ�ȭ�մϴ�.
			MyRndReportNum = Int((10000 * Rnd) + 1)   ' 1���� 10000���� ������ ���� �߻��մϴ�.
		    IframeReport.name = MyRndReportNum 
			IframeReport.location.href = path
		 End Sub	
		</SCRIPT>
	</HEAD>
	<BODY>
		<FORM id="frmThis" method="post" runat="server">
			<INPUT type="hidden" id="txtDBParams" runat="server" NAME="txtDBParams"><INPUT id="txtModuleDir" type="hidden" name="txtModuleDir" runat="server"><INPUT id="txtReportName" type="hidden" name="txtReportName" runat="server"><INPUT id="txtParams" type="hidden" name="txtParams" runat="server"><INPUT id="txtOpt" style="WIDTH: 104px; HEIGHT: 21px" type="hidden" size="12" name="txtOpt"
				runat="server">
			<TABLE align="center" border="0" width="100%" height="100%">
				<TR>
					<TD valign="middle" align="center">
						<IFRAME id="IframeReport" width="100%" height="100%" src="images/ReportING.jpg" frameborder="0"	style="BORDER-RIGHT: #6699ff 1px solid; BORDER-TOP: #6699ff 1px solid; BORDER-LEFT: #6699ff 1px solid; BORDER-BOTTOM: #6699ff 1px solid" scrolling=no>
						</IFRAME>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</BODY>
</HTML>
