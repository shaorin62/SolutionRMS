<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRPTMAIN.aspx.vb" Inherits="SC.SC_RPT_MAIN" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>CrystalReport Print</TITLE>
		<!--
		'****************************************************************************************
		'시스템구분 : SFAR/표준샘플/크리스탈리포트 공통웹페이지
		'실행  환경 : ASP.NET, VB.NET, COM+ 
		'프로그램명 : SC_RPT_MAIN.aspx
		'기      능 : RPT 파일을 출력한다.
		'파라  메터 : ????/SC_RPT_MAIN.aspx?rpt=mm&Param=Y&Opt=A
		'특이  사항 : 공통웹페이지(처리내역은 CodeBehind에서  처리)
		'----------------------------------------------------------------------------------------
		'HISTORY    :1) 2003/10/13 By esShin
		'****************************************************************************************
		-->
		<META content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<META content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<META content="VBScript" name="vs_defaultClientScript">
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
		 Sub Window_OnLoad

		  	Dim vntInParam, intNo, path
			vntInParam = window.dialogArguments

			intNo = ubound(vntInParam)
			'Modal 또는 Modeless 창을 띠워서 할 때
			path = "./SCRPTMAIN.asp?" & "DSN=" & frmThis.txtDBParams.value & "&ModuleDir=" & vntInParam(0) & "&ReportName=" & vntInParam(1) & "&Params=" & vntInParam(2) & "&Opt=" & vntInParam(3)
			
			'Open Window를 사용할 때(한글 문제 해결 못함)
			'path = "./SCRPTMAIN.asp?" & "DSN=" & frmThis.txtDBParams.value & "&ModuleDir=" & frmThis.txtModuleDir.value & "&ReportName=" & frmThis.txtReportName.value & "&Params=" & frmThis.txtParams.value & "&Opt=" & frmThis.txtOpt.value
	
			Dim MyRndReportNum
			Randomize								  ' 난수 발생기를 초기화합니다.
			MyRndReportNum = Int((10000 * Rnd) + 1)   ' 1에서 10000까지 무작위 값을 발생합니다.
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
