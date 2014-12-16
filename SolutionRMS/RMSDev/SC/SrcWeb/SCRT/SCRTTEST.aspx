<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTTEST.aspx.vb" Inherits="SC.SCRTTest" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>SCFUTest</TITLE>
		<META content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<META content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<META content="VBScript" name="vs_defaultClientScript">
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema"><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActivX COM ClassID -->
		<!--#INCLUDE VIRTUAL = "../../../Etc/SCUIClass.inc"-->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Sub Window_OnLoad
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitPageSetting mobjSCGLCtl,"SC"
	frmThis.txtParams.value = gStrUsrBU &":영업"
End Sub

'버튼클릭 이벤트      

Sub imgOnePrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	Dim Opt
	
	If frmThis.txtModuleDir.value="" then
		ModuleDir = "SC"
	Else
		ModuleDir = frmThis.txtModuleDir.value
	End if
	
	If frmThis.txtModuleDir.value="" then
		ReportName = "SCMENU.rpt"
	Else
		ReportName = frmThis.txtReportName.value
	End if
	
	If frmThis.txtParams.value = "" then
		Params = gStrUsrBU & ":구매"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
 
	if frmThis.txtOpt.value <> "B" then
		Opt = "A"	
	Else
	    Opt = "B"	
	end if		  

 	gShowReportWindow ModuleDir, ReportName, Params, Opt

	gFlowWait meWAIT_OFF
End Sub

Sub imgTwoPrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	Dim Opt,i
	
	If frmThis.txtModuleDir.value="" then
		ModuleDir = "SC"
	Else
		ModuleDir = frmThis.txtModuleDir.value
	End if
	
	If frmThis.txtModuleDir.value="" then
		ReportName = "SCMENU.rpt"
	Else
		ReportName = frmThis.txtReportName.value
	End if
	
	If frmThis.txtParams.value = "" then
		Params = gStrUsrBU & ":구매"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
 
	if frmThis.txtOpt.value <> "B" then
		Opt = "A"	
	Else
	    Opt = "B"	
	end if
	 
	For i=0 To 3
		Select Case i
			Case 0 : Params = gStrUsrBU &":영업"
			Case 1 : Params = gStrUsrBU &":구매"
			Case 2 : Params = gStrUsrBU &":생산"
			Case 3 : Params = gStrUsrBU &":재무"
		End Select
				
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	Next 
	gFlowWait meWAIT_OFF
End Sub


Sub btnIFRAME_Onclick 
	'gShowIFrameReport(아이프레임네임, 모듈,레포트명,파라미터,옵션)
    'gShowiFrameReport(iFrameName, Module,ReportName,Params, Opt)                                             
	 gShowiFrameReport ifrTest, "SC", "SCMENU.rpt", "HPC:영업", "A"
End Sub

//-->
		</SCRIPT>
	</HEAD>
	<BODY class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)">
		<FORM id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="790">
				<TBODY>
					<TR>
						<TD style="WIDTH: 790px">
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gif" border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD align="left" width="49" rowSpan="2"><IMG id="imgTEST" height="28" src="../../images/TitleIcon.gif" width="49"></TD>
												<TD align="left" height="4"></TD></TR>
											<TR>
												<TD class="TITLE">크리스탈 레포트 테스트
													페이지</TD></TR></TABLE></TD></TR></TABLE><FONT face="돋움" size="2">
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								&nbsp;<BR>&nbsp;* 레포트 낱장&nbsp;인쇄 &gt;&gt;</FONT><IMG id="imgOnePrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gif" width="54" align="absMiddle" border="0" name="imgOnePrint"><FONT face="돋움" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								*
								레포트 연속 인쇄 &gt;&gt;</FONT><IMG id="ImgTwoPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gif" width="54" align="absMiddle" border="0" name="ImgTwoPrint"> <FONT face="돋움" size="2">
								(연속으로 출력 불가능!! 에러발생)</FONT> <FONT face="돋움" size="2"><BR></FONT><FONT face="돋움" size="2"><BR></FONT>
							<TABLE id="Table1" style="WIDTH: 704px; HEIGHT: 192px" cellSpacing="1" cellPadding="1" width="704" border="1">
								<TR class="EVENROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16px">
										<P align="left"><FONT face="돋움" size="2">ModuleDir</FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 7.066pt"><FONT face="돋움" size="2"></FONT>
										<P align="left"><INPUT id="txtModuleDir" style="WIDTH: 64px; HEIGHT: 22px" type="text" size="5" value="SC"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16px">
										<P align="left"><FONT face="돋움" size="2">&nbsp;설명: 레포트가 실제 위치한 물리적인 모듈 디렉토리 <BR>&nbsp;(예:
												SC, CO, PO,
												AP&nbsp;등...)</FONT></P></TD></TR>
								<TR class="ODDROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 14px">
										<P align="left"><FONT face="돋움" size="2">ReportName </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 7.381pt">
										<P align="left"><INPUT id="txtReportName" style="WIDTH: 224px; HEIGHT: 22px" type="text" size="32" value="SCMENU.rpt"></P></TD>
									<TD class="LABEL" style="HEIGHT: 14px">
										<P align="left"><FONT face="돋움" size="2">&nbsp;레포트의 이름(예: SCMENU.rpt )</FONT></P></TD></TR>
								<TR class="EVENROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16.571pt">
										<P align="left"><FONT face="돋움" size="2">Params </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 16.571pt">
										<P align="left"><INPUT id="txtParams" style="WIDTH: 224px; HEIGHT: 22px" type="text" size="32"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16.571pt">
										<P align="left"><FONT face="돋움" size="2">&nbsp;파라미터 값(예: SJCC:영업)</FONT></P></TD></TR>
								<TR>
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16.571pt">
										<P align="left"><FONT face="돋움" size="2">Option </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 16.571pt">
										<P align="left"><INPUT id="txtOpt" style="WIDTH: 80px; HEIGHT: 22px" type="text" size="8" value="A" name="Text1"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16.571pt">
										<P align="left"><FONT face="돋움" size="2">&nbsp;화면출력:A 프린터출력:B</FONT></P></TD></TR>
								<TR class="ODDROW">
									<TD class="DATA" colSpan="3"><FONT face="굴림"></FONT>
										<P align="left"><FONT face="굴림"><FONT color="#ff0000"><STRONG><BR>*</STRONG>&nbsp;</FONT><STRONG><FONT color="#ff0000">주의사항<BR>
														&nbsp;1.&nbsp;파라미터는 String Type 으로 넘어
														갑니다.&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp;숫자형이나
														날짜형의 경우 레포트 SQL문에서 TO_NUMBER, TO_DATE 로 변환하여&nbsp;사용합니다.<BR>&nbsp;2. 레포트에서
														*&nbsp;는 % 로 인식되도록&nbsp;되어 있습니다.
														LIKE % 문 조건을&nbsp;사용하실때는 * 를 넘겨주시면
														됩니다.<BR></FONT></STRONG></FONT></P></TD></TR></TABLE>
							<P></P><FONT face="굴림">
								<P></FONT><FONT face="돋움" size="2">&nbsp;&nbsp; * 크리스탈 레포트 뷰어가 설치 안될때 아래의
								파일을
								다운받으셔서 수동으로 설치하십시요.</FONT><FONT face="돋움" size="2"><BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								레포트 뷰어 다운받기 --&gt;
								Crystal Report Viewer&nbsp;<BR></FONT><FONT face="굴림"><FONT size="2"><STRONG>&nbsp; </STRONG><FONT face="돋움">
										*&nbsp; 관리자
										주의사항&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;다중창을 열
										경우에는 같은 레포트여도 반드시 Window 창의 이름을 다르게 하여야
										합니다.<BR></FONT></FONT></P></FONT></TD></TR>
					<TR><TD><INPUT type="button" value="iFrameTest 버튼" id="btnIFRAME" class="button"></TD></TR>
					<TR>
						<TD><IFRAME id="ifrtest" src="" frameborder="0" style="BORDER-RIGHT: #6699ff 1px solid; BORDER-TOP: #6699ff 1px solid; BORDER-LEFT: #6699ff 1px solid; WIDTH: 704px; BORDER-BOTTOM: #6699ff 1px solid; HEIGHT: 152px" scrolling="no"></IFRAME></TD>
					</TR>
				</TBODY></TABLE></TD></TR></TBODY></TABLE></FORM>
	</BODY>
</HTML>
