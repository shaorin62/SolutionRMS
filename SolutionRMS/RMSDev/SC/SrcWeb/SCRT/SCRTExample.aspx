<%@ Page CodeBehind="SCRTExample.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="SC.SCRTExample" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>보고서 조회 예제</TITLE> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/SC/보고서 조회 예제(전표기준)(SCRTExample)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCRTExample.aspx
'기      능 : 보고서 조회
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007/10/25 By Kim Jung Hoon
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM Class ID -->
		<!-- #INCLUDE VIRTUAL=../../../Etc/SCUIClass.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mRowCnt

'=============================
' 이벤트 프로시져 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgPrint_onclick	'출력버튼 클릭시
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	
	gFlowWait meWAIT_ON
		with frmThis
		          			  
			ModuleDir = "SC"

			ReportName = "SCRTEXAMPLE.rpt"
			
			Params = .txtTYY_MM.value & ":" & .txtFYY_MM.value
            Opt = "A"
		end with
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	gFlowWait meWAIT_OFF
End Sub
'-----------------------------------
' 기타 change  이벤트
'-----------------------------------
Sub txtFYY_MM_onchange
	gSetChange
End Sub

Sub txtTYY_MM_onchange
	gSetChange
End Sub

'=============================
' UI업무 프로시져 
'=============================
'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage
 '서버업무객체생성
 '권한설정/공통파 메터/화면조정 등의 기본 작업을 수행
	gInitPageSetting mobjSCGLCtl,"SC" 

	InitPageData
End Sub

Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis

	End With
End Sub

Sub EndPage
    gEndPage    
End Sub 

-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gIf)">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="420" border="0" style="WIDTH: 420px">
				<TR>
					<TD style="WIDTH: 538px"><FONT face="굴림"></FONT>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIf"
							border="0">
							<TR>
								<td style="WIDTH: 293px" align="left" width="293" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../images/TitleIcon.gIf" width="49"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</TR>
										<tr>
											<td class="TITLE"><FONT face="굴림">보고서 조회 예제</FONT></td>
										</tr>
									</table>
								<TD vAlign="middle" align="center" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 180px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIf"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 100px; HEIGHT: 24px" cellSpacing="0" cellPadding="0"
										width="204" border="0">
										<TR>
											<TD></TD>
											<TD width="3"><FONT face="굴림"><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gIf'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gIf'" height="20" alt="자료를 인쇄합니다."
														src="../../../images/imgPrint.gIf" width="54" border="0" name="imgPrint"></FONT></TD>
											<TD></TD>
											<TD style="WIDTH: 161px; HEIGHT: 24px"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIf'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIf'" height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIf"
													width="54" border="0" name="imgClose"></TD>
											<TD></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblform1" style="WIDTH: 419px" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 790px" align="center"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 790px; HEIGHT: 8px" vAlign="middle">
									<TABLE id="Table1" style="WIDTH: 418px" cellSpacing="1" cellPadding="0" width="418" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 20%"><FONT face="굴림"> JOBCUST</FONT></TD>
											<TD class="DATA" width="80%"><FONT face="굴림"></FONT><FONT face="굴림">&nbsp;<INPUT class="INPUT" id="txtFYY_MM" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="7">&nbsp;&nbsp;~&nbsp;<INPUT class="INPUT" id="txtTYY_MM" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="7"></FONT>&nbsp;
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				<TR>
					<TD class="BOTTOMSPLIT" style="WIDTH: 790px; HEIGHT: 1px" width="790"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
		</form>
		</TD></TR></TABLE></SCRIPT>
	</body>
</HTML>
