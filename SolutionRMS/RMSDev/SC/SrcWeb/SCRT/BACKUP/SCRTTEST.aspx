<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTTEST.aspx.vb" Inherits="SC.SCRTTest" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SCFUTest</title> 
		<!--<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

'버튼클릭 이벤트      

Sub imgPrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	
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
		Params = "SJCC:구매"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
	
	gShowReportWindow ModuleDir, ReportName, Params
	''gShowReportWindow "SC","SCMENU.rpt","SJCC:영업"

	gFlowWait meWAIT_OFF
End Sub
//-->
		</script>
	</HEAD>
	<body>
		<form id="frmThis">
			<P><FONT face="굴림"></FONT>&nbsp;</P>
			<P><FONT face="굴림">*** 크리스탈 레포트 샘플 ***</FONT></P>
			<P>
				<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px" cellSpacing="1" cellPadding="1" width="75%" border="0">
					<TR>
						<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF" border="0" name="imgWaiting">
						</TD>
					</TR>
				</TABLE>
				<TABLE id="Table1" style="WIDTH: 619px; HEIGHT: 143px" cellSpacing="1" cellPadding="1" width="619" border="1">
					<TR class="EVENROW">
						<TD style="WIDTH: 370px"><FONT face="굴림">ModuleDir </FONT><INPUT id="txtModuleDir" type="text" value="SC"></TD>
						<TD><FONT face="굴림">설명: 레포트가 실제 위치한 물리적인
								<BR>
								모듈 디렉토리 (예: SC, CO, PO, AP&nbsp;등...)</FONT></TD>
					</TR>
					<TR class="ODDROW">
						<TD style="WIDTH: 370px"><FONT face="굴림">ReportName</FONT> <INPUT id="txtReportName" type="text" value="SCMENU.rpt"></TD>
						<TD><FONT face="굴림">레포트의 이름(예: SCMENU.rpt )</FONT></TD>
					</TR>
					<TR class="EVENROW">
						<TD style="WIDTH: 370px; HEIGHT: 16.571pt"><FONT face="굴림">Params</FONT> <INPUT id="txtParams" type="text" value="SJCC:영업"></TD>
						<TD style="HEIGHT: 16.571pt"><FONT face="굴림">파라미터 값(예: SJCC:영업)</FONT></TD>
					</TR>
					<TR class="ODDROW">
						<TD style="WIDTH: 370px">
							<P align="left"><FONT face="굴림" color="#ff3366"><STRONG>%% 주의: 파라미터를 넘길 경우 무조건 String 
										Type으로 넘기고 만약 숫자형이나 날짜형의 경우 String Type 데이터를 가지고 TO_NUMBER, TO_DATE 로 변환 하여 레포트 
										사용 %%&nbsp;&nbsp; --&gt; </STRONG></FONT><STRONG><FONT face="굴림" color="#ff3366">
										크리스탈 레포트교육.doc 참조</FONT></STRONG></P>
						</TD>
						<TD><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gif" width="54" border="0" name="imgPrint"><FONT face="굴림">&nbsp;</FONT></TD>
					</TR>
				</TABLE>
			</P>
			<P><FONT face="굴림"></FONT>&nbsp;</P>
			<P><FONT face="굴림"></FONT>&nbsp;</P>
		</form>
	</body>
</HTML>
