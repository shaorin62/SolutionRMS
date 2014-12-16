<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOINITPOP.aspx.vb" Inherits="SC.SCCOINITPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공지사항</title> 
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOMPPPOP.aspx
'기      능 : MPP 팝업
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/07 By KTY
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjSCCOGET 
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : mstrFields = vntInParam(i)			'조회추가필드
				case 1 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				case 2 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 3 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with	
end sub

Sub EndPage()
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												공지사항
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/back_p.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="2"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD align="center" class="TITLE">
									<FONT face="돋음" size="2">제목 </FONT>
								</TD>
								<td class="TITLE">
									&nbsp; !중요 [RMS] 접속관련 공지 입니다.
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT"></TD>
							</TR>
							<tr>
								<td colspan="2" class="TITLE" align="left">
									<P>
										<br>
										안녕하세요 RMS 담당자 입니다.
										<BR>
										현재 SK_P 와 사업 시스템 연동과 관련하여 RMS 접속
										<BR>
										아이디의 경우 기존 M&amp;C 에서 사용하시던 ID 와 PW 가
										<BR>
										임시 중지 되어있습니다.
										<BR>
										따라서 이전에 계정으로 로그인 하셔야 하는경우에는
										<BR>
										RMS 담당자에게 문의하여 주십시요..
										<BR>
										<br>
										신규 SK_P 쪽 계정으로 로그인 하실경우에 패스워드가
										<BR>
										신규 사번과 동일하게 입력하시면 접속이 가능합니다.
										<BR>
										감사합니다. <font color="red">[예 -&gt; 신규사번: 123456 패스워드:123456]</font>
										<br>
										<br>
										<font color="red">RMS - 문의사항, 02-6390-3981</font>
									</P>
								</td>
							</tr>
							<tr>
								<td align="right" colspan="2">
									<TABLE id="tblButton" style="WIDTH: 52px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="화면을 닫습니다." src="../../../images/imgClose.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
						</TABLE>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
