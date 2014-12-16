<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOVOCHDEL.aspx.vb" Inherits="SC.SCCOVOCHDEL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>전표삭제</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/공통/공통코드 팝업
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCDCC.aspx
'기      능 : CC 코드 조회를 위한 팝업
'파라  메터 : LOC CODE,OC Code,MU Code,CC Type,PU Code or Name,현재 사용중인 것만 조회할지 여부,
'			  코드 사용시점,조회추가필드,코드Like할지 여부
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/03/27 By KimKS
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjMDCMGET
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrChk
mstrChk = 0
'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	'EndPage
End Sub

sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	 SelectRtn
	gFlowWait meWAIT_OFF
end sub

Sub txtCodeName_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick ()
	if frmThis.sprSht.ActiveRow > 0 then
		sprSht_DblClick frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	else
		call Window_OnUnload()
	end if
end sub

Sub imgCancel_onclick
	gEndPage
End Sub
sub sprSht_DblClick (Col,Row)
	'선택된 로우 반환
	With frmThis
	if Row = 0 and Col >0 then
		mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
	Else
	'msgbox Col & Row
	window.returnvalue = mobjSCGLSpr.GetClip (.sprSht,1,.sprSht.ActiveRow,.sprSht.MaxCols,1,1)
	call Window_OnUnload()
	end if
	End With
end sub
'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
		
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	'set mobjMDCMGET = gCreateRemoteObject("cMDCM.ccMDCMGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		'mblnUseOnly = true: mstrUseDate="" : mstrFields = "": mblnLikeCode = true
		for i = 0 to intNo
			select case i
				case 0 : document.getElementById("TextBox1").innerText = vntInParam(i)	'CC Code or Name
				case 1 : document.getElementById("TextBox2").innerText = vntInParam(i)		'현재 사용중인 것만
			end select
		next
		
		'SpreadSheet 디자인
	end with	
	
	initpageData
	'자료조회
end sub

Sub preinitpageData

End Sub
Sub initPageData

with frmThis
	if .txtSU.value <> "" Then
	
		window.returnvalue = .txtGJAH.value & "|" & .txtBELNR.value & "|" & .txtTEXT.value & "|" & .txtSU.value
		gEndPage
		End If
	end with	
DeleteVOCH
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .chkAll.checked = True Then
		strCHK = ""
		Else
		strCHK = "All"
		End if
		
		vntData = mobjMDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			'조회해온 추가 필드를 Hidden
			for i = 3 to .sprSht.MaxCols
				strCols = strCols  & i & "|"
			next
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
end sub

-->
		</script>
		<base target="_self">
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<form id="form1" runat="server">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
				<TBODY>
					<TR>
						<TD>
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
								border="0">
								<TR>
									<td style="WIDTH: 148px" align="left" width="148" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/PopupIcon.gif" width="49"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE" id="objTitle">전표삭제</td>
											</tr>
										</table>
									</td>
									<TD vAlign="middle" align="right" height="28">
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<TABLE id="tblButton" style="WIDTH: 168px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="50" border="0">
											<TR>
												<TD width="120"></TD>
												<TD>
													<asp:ImageButton id="ImageButton1" runat="server" ImageUrl="../../../images/imgDelete.gIF"></asp:ImageButton></TD>
												<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
														height="20" alt="화면을 닫습니다." src="../../../images/imgCancel.gif" width="54" border="0"
														name="imgCancel"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="394" border="0" style="WIDTH: 394px">
								<TBODY>
									<TR>
										<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
									</TR>
									<TR>
										<TD class="KEYFRAME" style="HEIGHT: 1px" vAlign="middle" height="1"><FONT face="굴림">
												<TABLE class="KEY" id="tblKey" style="WIDTH: 392px; HEIGHT: 42px" height="42" cellSpacing="0"
													cellPadding="0" width="392" align="right" border="0">
													<TBODY>
														<TR>
															<TD class="SEARCHDATA" style="WIDTH: 581px; CURSOR: hand; HEIGHT: 12px" width="581"
																onclick="vbscript:Call gCleanField(txtCodeName,'')">&nbsp; 전표년도
																<asp:TextBox id="TextBox1" runat="server" Width="88px" ReadOnly="True"></asp:TextBox>&nbsp;&nbsp;전표번호&nbsp;&nbsp;
																<asp:TextBox id="TextBox2" runat="server" Width="169px" ReadOnly="True"></asp:TextBox></TD>
														</TR>
													</TBODY>
												</TABLE>
											</FONT>
										</TD>
									</TR>
		</form>
		</TBODY></TABLE> <FONT face="굴림"></FONT></TD></TR></TBODY></TABLE>
		<form name="frmThis">
			<div id="rfcfield" ><!--style="DISPLAY: none"-->
				<INPUT id="txtGJAH" style="WIDTH: 88px; HEIGHT: 21px" type="text" size="9" name="txtGJAH" value='<%# DataBinder.Eval(SapProxy31, "tblVoch.Gjahr") %>'>
				<INPUT id="txtBELNR" style="WIDTH: 80px; HEIGHT: 21px" type="text" size="8" name="txtBELNR" value='<%# DataBinder.Eval(SapProxy31, "tblVoch.Belnr") %>'>
				<INPUT id="txtTEXT" style="WIDTH: 96px; HEIGHT: 21px" type="text" size="10" name="txtTEXT" value='<%# DataBinder.Eval(SapProxy31, "tblVoch.Text") %>'>
				<INPUT id="txtSU" style="WIDTH: 112px; HEIGHT: 21px" type="text" size="13" name="txtSU" value='<%# DataBinder.Eval(SapProxy31, "tblVoch.Subrc") %>'>
			</div>
		</form>
		<Script language="vbscript">
Sub DeleteVOCH
	with frmThis
		
	End with
End Sub
		</Script>
	</body>
</HTML>
