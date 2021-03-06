<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMTRANSINTERNETPOP.aspx.vb" Inherits="MD.MDCMTRANSINTERNETPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>거래명세서 조회</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/공통/공통코드 팝업
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPOP1.aspx
'기      능 : JOBNO 조회를 위한 팝업
'파라  메터 : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , 조회추가필드, 현재 사용중인 것만 조회할지 여부,
'			  코드 사용시점, 코드Like할지 여부
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
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

Dim mobjMDCMINTERNETCOMMI
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mtranscommiflag, mtransTblflag
Dim mstrGUBUN

'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
end sub

Sub txtJOBNO_onkeydown
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
	call Window_OnUnload()
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
	set mobjMDCMINTERNETCOMMI	= gCreateRemoteObject("cMDIT.ccMDITINTERNETCOMMI")
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
				case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
				case 1 : .txtTRANSNO.value = vntInParam(i)
				case 2 : .txtCLIENTCODE.value  = vntInParam(i)	'조회추가필드
				case 3 : .txtCLIENTNAME.value  = vntInParam(i)	'조회추가필드
				case 4 : mstrGUBUN  = vntInParam(i)		'조회추가필드
				case 5 : mtransTblflag  = vntInParam(i)			'조회추가필드
			end select
		next
		'msgbox mtranscommiflag
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
        With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 5, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME|GBN"
		
		txtLB.innerHTML ="청구지코드&nbsp;"
		txtLBNAME.innerHTML ="청구지명&nbsp;"
		mobjSCGLSpr.SetHeader .sprSht, "년월|번호|CLIENTCODE|매체명|계산서생성"
		
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "10 | 9 | 0| 14|9"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "TRANSYEARMON"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "TRANSNO"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "CLIENTNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "GBN"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", true
		
		mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
		mobjSCGLSpr.SetCellAlign2 .sprSht, "TRANSYEARMON|TRANSNO|CLIENTCODE|GBN",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME",-1,-1,0,2,false

        End With
	end with	
	'자료조회	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDCMINTERNETCOMMI = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
	Dim EndFlag
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .chkFlag.checked = True Then
			EndFlag = 1
		Else
			EndFlag = 0
		End If
		vntData = mobjMDCMINTERNETCOMMI.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME.value, mstrGUBUN, EndFlag)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
end sub

-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
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
											<td class="TITLE" id="objTitle">
												거래명세서&nbsp;조회
											</td>
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
										width="168" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgConfirmOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirm.gif'"
													height="20" alt="자료를 선택합니다." src="../../../images/imgConfirm.gif" width="54" border="0"
													name="imgConfirm"></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="화면을 닫습니다." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
										<TABLE class="KEY" id="tblKey" style="WIDTH: 392px" cellSpacing="0" cellPadding="0" width="392"
											border="0">
											<TBODY>
												<TR height="24">
													<TD align="right" WIDTH="97" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON,txtTRANSNO)">
														거래명세서번호&nbsp;</TD>
													<TD align="left"><INPUT class="INPUT" id="txtTRANSYEARMON" type="text" size="6" name="txtTRANSYEARMON" style="WIDTH: 73px; HEIGHT: 22px"></TD>
													<td>&nbsp;- <INPUT class="INPUT" id="txtTRANSNO" style="WIDTH: 48px; HEIGHT: 22px" type="text" size="2"
															name="txtTRANSNO"></td>
													<td vAlign="middle" style="WIDTH: 104px" align="right">계산서완료구분 &nbsp;</td>
													<td><INPUT type="checkbox" id="chkFlag" name="chkFlag"></td>
								</TD>
							</TR>
							<TR>
								<TD id="txtLB" align="right" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,'')">
									&nbsp;</TD>
								<TD style="WIDTH: 72px" align="left"><INPUT class="INPUT" id="Text1" style="WIDTH: 73px; HEIGHT: 22px" type="text" size="6"
										name="txtCLIENTCODE">
								</TD>
								<TD id="txtLBNAME" WIDTH="80" align="right" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)">
									&nbsp;</TD>
								<TD align="left" colspan="2"><INPUT class="INPUT" id="txtCLIENTNAME" type="text" size="18" name="txtCLIENTNAME" style="WIDTH: 144px; HEIGHT: 22px">
								</TD>
							</TR>
						</TABLE>
					</FONT>
				</TD>
			</TR>
			<TR>
				<TD class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
			</TR>
			<TR>
				<TD align="center"><FONT face="굴림">
						<OBJECT id="sprSht" style="WIDTH: 392px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="10372">
							<PARAM NAME="_ExtentY" VALUE="7250">
							<PARAM NAME="_StockProps" VALUE="64">
							<PARAM NAME="Enabled" VALUE="-1">
							<PARAM NAME="AllowCellOverflow" VALUE="0">
							<PARAM NAME="AllowDragDrop" VALUE="0">
							<PARAM NAME="AllowMultiBlocks" VALUE="0">
							<PARAM NAME="AllowUserFormulas" VALUE="0">
							<PARAM NAME="ArrowsExitEditMode" VALUE="0">
							<PARAM NAME="AutoCalc" VALUE="-1">
							<PARAM NAME="AutoClipboard" VALUE="-1">
							<PARAM NAME="AutoSize" VALUE="0">
							<PARAM NAME="BackColorStyle" VALUE="0">
							<PARAM NAME="BorderStyle" VALUE="1">
							<PARAM NAME="ButtonDrawMode" VALUE="0">
							<PARAM NAME="ColHeaderDisplay" VALUE="2">
							<PARAM NAME="ColsFrozen" VALUE="0">
							<PARAM NAME="DAutoCellTypes" VALUE="1">
							<PARAM NAME="DAutoFill" VALUE="1">
							<PARAM NAME="DAutoHeadings" VALUE="1">
							<PARAM NAME="DAutoSave" VALUE="1">
							<PARAM NAME="DAutoSizeCols" VALUE="2">
							<PARAM NAME="DInformActiveRowChange" VALUE="1">
							<PARAM NAME="DisplayColHeaders" VALUE="1">
							<PARAM NAME="DisplayRowHeaders" VALUE="1">
							<PARAM NAME="EditEnterAction" VALUE="0">
							<PARAM NAME="EditModePermanent" VALUE="0">
							<PARAM NAME="EditModeReplace" VALUE="0">
							<PARAM NAME="FormulaSync" VALUE="-1">
							<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
							<PARAM NAME="GridColor" VALUE="12632256">
							<PARAM NAME="GridShowHoriz" VALUE="1">
							<PARAM NAME="GridShowVert" VALUE="1">
							<PARAM NAME="GridSolid" VALUE="1">
							<PARAM NAME="MaxCols" VALUE="500">
							<PARAM NAME="MaxRows" VALUE="500">
							<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
							<PARAM NAME="NoBeep" VALUE="0">
							<PARAM NAME="NoBorder" VALUE="0">
							<PARAM NAME="OperationMode" VALUE="0">
							<PARAM NAME="Position" VALUE="0">
							<PARAM NAME="ProcessTab" VALUE="0">
							<PARAM NAME="Protect" VALUE="-1">
							<PARAM NAME="ReDraw" VALUE="1">
							<PARAM NAME="RestrictCols" VALUE="0">
							<PARAM NAME="RestrictRows" VALUE="0">
							<PARAM NAME="RetainSelBlock" VALUE="-1">
							<PARAM NAME="RowHeaderDisplay" VALUE="1">
							<PARAM NAME="RowsFrozen" VALUE="0">
							<PARAM NAME="ScrollBarExtMode" VALUE="0">
							<PARAM NAME="ScrollBarMaxAlign" VALUE="-1">
							<PARAM NAME="ScrollBars" VALUE="3">
							<PARAM NAME="ScrollBarShowMax" VALUE="-1">
							<PARAM NAME="SelectBlockOptions" VALUE="15">
							<PARAM NAME="ShadowColor" VALUE="-2147483633">
							<PARAM NAME="ShadowDark" VALUE="-2147483632">
							<PARAM NAME="ShadowText" VALUE="-2147483630">
							<PARAM NAME="StartingColNumber" VALUE="1">
							<PARAM NAME="StartingRowNumber" VALUE="1">
							<PARAM NAME="UnitType" VALUE="1">
							<PARAM NAME="UserResize" VALUE="3">
							<PARAM NAME="VirtualMaxRows" VALUE="-1">
							<PARAM NAME="VirtualMode" VALUE="0">
							<PARAM NAME="VirtualOverlap" VALUE="0">
							<PARAM NAME="VirtualRows" VALUE="0">
							<PARAM NAME="VirtualScrollBuffer" VALUE="0">
							<PARAM NAME="VisibleCols" VALUE="0">
							<PARAM NAME="VisibleRows" VALUE="0">
							<PARAM NAME="VScrollSpecial" VALUE="0">
							<PARAM NAME="VScrollSpecialType" VALUE="0">
							<PARAM NAME="Appearance" VALUE="0">
							<PARAM NAME="TextTip" VALUE="0">
							<PARAM NAME="TextTipDelay" VALUE="500">
							<PARAM NAME="ScrollBarTrack" VALUE="0">
							<PARAM NAME="ClipboardOptions" VALUE="15">
							<PARAM NAME="CellNoteIndicator" VALUE="0">
							<PARAM NAME="ShowScrollTips" VALUE="0">
							<PARAM NAME="DataMember" VALUE="">
							<PARAM NAME="OLEDropMode" VALUE="0">
						</OBJECT>
					</FONT>
				</TD>
			</TR>
			<TR>
				<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
			</TR>
		</TABLE>
		<FONT face="굴림"></FONT></TD></FORM></TR></TBODY></TABLE>
	</body>
</HTML>
