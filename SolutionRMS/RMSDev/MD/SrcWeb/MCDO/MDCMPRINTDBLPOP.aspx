<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTDBLPOP.aspx.vb" Inherits="MD.MDCMPRINTDBLPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>광고주/매체 선택 조회</title> 
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMPRINTDBLPOP.aspx
'기      능 : 광고주/매체 선택 조회 팝업
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/04 By Kim Tae Yub
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
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjMDSRPRINTMULTILIST
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrCheck, mstrCheckMED
Dim mstrFLAG
mstrCheck = True
mstrCheckMED = True

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

Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtMEDNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick ()
	Dim i
	Dim strCLIENT_Return
	Dim strMED_Return
	Dim strReturn
	Dim strCNT
	Dim strFLAG
	DiM strFLAG2
	With frmThis
		strReturn = ""
		strCNT = 0
		strFLAG = TRUE	'처음들어갈때 이거 없으면 "|" 이게 맨앞으로 오게되므로..
		for i=1 to .sprSht_CLIENT.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CHK",i) = 1 THEN
				IF strFLAG THEN
					strCLIENT_Return = mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",i)
					strFLAG = FALSE
				ELSE
					strCLIENT_Return = strCLIENT_Return & "|" &mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",i)
				END IF
			END IF
		Next
		
		
		strFLAG2 = TRUE	'처음들어갈때 이거 없으면 "|" 이게 맨앞으로 오게되므로..
		for i=1 to .sprSht_MED.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_MED,"CHK",i) = 1 THEN
				IF strFLAG2 THEN
					strMED_Return = mobjSCGLSpr.GetTextBinding(.sprSht_MED,"MEDCODE",i)
					strFLAG2 = FALSE
				ELSE
					strMED_Return = strMED_Return & "|" &mobjSCGLSpr.GetTextBinding(.sprSht_MED,"MEDCODE",i)
				END IF
			END IF
		Next
		
		IF strCLIENT_Return = "" THEN
			gErrorMsgBox "선택한 광고주가 없습니다.","조회안내!"
			exit sub
		END IF
		
		IF strMED_Return = "" THEN
			gErrorMsgBox "선택한 매체가 없습니다.","조회안내!"
			exit sub
		END IF
		
		strReturn = .txtYEAR.value & "♥" & strCLIENT_Return & "♥" & strMED_Return
		window.returnvalue = strReturn
		call Window_OnUnload()
	END WITH
end sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

sub sprSht_CLIENT_DblClick (Col,Row)
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CLIENT, ""
		END IF
	End With
end sub

sub sprSht_MED_DblClick (Col,Row)
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_MED, ""
		END IF
	End With
end sub

'시트 클릭 이벤트
Sub sprSht_CLIENT_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_CLIENT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if
	end with
End Sub

Sub sprSht_MED_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_MED, 1, 1, , , "", , , , , mstrCheckMED
			if mstrCheckMED = True then 
				mstrCheckMED = False
			elseif mstrCheckMED = False then 
				mstrCheckMED = True
			end if
		end if
	end with
End Sub

'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMDSRPRINTMULTILIST	= gCreateRemoteObject("cMDSC.ccMDSCPRINTMULTILIST")
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
				case 0 : .txtYEAR.value = vntInParam(i)	
				case 1 : mstrFLAG = vntInParam(i)			'조회추가필드
				case 2 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				case 3 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 4 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
        With frmThis
			'광고주 그리드
            gSetSheetColor mobjSCGLSpr, .sprSht_CLIENT
			mobjSCGLSpr.SpreadLayout .sprSht_CLIENT, 3, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_CLIENT, "CHK | CLIENTCODE | CLIENTNAME"
			mobjSCGLSpr.SetHeader .sprSht_CLIENT, "선택|코드|광고주"
			mobjSCGLSpr.SetColWidth .sprSht_CLIENT, "-1", "4|0 | 23"
			mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "0", "20"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CLIENT, "CHK"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht_CLIENT, "CLIENTCODE"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht_CLIENT, "CLIENTNAME"
			mobjSCGLSpr.ColHidden .sprSht_CLIENT, "CLIENTCODE", true
			
			mobjSCGLSpr.SetScrollBar .sprSht_CLIENT,2,False,0,-1
			mobjSCGLSpr.SetCellAlign2 .sprSht_CLIENT, "CLIENTCODE",-1,-1,2,2,false
			
			'매체사 그리드
			gSetSheetColor mobjSCGLSpr, .sprSht_MED
			mobjSCGLSpr.SpreadLayout .sprSht_MED, 3, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_MED, "CHK | MEDCODE | MEDNAME"
			mobjSCGLSpr.SetHeader .sprSht_MED, "선택|코드|매체명"
			mobjSCGLSpr.SetColWidth .sprSht_MED, "-1", "4|0 | 23"
			mobjSCGLSpr.SetRowHeight .sprSht_MED, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_MED, "0", "20"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_MED, "CHK"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht_MED, "MEDCODE"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht_MED, "MEDNAME"
			mobjSCGLSpr.ColHidden .sprSht_MED, "MEDCODE", true
			
			mobjSCGLSpr.SetScrollBar .sprSht_MED,2,False,0,-1
			mobjSCGLSpr.SetCellAlign2 .sprSht_MED, "MEDCODE",-1,-1,2,2,false
        End With
	end with	
	'자료조회	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDSRPRINTMULTILIST = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim vntData1
   	Dim i, strCols
   	Dim lngRowCnt
   	Dim lngColCnt

	On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		lngRowCnt=clng(0)
		lngColCnt=clng(0)

		vntData = mobjMDSRPRINTMULTILIST.GetCUSTDBLLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEAR.value,.txtCUSTNAME.value, mstrFLAG)
		
		vntData1 = mobjMDSRPRINTMULTILIST.GetMED_DBLLIST(gstrConfigXml,lngRowCnt,lngColCnt,.txtYEAR.value,.txtMEDNAME.value, mstrFLAG)		

		if not gDoErrorRtn ("GetCUSTDBLLIST") then
			mobjSCGLSpr.SetClip .sprSht_CLIENT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   		end if
   		
   		if not gDoErrorRtn ("GetMED_DBLLIST") then
			mobjSCGLSpr.SetClip .sprSht_MED, vntData1, 1, 1, lngColCnt, lngRowCnt, True
   		end if
   	end with
end sub

-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="560" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 300px" align="left" width="300" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" valign=bottom>
												광고주/매체 선택 및 조회
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 300px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</tr>
						</table>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20">
										<TABLE class="SEARCHDATA" id="tblKey" style="WIDTH: 560px" cellSpacing="0" cellPadding="0" width="392"
											align="right" border="0">
											<TBODY>
												<TR>
													<TD class="SEARCHLABEL" width="60" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNO,'')">
														년도</TD>
													<TD class="SEARCHDATA" width="80"><INPUT class="INPUT" id="txtYEAR" type="text" size="7" name="txtYEAR" style="WIDTH: 79px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,'')">
														광고주명&nbsp;</TD>
													<TD class="SEARCHDATA" width="130" style="WIDTH: 147px"><INPUT class="INPUT" id="txtCUSTNAME" style="WIDTH: 144px; HEIGHT: 22px" type="text" size="18"
															name="txtCUSTNAME" tabIndex="1"></TD>
													<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,'')">
														매체명&nbsp;</TD>
													<TD class="SEARCHDATA" width="130"><INPUT class="INPUT" id="txtMEDNAME" style="WIDTH: 131px; HEIGHT: 22px" type="text" size="16"
															name="txtMEDNAME" tabIndex="1"></TD>
												</TR>
											</TBODY>
										</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"></TD>
							</TR>
							<TR>
								<TD align="center">
									<table>
										<tr>
											<td style="WIDTH: 280px">
												<OBJECT id="sprSht_CLIENT" style="WIDTH: 275px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
													>
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="7276">
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
											</td>
											<td style="WIDTH: 280px">
												<OBJECT id="sprSht_MED" style="WIDTH: 275px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="7276">
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
											</td>
										</tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"></TD>
							</TR>
						</TABLE>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
