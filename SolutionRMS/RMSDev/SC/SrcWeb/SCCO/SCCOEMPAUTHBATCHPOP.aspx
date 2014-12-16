<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOEMPAUTHBATCHPOP.aspx.vb" Inherits="SC.SCCOEMPAUTHBATCHPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>사원 조회</title> 
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
'HISTORY    :1) 2009/07/15 By KTY
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

Dim mobjSCCOROLEMST 
Dim mlngRowCnt, mlngColCnt
Dim mstrUSERID
Dim mstrUSERNAME
Dim mstrCheck

mstrCheck = True

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

'선택된 사용자의 권한 을 기준 사용자의 권한으로 일괄 적용한다.
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'시트 클릭
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
		
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	end With
End Sub

sub sprSht_DblClick (Col,Row)
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	End With
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjSCCOROLEMST = gCreateRemoteObject("cSCCO.ccSCCOROLEMST")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		
		
		for i = 0 to intNo
			select case i
				case 0 : .txtSETUSERID.value = vntInParam(i)	
				case 1 : .txtSETUSERNAME.value = vntInParam(i)
			end select
		next
	
		mobjSCGLCtl.DoEventQueue
	
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 5, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | EMPNO | EMP_NAME |CC_CODE | CC_NAME"
		mobjSCGLSpr.SetHeader .sprSht,			"선택|사원코드|사원명|부서코드|부서명"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "    4|      10|    16|       0|    25"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "EMPNO | EMP_NAME |CC_CODE | CC_NAME", -1, -1, 100
		mobjSCGLSpr.ColHidden .sprSht, "CC_CODE",true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EMPNO",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EMP_NAME|CC_NAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "EMPNO | EMP_NAME |CC_CODE | CC_NAME"
		
	
		mstrUSERID	 =  .txtSETUSERID.value
		mstrUSERNAME =  .txtSETUSERNAME.value
	'자료조회
		SelectRtn
	end with
end sub

Sub EndPage()
	set mobjSCCOROLEMST = Nothing
	gEndPage
End Sub


'****************************************************************************************
' 데이터 조회
'****************************************************************************************
sub SelectRtn ()
   	Dim vntData
   	Dim strUSERNO
   	Dim strUSERNAME
   	Dim strDEPT_CD
   	Dim strDEPT_NAME

	'On error resume next
	with frmThis
		strUSERNO = "" : strUSERNAME = "" : strDEPT_CD = "" : strDEPT_NAME = ""
		
		strUSERNO	= .txtEMPNO.value
		strUSERNAME = .txtEMPNAME.value
		strDEPT_CD = .txtDEPT_CD.value
		strDEPT_NAME = .txtDEPT_NAME.value
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjSCCOROLEMST.GetBATCHSCEMP(gstrConfigXml, mlngRowCnt, mlngColCnt, strUSERNO, strUSERNAME,strDEPT_CD,strDEPT_NAME)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			if mlngRowCnt <> 0 then
   				.sprSht.focus()
   			else
   				.sprSht.MaxRows = 0
   			end if 
   		end if
   	end with
end sub


'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn, intRtn2
   	Dim vntData

	With frmThis
	
		'자신이 자신의 사용자 권한을 일괄로 줄수 없도록 VALIDATION  처리 함.
		If DataValidation =False Then exit Sub

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | EMPNO | EMP_NAME |CC_CODE | CC_NAME")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtn2 = gYesNoMsgbox( mstrUSERNAME & "의 권한을 선택하신 사용자들에게 일괄 적용 하시겠습니까?"," 일괄 적용 확인!")
		If intRtn2 <> vbYes Then exit Sub
		
		if mstrUSERID = "" then
			gErrorMsgBox "변경을 적용할 사용자의 데이터가 입력 되지 않았습니다.팝업을 다시 생성 하세요","저장안내"
			exit sub
		end if 
		
		intRtn = mobjSCCOROLEMST.ProcessRtn_BATCH(gstrConfigXml,vntData, mstrUSERID)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
   		End If
   	end With
End Sub


'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
	Dim vntData
	Dim i
	Dim strSETUSERID, strUSERID
   	
	'On error resume Next
	With frmThis
   		
   		strSETUSERID = "" : strUSERID = ""
   		
   		strSETUSERID = .txtSETUSERID
   		for i = 1 to .sprSht.MaxRows
   			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" then
   				strUSERID = mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",i)
   				
   				IF strSETUSERID = strUSERID THEN 
   					gErrorMsgBox "선택하신 사용자중 기준이 되는 사용자가 선택되어 있습니다." & vbcrlf & " 사용자 1의 권한을 사용자1 의 권한으로 처리 하실수 없습니다.","삭제안내!"
					exit Function
   				END IF 
   			end if
   		next 
   		
   	End With
	DataValidation = True
End Function


-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="500" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td class="TITLE" id="objTitle">
												사용자 선택
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 150px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" height="1"></td>
								<TD align="right"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
										onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" alt="자료를 인쇄합니다." src="../../../images/imgSave.gIF"
										border="0" name="imgSave">
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" style="WIDTH: 392px" cellSpacing="0" cellPadding="0"
											width="392" align="right" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
													사원 코드</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtEMPNO" size="9" name="txtEMPNO" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
													사원&nbsp;명&nbsp;</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 140px; HEIGHT: 22px" size="18" name="txtEMPNAME"
														tabIndex="1"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPT_CD,txtDEPT_NAME)">
													부서 코드</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtDEPT_CD" size="9" name="txtDEPT_CD" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPT_CD,txtDEPT_NAME)">
													부서 명&nbsp;</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtDEPT_NAME" style="WIDTH: 140px; HEIGHT: 22px" size="18" name="txtDEPT_NAME"
														tabIndex="1"></TD>
											</TR>
										</TABLE>
										<INPUT style="Z-INDEX: 0; WIDTH: 8px; HEIGHT: 21px" id="txtSETUSERID" dataSrc="#xmlBind"
											type="hidden" name="txtSETUSERID"> <INPUT style="Z-INDEX: 0; WIDTH: 8px; HEIGHT: 21px" id="txtSETUSERNAME" dataSrc="#xmlBind"
											type="hidden" name="txtSETUSERNAME"> </FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT style="WIDTH: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="13202">
											<PARAM NAME="_ExtentY" VALUE="7824">
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
						<FONT face="굴림"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
