<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORDIVPOP.aspx.vb" Inherits="MD.MDCMOUTDOORDIVPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>옥외 청약 데이터 분할</title> 
		<!--
'****************************************************************************************
'시스템구분 : 대대행사 관리팝업
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMEXCUTIONPOP.aspx
'기      능 : JOBNO 조회를 위한 팝업
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 20120326 by OH SE HOON
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

Dim mobjMDOTOUTDOOR
Dim mlngRowCnt, mlngColCnt
'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgClose_onclick()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub

sub imgDelRow_onclick ()
	With frmThis
		DeleteRtn
	End With 
end sub

Sub sprSht_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW  and KeyCode <> meCR then exit sub  
	
	With frmThis
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",	.sprSht.ActiveRow, .txtYEARMON.value
		mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",		.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"TITLE",		.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"TITLE",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"TOTALAMT",	.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"TOTALAMT",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMI_RATE",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",		.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"OUT_AMT",	.sprSht.ActiveRow, 0
	
	End With
End Sub

'-----------------------------
' Spread Sheet Event
'-----------------------------	
Sub sprSht_change(ByVal Col,ByVal Row)
	Dim intAMT
	Dim intOUT_AMT
	Dim intCOMMISSION
	Dim intCOMMI_RATE
	
	with frmThis
		intAMT = 0
		intOUT_AMT = 0
		
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") OR Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht, "AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht, "OUT_AMT",Row)
			
			intCOMMISSION = intAMT - intOUT_AMT
			
			IF intAMT = 0 THEN
				intCOMMI_RATE = 0
			ELSE 
				intCOMMI_RATE = intCOMMISSION / intAMT * 100
			END IF
			
			mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION", Row, intCOMMISSION
			mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE", Row, intCOMMI_RATE
		end if
	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
End Sub	

'시트 더블클릭 
sub sprSht_DBLClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		ELSE
		
		end if
	end with
end sub

'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	Dim intNo,i,vntInParam
	Dim strATTR01
	
	set mobjMDOTOUTDOOR	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR")
	
	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		
		.txtATTR01.value = vntInParam(0)
		
		strATTR01 = split(vntInParam(0),"-")
		
		.txtYEARMON.value = strATTR01(0)
		.txtSEQ.value = strATTR01(1)
		
		gSetSheetDefaultColor()
			
        gSetSheetColor mobjSCGLSpr, .sprSht 
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | CLIENTNAME | TITLE | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MEMO | COMMI_TRANS_NO | TRU_VOCH_NO | ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|년월|번호|광고주|계약명|총계약금액|월청구액|월지급액|내수율|내수액|비고|거래명세표|매입전표|ATTR01"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   0|    18|    20|        11|      11|      10|     4|    10|  20|         0|       0|     0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME | TITLE | MEMO | COMMI_TRANS_NO | TRU_VOCH_NO"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOTALAMT | AMT | OUT_AMT | COMMISSION | ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, " YEARMON | SEQ | CLIENTNAME | COMMI_TRANS_NO | TRU_VOCH_NO "
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | COMMI_TRANS_NO | TRU_VOCH_NO", true
	
		.sprSht.focus
	End With
    
	SelectRtn
end sub

Sub EndPage()
	set mobjMDOTOUTDOOR = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		vntData = mobjMDOTOUTDOOR.SelectRtn_OUTDOORDIV(gstrConfigXml,mlngRowCnt,mlngColCnt, .txtATTR01.value)

		if not gDoErrorRtn ("SelectRtn_OUTDOORDIV") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
   		end if
   	end with
end sub

Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	
	with frmThis
   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "저장할 데이터를 입력 하십시오"
			Exit Sub
		end if
		
		'저장시 빈로우 삭제후 저장
   		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht, "AMT",intCnt) = "0" AND mobjSCGLSpr.GetTextBinding(.sprSht, "OUT_AMT",intCnt) = "0" then
			mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End If
		Next
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTNAME | TITLE | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MEMO")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtn = mobjMDOTOUTDOOR.ProcessRtn_DIV(gstrConfigXml,vntData, .txtATTR01.value)
	
		if not gDoErrorRtn ("ProcessRtn_DIV") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox intRtn & "건의 자료가 저장" & mePROC_DONE , "저장안내!"
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master 입력 데이터 Validation : 필수 입력항목 검사
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  		for intCnt = 1 to .sprSht.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) = 0  AND mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",intCnt) = 0 Then 
					gErrorMsgBox intCnt & " 번째 행의 입력내용 을 확인하십시오","입력오류"
					Exit Function
			 End if
		next
   	End with
	DataValidation = true
End Function

'---------------------
'----데이터 삭제------
'---------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '실제데이터여부 플레
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_VOCH_NO",i) <> "" Then
					gErrorMsgBox "선택하신 " & i & "행의 자료는 거래명세표/매입전표가 존재 합니다.  " & vbcrlf & "  먼저 거래명세표/매입전표를 삭제 하십시오!","삭제안내!"
					exit Sub
				elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",i) = "" THEN  
					gErrorMsgBox "선택하신 " & i & "행의 자료는 분할시 최초 분할했던 원본 데이터 입니다.  " & vbcrlf & "  원본 데이터는 삭제할 수 없습니다.!","삭제안내!"
					exit Sub
				ELSE
					lngchkCnt = lngchkCnt +1
				End If
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDOTOUTDOOR.DeleteRtn(gstrConfigXml,strYEARMON,dblSEQ)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
   		End If

		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn
			If .sprSht.MaxRows = 0 Then
				Window_OnUnload
			End If 
		End If
	End With
	err.clear	
End Sub

-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="573" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td style="WIDTH: 300px" align="left" width="300" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" vAlign="bottom">옥외 청약 데이터 분할
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 225px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
												height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0" name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="HEIGHT: 340px" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="0" cellPadding="0" width="100%" align="right"
										border="0">
										<TBODY>
											<TR>
												<TD class="SEARCHLABEL" width="60">청약번호
												</TD>
												<td class="SEARCHDATA"><INPUT class="NOINPUT" id="txtYEARMON" style="WIDTH: 80px; HEIGHT: 22px" readOnly size="8"
														name="txtYEARMON"><INPUT class="NOINPUT" id="txtSEQ" style="WIDTH: 56px; HEIGHT: 22px" readOnly size="4"
														name="txtSEQ">
												</td>
											</TR>
										</TBODY>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD align = "right"><INPUT id="txtATTR01" tabIndex="1" type="hidden" name="txtATTR01"><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
									style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0"
									name="imgAddRow"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
									onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF"
									width="54" border="0" name="imgSave"><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
									style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'" alt="한 행 삭제" src="../../../images/imgDelRow.gif"
									width="54" border="0" name="imgDelRow">
								</TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT style="WIDTH: 574px; HEIGHT: 251px" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15187">
											<PARAM NAME="_ExtentY" VALUE="6641">
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
											<PARAM NAME="EditEnterAction" VALUE="5">
											<PARAM NAME="EditModePermanent" VALUE="0">
											<PARAM NAME="EditModeReplace" VALUE="0">
											<PARAM NAME="FormulaSync" VALUE="-1">
											<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
											<PARAM NAME="GridColor" VALUE="12632256">
											<PARAM NAME="GridShowHoriz" VALUE="1">
											<PARAM NAME="GridShowVert" VALUE="1">
											<PARAM NAME="GridSolid" VALUE="1">
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
					</FORM>
				</TD>
			</TR>
		</TABLE>
	</body>
</HTML>
