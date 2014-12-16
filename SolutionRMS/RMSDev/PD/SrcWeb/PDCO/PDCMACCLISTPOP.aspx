<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMACCLISTPOP.aspx.vb" Inherits="PD.PDCMACCLISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>진행비내역 관리</title> 
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
'HISTORY    :1) 2003/05/21 By 황덕수
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

Dim mobjPDCMGET
Dim mobjPDCMEXE
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrStatus
Const meTab = 9
'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgClose_onclick()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
If mstrStatus = "END" Then
	gErrorMsgBox "확정건에대하여 저장이 불가능 합니다.","처리안내"
	Exit Sub
End If
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtJOBNO_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub



Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------
' Spread Sheet Event
'-----------------------------	
'onblour 이벤트
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub
Sub txtDIVAMT_onblur
	with frmThis
		call gFormatNumber(.txtDIVAMT,0,true)
	end with
End Sub


Sub sprSht_change(ByVal Col,ByVal Row)
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
	SUM_AMT
End Sub	

sub sprSht_DblClick (Col,Row)
	'선택된 로우 반환
	'window.returnvalue = mobjSCGLSpr.GetClip (frmThis.sprSht,1,frmThis.sprSht.ActiveRow,frmThis.sprSht.MaxCols,1,1)
	'call Window_OnUnload()
end sub
sub imgAddRow_onclick ()
If mstrStatus = "END" Then
	gErrorMsgBox "완료건에대하여 입력이 불가능 합니다.","처리안내"
	Exit Sub
End If
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub
sub imgDelRow_onclick ()
If mstrStatus = "END" Then
	gErrorMsgBox "완료건에대하여 삭제가 불가능 합니다.","처리안내"
	Exit Sub
End If
	With frmThis
		call sprSht_Keydown(meDEL_ROW, 0)
	End With 
end sub

Sub sprSht_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 8 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub

Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, .txtJOBNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"ACCDAY",.sprSht.ActiveRow, gNowDate
		
	End with
End Sub


'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	Dim intNo,i,vntInParam
	
	set mobjPDCMEXE = gCreateRemoteObject("cPDCO.ccPDCOEXE")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	with frmThis
	
		.txtJOBNO.style.visibility = "hidden"
		.txtDIVAMT.style.visibility = "hidden"
	
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		'PREESTNO,YEARMON,JOBNO,CREDAY,DIVAMT
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)	
				case 1 : mstrStatus = vntInParam(i)
				'case 2 : .txtJOBNO.value = vntInParam(i)
				'case 3 : .txtCREDAY.value = vntInParam(i)
				'case 4 : .txtDIVAMT.value = vntInParam(i)
			end select
		next
		
		'★★★★★★★★★★★★★★★★★★IN 파라메터 및 조회를 위한 추가 파라메터 까지
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
		'txtDIVAMT_onblur
	End with
        With frmThis
			'메인쉬트
            gSetSheetColor mobjSCGLSpr, .sprSht 
			mobjSCGLSpr.SpreadLayout .sprSht, 5, 0
			mobjSCGLSpr.SpreadDataField .sprSht, "JOBNO|SEQ|BREAKDOWN|AMT|ACCDAY"
			mobjSCGLSpr.SetHeader .sprSht,         "제작번호|순번|내역|진행비|거래일"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "       0|   0|  28|    15|  10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ACCDAY", -1, -1, 10
			mobjSCGLSpr.ColHidden .sprSht, "JOBNO|SEQ", true
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "BREAKDOWN", -1, -1, 255
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
			mobjSCGLSpr.SetCellAlign2 .sprSht, "BREAKDOWN",-1,-1,0,2,false
			
			'Sum 쉬트
			gSetSheetColor mobjSCGLSpr, .sprShtSum
			mobjSCGLSpr.SpreadLayout .sprShtSum, 5, 1, 0,0,1,1,1,false,true,true,1
			mobjSCGLSpr.SpreadDataField .sprShtSum, "JOBNO|SEQ|BREAKDOWN|AMT|ACCDAY"
			mobjSCGLSpr.SetText .sprShtSum, 2, 1, "합 계"
			mobjSCGLSpr.SetScrollBar .sprShtSum, 0
			mobjSCGLSpr.SetBackColor .sprShtSum,"1|2",rgb(205,219,215),false
			mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "AMT", -1, -1, 0
			mobjSCGLSpr.ColHidden .sprShtSum, "JOBNO|SEQ", true
			mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
			mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "15"
			.sprSht.focus
        End With
        
        SelectRtn
        SUM_AMT
end sub

Sub EndPage()
	set mobjPDCMEXE = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMEXE.SelectRtn_ACCLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,frmthis.txtJOBNO.value)

		if not gDoErrorRtn ("SelectRtn_ACCLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
		
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'Call SUM_AMT ()
   		end if
   	end with
end sub
Sub DeleteRtn_DTL
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	dim strJOBNO,strCUST,strSEQ
	Dim lngSUMAMT,lngSUMAMT2
	Dim strPREESTNO
	Dim dblSEQ
	Dim strSUMAMT
	On error resume next
	
	with frmThis
		'한 건씩 삭제할 경우
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)

		if gDoErrorRtn ("DeleteRtn_Dtl") then exit sub

		if intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		
		strJOBNO = ""
		strSEQ = 0
	
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			strJOBNO = Trim(.txtJOBNO.value) 
			dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))	
			if mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)) <> ""  then
				intRtn = mobjPDCMEXE.DeleteRtn_ACCLIST(gstrConfigXml,strJOBNO,dblSEQ)
			end if
			
			if not gDoErrorRtn ("DeleteRtn_ACCLIST") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				'합계재계산
				gWriteText "", "자료가 삭제" & mePROC_DONE
				
   			end if
		next
		SUM_AMT
		strSUMAMT = mobjSCGLSpr.GetTextBinding( .sprShtSum,"AMT",1)
		strSUMAMT = replace(strSUMAMT,",","")
		intRtn = mobjPDCMEXE.DeleteUpdate_ACCLIST(gstrConfigXml,strJOBNO,strSUMAMT)
		if not gDoErrorRtn ("DeleteUpdate_ACCLIST") then
		Else
		gErrorMsgBox "삭제후 외주비정산 금액 갱신에 실패하였습니다.","삭제안내!"		
   		end if
		mobjSCGLSpr.DeselectBlock .sprSht
		mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		
	end with
End Sub

'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
	End with
end sub
'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub
Sub SUM_AMT()
	Dim lngCnt
	Dim strSUMDEMANDAMT
	Dim strDIVAMT
	strSUMDEMANDAMT = 0
	With frmThis
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		
		mobjSCGLSpr.SetTextBinding .sprShtSum,"AMT",1, strSUMDEMANDAMT
	End With
End Sub


Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strJOBNO,strDEMANDAMT,strJOBYEARMON
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt,intCnt2
	Dim strSUMAMT 
	with frmThis
   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNO|SEQ|AMT|BREAKDOWN|ACCDAY")
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "디테일 데이터를 입력 하십시오"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		intRtn = mobjPDCMEXE.ProcessRtn_ACCLIST(gstrConfigXml,vntData,.txtJOBNO.value )
	
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
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
  	
		'AND mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt) = "" AND (mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = 0) 
   	End with
	DataValidation = true
End Function

-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%"  HEIGHT="100%"  border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">진행비내역&nbsp;관리
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 225px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 128px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
													alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"
													align="absMiddle"></TD>
											<TD style="WIDTH: 126px"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF"
													width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'"
													alt="한 행 삭제" src="../../../images/imgDelRow.gif" width="54" border="0" name="imgDelRow"
													align="absMiddle">
											</TD>		
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
							<TABLE id="tblTitle" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</tr>
						</table>
						<TABLE id="tblBody" style="HEIGHT: 100%" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 90%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="20214">
											<PARAM NAME="_ExtentY" VALUE="9287">
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
											<!--904-->
										</OBJECT>
										<OBJECT id="sprShtSum" style="WIDTH: 100%; HEIGHT: 23px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="20214">
											<PARAM NAME="_ExtentY" VALUE="609">
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
						<FONT face="굴림"><INPUT class="NOINPUT" id="txtJOBNO" style="WIDTH: 144px; HEIGHT: 22px" readOnly type="text"
								size="18" name="txtJOBNO"><INPUT class="NOINPUT" id="txtDIVAMT" style="WIDTH: 200px; HEIGHT: 22px" tabIndex="1" readOnly
								type="text" size="28" name="txtDIVAMT"></FONT>
				</TD>
				</FORM></TR>
		</TABLE>
	</body>
</HTML>
