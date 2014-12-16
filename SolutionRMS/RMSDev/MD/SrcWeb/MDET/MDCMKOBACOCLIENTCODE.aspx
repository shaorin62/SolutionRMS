<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMKOBACOCLIENTCODE.aspx.vb" Inherits="MD.MDCMKOBACOCLIENTCODE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>KOBACO 광고주코드 등록</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'컨트롤작성 : 
'엔티티작성 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
'전역변수 설정
Dim mobjMDETREALMEDCODEMST
Dim mobjMDCOGET
Dim mlngRowCnt,mlngColCnt
Dim mstrGUBUN

Dim mUploadFlag
mstrGUBUN = "KOBACO"

CONST meTAB = 9
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub


Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub


'텝처리 (코바코)
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
		
	pnlTab_KOBACO.style.visibility = "visible" 
	pnlTab_SBS.style.visibility = "hidden" 	
	
	document.getElementById("strMsgBox").innerHTML = "코바코 광고주 코드"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "KOBACO"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'텝처리 (SBS)
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TABON
	
	pnlTab_KOBACO.style.visibility = "hidden" 
	pnlTab_SBS.style.visibility = "visible" 
	
	document.getElementById("strMsgBox").innerHTML = "SBS 광고주 코드"
		
	gFlowWait meWAIT_ON
	mstrGUBUN = "SBS"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgHIGHCUSTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	
	With frmThis
		vntInParams = array(trim(.txtHIGHCUSTCODE.value), trim(.txtCUSTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtHIGHCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtHIGHCUSTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCUSTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCUSTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtHIGHCUSTCODE.value),trim(.txtCUSTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtHIGHCUSTCODE.value = trim(vntData(0,1))
					.txtCUSTNAME.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim intCnt
	Dim strKOBACOCUSTCODE
	Dim strRow
	
	strRow = Row
	
	with frmThis
		strKOBACOCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"KOBACOCUSTCODE",Row)
	
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding( .sprSht,"KOBACOCUSTCODE",intCnt) <> "" Then	
				if intCnt = strRow Then
				Else
					If strKOBACOCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"KOBACOCUSTCODE",intCnt) Then
						gErrorMsgBox "사용중인코드 입니다.","입력안내"
						mobjSCGLSpr.SetTextBinding .sprSht,"KOBACOCUSTCODE",Row, ""
						exit Sub
					End If
				End If
			End If	
		Next 
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub sprSht_SBS_Change(ByVal Col, ByVal Row)
	Dim intCnt
	Dim strSBSCUSTCODE
	Dim strRow
	
	strRow = Row
	
	with frmThis
		strSBSCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_SBS,"SBSCUSTCODE",Row)
	
		For intCnt = 1 To .sprSht_SBS.MaxRows
			If mobjSCGLSpr.GetTextBinding( .sprSht_SBS,"SBSCUSTCODE",intCnt) <> "" Then	
				if intCnt = strRow Then
				Else
					If strSBSCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_SBS,"SBSCUSTCODE",intCnt) Then
						gErrorMsgBox "사용중인코드 입니다.","입력안내"
						mobjSCGLSpr.SetTextBinding .sprSht_SBS,"SBSCUSTCODE",Row, ""
						exit Sub
					End If
				End If
			End If	
		Next 
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_SBS, Col, Row
End Sub

'시트 더블클릭
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser .sprSht, ""
		end if
	end with
end sub

sub sprSht_SBS_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser .sprSht_SBS, ""
		end if
	end with
end sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Set mobjMDETREALMEDCODEMST	= gCreateRemoteObject("cMDET.ccMDETREALMEDCODEMST")
	Set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	gInitComParams mobjSCGLCtl,"MC"
	
	With frmThis
		'탭 위치 설정 및 초기화
		mobjSCGLCtl.DoEventQueue
		gSetSheetDefaultColor
	
	End With
	InitPageData	
	btnTab1_onclick
End Sub

Sub gridLayOut
	with frmThis
		if mstrGUBUN = "KOBACO" THEN
			'**************************************************
			'***KOBACO Sheet 디자인
			'**************************************************	
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 3, 0, 0
			mobjSCGLSpr.SpreadDataField .sprSht, "HIGHCUSTCODE | CUSTNAME | KOBACOCUSTCODE"
			mobjSCGLSpr.SetHeader .sprSht,			"광고주코드|광고주명|코바코운영코드"
			mobjSCGLSpr.SetColWidth .sprSht, "-1",  "22|60|40"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "KOBACOCUSTCODE ", -1, -1, 8 '
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTNAME ", -1, -1, 255 'By 태호: 글씨 짤림
			mobjSCGLSpr.SetCellAlign2 .sprSht, "HIGHCUSTCODE | KOBACOCUSTCODE",-1,-1,2,2,false
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"HIGHCUSTCODE | CUSTNAME"
			
			'pnlTab_KOBACO.style.visibility = "visible" 
			'pnlTab_SBS.style.visibility = "hidden" 
		ELSE
		
			'**************************************************
			'***SBS Sheet 디자인
			'**************************************************	
			gSetSheetColor mobjSCGLSpr, .sprSht_SBS
			mobjSCGLSpr.SpreadLayout .sprSht_SBS, 3, 0, 0
			mobjSCGLSpr.SpreadDataField .sprSht_SBS, "HIGHCUSTCODE | CUSTNAME | SBSCUSTCODE"
			mobjSCGLSpr.SetHeader .sprSht_SBS,			"광고주코드|광고주명|SBS운영코드"
			mobjSCGLSpr.SetColWidth .sprSht_SBS, "-1",  "22|60|40"
			mobjSCGLSpr.SetRowHeight .sprSht_SBS, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht_SBS, "-1", "13"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SBS, "CUSTNAME | SBSCUSTCODE", -1, -1, 255 
			mobjSCGLSpr.SetCellAlign2 .sprSht_SBS, "HIGHCUSTCODE | SBSCUSTCODE",-1,-1,2,2,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_SBS,true,"HIGHCUSTCODE | CUSTNAME"
			
			'pnlTab_KOBACO.style.visibility = "hidden" 
			'pnlTab_SBS.style.visibility = "visible" 
		END IF 
	end with

END SUB

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'화면 초기값 설정
	gClearAllObject frmThis
	with frmThis
		gridLayOut
		.sprSht.MaxRows = 0	
		.sprSht_SBS.MaxRows = 0	
		
		document.getElementById("strMsgBox").innerHTML = "코바코 광고주 코드"
	END WITH
End Sub

Sub EndPage()
	set mobjMDETREALMEDCODEMST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

Sub SelectRtn (mstrGUBUN)
   	Dim vntData
   	Dim i, strCols
   	Dim strKOBACOCUSTCODE
   	Dim strHIGHCUSTCODE
   	Dim strCUSTNAME
   	
	With frmThis
		.sprSht.MaxRows = 0	
		.sprSht_SBS.MaxRows = 0	
		
		gridLayOut
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strKOBACOCUSTCODE	= .txtKOBACOCUSTCODE.value
		strHIGHCUSTCODE		= .txtHIGHCUSTCODE.value
		strCUSTNAME			= .txtCUSTNAME.value
		
		vntData = mobjMDETREALMEDCODEMST.SelectRtn_Client(gstrConfigXml,mlngRowCnt,mlngColCnt,strKOBACOCUSTCODE,strHIGHCUSTCODE, strCUSTNAME, mstrGUBUN)
		
		If not gDoErrorRtn ("SelectRtn_Client") then
			IF mstrGUBUN = "KOBACO" THEN
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			ELSE
				mobjSCGLSpr.SetClipBinding .sprsht_SBS, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			END IF 
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE			
   		End if
   	End With
End Sub

'저장로직
Sub ProcessRtn()
	Dim intRtn
   	Dim vntData
   	
	with frmThis
		'쉬트의 변경된 데이터만 가져온다.
		if mstrGUBUN = "KOBACO" then 
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"HIGHCUSTCODE | CUSTNAME | KOBACOCUSTCODE")
		else
			vntData = mobjSCGLSpr.GetDataRows(.sprsht_SBS,"HIGHCUSTCODE | CUSTNAME | SBSCUSTCODE")
		end if 
		
		'처리 업무객체 호출
		intRtn = mobjMDETREALMEDCODEMST.ProcessRtn_Client(gstrConfigXml,vntData,mstrGUBUN)
		if not gDoErrorRtn ("ProcessRtn_Client") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox intRtn & " 건 이 저장되었습니다.","저장안내"
			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="98%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="113" background="../../../images/back_p.gIF"
													border="0">
													<TR>
														<TD align="left" width="100%" height="2"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">코바코 광고주 등록</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 120px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtKOBACOCUSTCODE,'')"
												width="86"><span id="strMsgBox"></span>
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 96px"><INPUT class="INPUTL" id="txtKOBACOCUSTCODE" style="WIDTH: 150px; HEIGHT: 22px" accessKey="NUM"
													maxLength="20" size="10" name="txtKOBACOCUSTCODE"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtHIGHCUSTCODE,txtCUSTNAME)"
												width="86">광고주
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUTL" id="txtCUSTNAME" style="WIDTH: 312px; HEIGHT: 22px" maxLength="255"
													size="46" name="txtCUSTNAME"> <IMG id="ImgHIGHCUSTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgHIGHCUSTCODE"> <INPUT class="INPUT" id="txtHIGHCUSTCODE" style="WIDTH: 68px; HEIGHT: 22px" maxLength="6"
													size="6" name="txtHIGHCUSTCODE">
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD vAlign="middle" align="right" height="28">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
																type="button" value="KOBACO" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
																type="button" value="SBS" name="btnTab2">
														</TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab_KOBACO" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31882">
											<PARAM NAME="_ExtentY" VALUE="16007">
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
									</DIV>
									<DIV id="pnlTab_SBS" style="LEFT:7px; VISIBILITY:hidden; WIDTH:100%; POSITION:relative; HEIGHT:100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_SBS" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="16007">
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
											<PARAM NAME="MaxCols" VALUE="19">
											<PARAM NAME="MaxRows" VALUE="0">
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
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
