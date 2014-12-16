<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMROLELIST.aspx.vb" Inherits="PD.PDCMROLELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>사용자 권한관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/마감 등록 화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
Dim mobjPDCMROLE
Dim mobjPDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mstrCheck
mstrCheck = True
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

Sub imgQuery_onclick
Dim vntData

with frmThis
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End with
	
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub imgSave_onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgAuthCopy_onclick

	Dim vntRet
	Dim vntInParams
	Dim strUSERID
	Dim strUSERNAME
	Dim strCOPYUSERID
	Dim strCOPYUSERNAME
	Dim intRtnSave
	Dim intRtn
	with frmThis
		
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "복사할 사원을 우선 조회하십시오.","권한복사 안내"
			Exit Sub
		Else 
			strUSERNAME = .txtEMPNAME.value 
			strUSERID = .txtEMPNO.value 
		End If
		
		vntRet = gShowModalWindow("PDCMEMPAUTHPOP.aspx","" , 413,435)
		if isArray(vntRet) then
			strCOPYUSERID = trim(vntRet(0,0))
			strCOPYUSERNAME = trim(vntRet(1,0))
     	end if
     	If strCOPYUSERID <> "" And strCOPYUSERNAME <> "" Then
     		intRtnSave = gYesNoMsgbox( strCOPYUSERNAME & " 님 권한을 " & strUSERNAME & " 님 권한 으로 적용 하시겠습니까?","권한복사 안내")
			IF intRtnSave <> vbYes then exit Sub
			
			intRtn = mobjPDCMROLE.ProcessRtn_Copy(gstrConfigXml,strUSERID,strCOPYUSERID)
			if not gDoErrorRtn ("ProcessRtn_Copy") then
					if intRtn > 0 Then
					gErrorMsgBox "권한이 적용되었습니다.","권한복사 안내"
					End If
				SelectRtn
   			end if
     	Else
     		'gErrorMsgBox "권한복사 대상자를 선택 하지 않으셨으므로, 작업이 중단됩니다.","권한복사 안내!" '선택 취소 시 메세지 보여달라고 할경우 주석 제거
     	End If
	End with
	gSetChange

End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	dim intcnt
	with frmThis
		If Row = 0 and Col = 4  then 
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 4,4,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
				
			next
		end if
	end with
End Sub
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()

	'탭 위치 설정 및 초기화
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "152px"
	'pnlTab1.style.height ="300px"
	'pnlTab1.style.left= "7px"
	
	
	'서버업무객체 생성	
	Set mobjPDCMROLE = gCreateRemoteObject("cPDCO.ccPDCOROLEMST")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'mobjPDCMGET
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
	
    Call Grid_Layout()
    	'화면 초기값 설정
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 5, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "SYSTEMLEVELNAME|PROGRAMID|PROGRAMNAME|USEYN|LOOPCHK"
		mobjSCGLSpr.SetHeader .sprSht,		    "시스템구분|프로그램ID|프로그램명|사용구분|업데이트"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "20        |15      |78          |10      |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USEYN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PROGRAMID",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SYSTEMLEVELNAME|PROGRAMNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SYSTEMLEVELNAME|PROGRAMID|PROGRAMNAME"
		mobjSCGLSpr.ColHidden .sprSht, "LOOPCHK",true
		mobjSCGLSpr.CellGroupingEach .sprSht,"SYSTEMLEVELNAME"
		
	End with
	
	'DateClean
	'pnlTab1.style.visibility = "" 
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'검색조건 시작일
Sub imgFrom_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtFrom,.imgFrom,"txtFrom_onchange()"
		gSetChange
	end with
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub

'검색조건 종료일
Sub imgTo_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtTo,.imgTo,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtYEARMON_onchange
	gSetChange
End Sub
Sub txtTo_onchange
	gSetChange
End Sub


Sub SelectRtn ()

   	Dim vntData
   	Dim i, strCols
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		If .txtEMPNO.value = "" Then
			gErrorMsgbox "사원번호는 반드시 선택 하십시오.","조회안내"
			.sprSht.MaxRows = 0
			Exit Sub
		End If
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMROLE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value,.txtEMPNAME.value,.cmbGUBN.value )
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 1 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			Else
   			.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub ProcessRtn()
	Dim intRtn
   	Dim vntData
   	Dim intRtnSave
   	Dim strUSERID
   	Dim intCnt
	with frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		'For intCnt = 1 To .sprSht.MaxRows
		'	mobjSCGLSpr.SetTextBinding .sprSht,"LOOPCHK",intCnt, "T"
		'	sprSht_Change 5, intcnt
		'Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SYSTEMLEVELNAME|PROGRAMID|PROGRAMNAME|USEYN|LOOPCHK")
		
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "사원번호를 확인하십시오.","저장안내"
			.sprSht.MaxRows = 0
			Exit Sub
		Else
		strUSERID =  .txtEMPNO.value
		End If
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtnSave = gYesNoMsgbox("사용자권한을 적용하시겠습니까?","저장안내")
		IF intRtnSave <> vbYes then exit Sub
		
		
		'처리 업무객체 호출
		intRtn = mobjPDCMROLE.ProcessRtn(gstrConfigXml,vntData,strUSERID)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 0 Then
			gErrorMsgBox intRtn & " 건 의 권한이 적용되었습니다.","저장안내"
			End If
			SelectRtn
   		end if
   	end with
End Sub

Sub EndPage()
	set mobjPDCMROLE = Nothing
	set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
	End with
End Sub
'-----------------------------------------------------------------------------------------
' 사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'실제 데이터List 가져오기
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()					' 포커스 이동
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
			
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					'.txtMEMO.focus()
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="3">
				<!--Top TR Start-->
				<TR>
					<TD >
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;사용자 권한관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="50" border="0">
										<TR>
											<TD></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="top" align="LEFT" colSpan="2"><FONT face="굴림">
										<TABLE id="tblKey" cellSpacing="1" cellPadding="0" width="1040" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, '')"
													width="90">&nbsp;시스템구분
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 390px"><SELECT id="cmbGUBN" title="시스템구분" style="WIDTH: 168px" name="cmbGUBN">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="SCCM">공통</OPTION>
														<OPTION value="MDEL">공중파</OPTION>
														<OPTION value="MDCA">케이블</OPTION>
														<OPTION value="MDPR">인쇄</OPTION>
														<OPTION value="MDIN">인터넷</OPTION>
														<OPTION value="MDOU">옥외</OPTION>
														<OPTION value="PDCM">제작의뢰</OPTION>
														<OPTION value="PDMA">제작관리</OPTION>
														<OPTION value="READ">광고비집행현황</OPTION>
														<OPTION value="REME">매체별집행내역</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO, txtEMPNAME)"
													width="90">사용자</TD>
												<TD class="SEARCHDATA" style="WIDTH: 398px"><INPUT class="INPUT_L" id="txtEMPNAME" title="사원조회" style="HEIGHT: 22px" type="text" maxLength="255"
														align="left" size="38" name="txtEMPNAME"><IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
														width="23" align="absMiddle" border="0" name="ImgEMPNO"><INPUT class="INPUT" id="txtEMPNO" title="사번조회" style="WIDTH: 88px; HEIGHT: 22px" type="text"
														maxLength="8" align="left" size="9" name="txtEMPNO" readOnly></TD>
												<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
										<!--적용시작-->
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 1040px; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;사용자 권한설정</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px"  cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgAuthCopy" onmouseover="JavaScript:this.src='../../../images/imgAuthCopyOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAuthCopy.gIF'"
																	height="20" alt="기존사용자 권한을 선택 하여 현재 조회 되어있는 사용자 에게 적용합니다." src="../../../images/imgAuthCopy.gIF"
																	border="0" name="imgAuthCopy" align="absMiddle">
															</TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave" align="absMiddle">
															</TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"
																	align="absMiddle"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
										<!--적용끝-->
									</FONT>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; "></TD>
							</tr>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 98%; POSITION: relative" ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="23098">
											<PARAM NAME="_ExtentY" VALUE="18256">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
								</TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR> 
			<!--List End-->
			<!--BodySplit Start-->
			<TR>
				<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
			</TR>
			<!--BodySplit End-->
			<!--Brench Start-->
			<TR>
				<TD class="BRANCHFRAME" style="WIDTH: 1040px"><FONT face="굴림" color="#666666" size="3"></FONT>
					<!--<INPUT class="BUTTON" id="btn1" style="WIDTH: 123px; HEIGHT: 16pt" type="button" value="분기버튼"
											name="Button">--></TD>
			</TR>
			<!--Brench End-->
			<!--Bottom Split Start-->
			<TR>
				<TD class="BOTTOMSPLIT" style="WIDTH: 1040px" id="lblstatus"><FONT face="굴림"></FONT></TD>
			</TR>
			<!--Bottom Split End--> </TABLE> 
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
