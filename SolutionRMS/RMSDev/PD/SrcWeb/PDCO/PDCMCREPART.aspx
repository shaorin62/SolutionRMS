<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCREPART.aspx.vb" Inherits="PD.PDCMCREPART" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>사용자 권한관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 권한 관리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/12/01 By KIM TAE HO
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
Dim mobjPDCOCREPART
Dim mlngRowCnt,mlngColCnt
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
	mobjSCGLSpr.ExportExcelFile frmThis.sprSht
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


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정

	with frmThis

		if  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLASS_CODE")  Then
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",Row) = "PD_GRAPHICKIND" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",.sprSht.ActiveRow, "PA01"
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",Row) = "PD_ELECKIND" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",.sprSht.ActiveRow, "PA02"
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",Row) = "PD_PROMOTIONKIND" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",.sprSht.ActiveRow, "PA05"
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",Row) = "PD_INTERNETKIND" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",.sprSht.ActiveRow, "PA07"
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",Row) = "PD_OTHERSKIND" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",.sprSht.ActiveRow, "PA08"
			End If
		End If
	End with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


Sub cmbJOBGUBN1_onChange ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	Dim vntJOBGUBN
	
    With frmThis   
		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntJOBGUBN = mobjPDCOCREPART.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  '제작구분

		if not gDoErrorRtn ("COMBO_TYPE") then 
			mobjSCGLSpr.TypeComboBox = True 
			gLoadComboBox .cmbJOBGUBN1, vntJOBGUBN, False
   		end if    	
   	end with     	
End Sub
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjPDCOCREPART = gCreateRemoteObject("cPDCO.ccPDCOCREPART")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
	
    Call Grid_Layout()
    	'화면 초기값 설정
	InitPageData	
End Sub

Sub Grid_Layout()
	
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CLASS_CODE|CODE|SC_BU_CODE|CODE_NAME|SORT_SEQ|USE_YN|UPDATE_YN|ATTR02|DEBTOR|ACCOUNT|ATTR01|INSERTYN"
		mobjSCGLSpr.SetHeader .sprSht,		    "제작종류|분류코드|공통|매체분류명|정렬구분|사용구분|수정구분|제작코드|수익인식|외주용역비|JOB종류|저장여부"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "20      |10      |10  |25        |8       |10      |10      |12      |24      |24        |10     |10"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CODE|SORT_SEQ|ATTR02",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLASS_CODE|CODE_NAME|DEBTOR|ACCOUNT",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CLASS_CODE|CODE|SC_BU_CODE|CODE_NAME|SORT_SEQ|USE_YN|UPDATE_YN|ATTR02|DEBTOR|ACCOUNT|ATTR01|INSERTYN"
		mobjSCGLSpr.ColHidden .sprSht, "SC_BU_CODE|USE_YN|UPDATE_YN|ATTR01|ATTR02|INSERTYN",true
		'mobjSCGLSpr.CellGroupingEach .sprSht,"CLASS_CODE"
	End with
	
	
	
	pnlTab1.style.visibility = "visible" 
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
	
	with frmThis
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCOCREPART.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.cmbJOBGUBN1.value )
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 1 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CLASS_CODE|CODE|SC_BU_CODE|CODE_NAME|SORT_SEQ|USE_YN|UPDATE_YN|ATTR02|DEBTOR|ACCOUNT|ATTR01"
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
	if DataValidation =false then exit sub 	
	
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CLASS_CODE|CODE|SC_BU_CODE|CODE_NAME|SORT_SEQ|USE_YN|UPDATE_YN|ATTR02|DEBTOR|ACCOUNT|ATTR01|INSERTYN")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtnSave = gYesNoMsgbox("적용된 항목은 RMS운영팀 이외에 삭제,수정 하실수 없습니다." & vbcrlf & "제작매체분류를 적용하시겠습니까?","저장안내")
		IF intRtnSave <> vbYes then exit Sub
		
		'처리 업무객체 호출
		intRtn = mobjPDCOCREPART.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 0 Then
				gErrorMsgBox "저장되었습니다.","저장안내"
			End If
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
   	Dim intCnt
   	Dim dblSumAmt
   	Dim dblAMT
	'On error resume next
	with frmThis
  	
		
   		
   		
   		
		
   		for intCnt = 1 to .sprSht.MaxRows
   			'Sheet 필수 입력사항
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS_CODE",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht,"DEBTOR",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht,"ACCOUNT",intCnt) = "" Then 
				gErrorMsgBox "제작종류,수익인식,외주용역비 는 필수 입니다.","저장오류"
				Exit Function
			End if	
		next
   		
   	End with
   	
	DataValidation = true
End Function

Sub EndPage()
	set mobjPDCOCREPART = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
		CALL COMBO_TYPE()
		.cmbJOBGUBN1.selectedIndex = -1
		Get_COMBO_CLASSCODE
		Get_COMBO_DEBTOR
		Get_COMBO_ACCOUNT
	End with
End Sub

Sub Get_COMBO_CLASSCODE ()		
	Dim vntData_Demand
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCOCREPART.GetDataType_class(gstrConfigXml,mlngRowCnt,mlngColCnt)
						

		If not gDoErrorRtn ("GetDataType_class") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CLASS_CODE",,,vntData_Demand,,160		
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

Sub Get_COMBO_DEBTOR ()		
	Dim vntData_Demand
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCOCREPART.GetDataType_debtor(gstrConfigXml,mlngRowCnt,mlngColCnt)
						

		If not gDoErrorRtn ("GetDataType_debtor") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DEBTOR",,,vntData_Demand,,190		
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

Sub Get_COMBO_ACCOUNT ()		
	Dim vntData_Demand
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCOCREPART.GetDataType_account(gstrConfigXml,mlngRowCnt,mlngColCnt)
						

		If not gDoErrorRtn ("GetDataType_account") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "ACCOUNT",,,vntData_Demand,,190		
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		


Sub imgRowAdd_onclick ()
Dim intCnt
Dim intChk
with frmThis
	intChk =0 
	For intCnt = 1 To .sprSht.MaxRows
		If mobjSCGLSpr.GetTextBinding(.sprSht,"INSERTYN",intCnt) = "Y" Then
			intChk = intChk +1
		End IF
		If intChk <> 0 Then
			gErrorMsgBox "저장은 한번에 한코드씩 가능 합니다.","행추가안내"
			Exit Sub
		End If
	Next
End with
call sprSht_Keydown(meINS_ROW, 0)
End Sub

Sub sprSht_Keydown(KeyCode, Shift)

	Dim intRtn
	
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: 
		End Select

End Sub

Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"SC_BU_CODE",.sprSht.ActiveRow, "MC"
		mobjSCGLSpr.SetTextBinding .sprSht,"USE_YN",.sprSht.ActiveRow, "Y"
		mobjSCGLSpr.SetTextBinding .sprSht,"UPDATE_YN",.sprSht.ActiveRow, "N"
		mobjSCGLSpr.SetTextBinding .sprSht,"ATTR02",.sprSht.ActiveRow, "K"
		mobjSCGLSpr.SetTextBinding .sprSht,"INSERTYN",.sprSht.ActiveRow, "Y"
		mobjSCGLSpr.SetCellsLock2 .sprSht,false,"CLASS_CODE|SC_BU_CODE|CODE_NAME|USE_YN|UPDATE_YN|ATTR02|DEBTOR|ACCOUNT|ATTR01",.sprSht.ActiveRow,.sprSht.ActiveRow,false
	End With
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="108" background="../../../images/back_p.gIF"
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
											<td class="TITLE">제작매체 분류관리</td>
										</tr>
									</table>
								</td>
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
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1">
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" onclick="vbscript:Call gCleanField(cmbJOBGUBN1, '')"
												width="90">&nbsp;매체종류
											</TD>
											<TD class="SEARCHDATA" ><SELECT id="cmbJOBGUBN1" title="시스템구분" style="WIDTH: 168px" name="cmbJOBGUBN1">
												</SELECT></TD>
											<td class="SEARCHDATA2" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</td>
										</TR>
									</TABLE>
									<!--적용시작-->
									<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF"
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
											<td class="TITLE">제작매체 분류</td>
										</tr>
									</table>
								</td>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'"
																height="20" alt="입력을 위하여 행을 추가합니다." src="../../../images/imgRowAdd.gIF"
																border="0" name="imgRowAdd" align="absMiddle">
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
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="42466">
											<PARAM NAME="_ExtentY" VALUE="16272">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
