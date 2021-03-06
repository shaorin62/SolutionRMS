<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOCUSTMPPLIST.aspx.vb" Inherits="SC.SCCOCUSTMPPLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>거래선 등록(MPP)</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOCUSTMPPLIST.aspx
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/09 By KTY
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script type="text/javascript">
		
function Set_IframeValue(strBUSINO,intCNT) {
	var value1  = strBUSINO;
	var value2  = intCNT;
	//iframe 서버컨트롤 텍스트 박스 busino 입력
	var textbox1 = frmSapCon.document.getElementById("<%=txtSAPBUSINO.ClientID%>");
	var textbox2 = frmSapCon.document.getElementById("<%=txtCNT.ClientID%>");
	
	textbox1.value = value1;
	textbox2.value = value2;
	window.frames[0].document.forms[0].submit();
}

		</script>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOCUSTLIST '공통코드, 클래스
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'---------------------------------------------------
' 신규 SAP 값받아오기
'---------------------------------------------------
Sub Set_CustValue (strVALUE)
	With frmThis
		If MID(strVALUE,InStr(1,strVALUE,"|"),len(strVALUE)) <> "||||||||||||||" Then
			gErrorMsgBox "SAP에 존재하는 거래처 사업자번호입니다. 다른 사업자번호를 입력하세요","저장안내"
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",	.sprSht_CUST.ActiveRow, ""
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
			Exit Sub
		End If
	End With
End Sub

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' 명령 버튼 클릭 이벤트
'---------------------------------------------------
'-----------------------------------
'조회
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'행추가
'-----------------------------
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_CUST_Keydown(meINS_ROW, 0)
		.txtBUSINO.focus
		.sprSht_CUST.focus
	End With 
end sub

'-----------------------------------
' 저장   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht_CUST.MaxRows = 0 Then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn_MPPHDR
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_CUST
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'삭제
'-----------------------------------
Sub imgDelete_onclick ()
	Dim i
	If frmThis.sprSht_CUST.MaxRows = 0 Then
		gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 닫기
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_CUST_Change(ByVal Col, ByVal Row)
	Dim i
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_CUST,"BUSINO") Then
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow) <> "" Then
				If Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))) = 10 Then
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),1,3) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),4,2) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),6,5)
				elseIf Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))) = 13 Then
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),1,6) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),7,7)
				else
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))
				End If
				
				If Busino_Check =False Then Exit Sub
			End If
			
			Set_IframeValue TRIM(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",Row)) , 1
		End If
		mobjSCGLSpr.CellChanged .sprSht_CUST, Col, Row
	End With
End Sub

'상단그리드 사업자 번호 입력시 사업자번호 중복 체크
Function Busino_Check ()
	Busino_Check = false
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCCOCUSTLIST.Busino_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												trim(Replace(mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),"-","")), "P")
		
		If mlngRowCnt > 0 Then
			gErrorMsgBox "상위거래처에 중복된 사업자번호가 있습니다.",""
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow,""
			'포커스를 시트로 이동시킨다.
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
			Exit Function
   		End if
   	End With
   	Busino_Check = True
End Function

'-----------------------------------
'쉬트 더블클릭
'-----------------------------------
Sub sprSht_CUST_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CUST, ""
		End if
	End With
End Sub

'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_CUST_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_CUST, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht_CUST,false,frmThis.sprSht_CUST.ActiveRow,mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),true
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"CUSTTYPE",frmThis.sprSht_CUST.ActiveRow, "비계열"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"USE_FLAG",frmThis.sprSht_CUST.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_CUST, 1,frmThis.sprSht_CUST.MaxRows
		frmThis.txtBUSINO.focus
		frmThis.sprSht_CUST.focus
		
	End If
End Sub

Sub txtCLIENTNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtBUSINO_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjSCCOCUSTLIST = gCreateRemoteObject("cSCCO.ccSCCOCUSTLIST")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
	
		gSetSheetColor mobjSCGLSpr, .sprSht_CUST	
		mobjSCGLSpr.SpreadLayout .sprSht_CUST, 7, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_CUST, "CHK | BUSINO | CUSTNAME | HIGHCUSTCODE | USE_FLAG | CUSTTYPE | UUSER"
		mobjSCGLSpr.SetHeader .sprSht_CUST,		  "선택|사업자번호|거래처명|코드|사용|계열|입력자"
		mobjSCGLSpr.SetColWidth .sprSht_CUST, "-1", " 4|        13|      35|   8|   5|   8|     6"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CUST, "CHK | USE_FLAG"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | CUSTNAME | HIGHCUSTCODE | CUSTTYPE | ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_CUST, true, "BUSINO | HIGHCUSTCODE | UUSER"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_CUST, "CUSTTYPE", -1, -1, "계열" & vbTab & "비계열" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | CUSTNAME | HIGHCUSTCODE", -1, -1, 200
		mobjSCGLSpr.SetCellAlign2 .sprSht_CUST, "BUSINO | HIGHCUSTCODE | CUSTTYPE" ,-1,-1,2,2,false
		
		.sprSht_CUST.style.visibility = "visible"

    End With

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOCUSTLIST = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis

	'초기 데이터 설정
	With frmThis
		.sprSht_CUST.MaxRows = 0
	End With
End Sub

'------------------------------------------
' HDR 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTNAME, strBUSINO
   	Dim intCnt
   	
	With frmThis

		'Sheet초기화
		.sprSht_CUST.MaxRows = 0
		
		'변수 초기화
		strCLIENTNAME = "" : strBUSINO = ""
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value 
		strBUSINO		= .txtBUSINO.value
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_MPPHDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strCLIENTNAME, strBUSINO)

		If not gDoErrorRtn ("SelectRtn_MPPHDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_CUST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			For intCnt = 1 To .sprSht_CUST.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"USE_FLAG",intCnt) = "0" Then
					mobjSCGLSpr.SetCellShadow .sprSht_CUST, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
				End If
			Next
			
   			gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
   		End if
   	End With
End Sub

'------------------------------------------
' HDR 수정/저장 처리 
'------------------------------------------
Sub ProcessRtn_MPPHDR ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	With frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_CUST, "BUSINO|CUSTNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 사업자번호/거래처명은 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_CUST,"CHK | BUSINO | CUSTNAME | HIGHCUSTCODE | USE_FLAG | CUSTTYPE")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_MPPHDR(gstrConfigXml,vntData, "P")
	
		If not gDoErrorRtn ("ProcessRtn_MPPHDR") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_CUST,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			SelectRtn
   		End If
   	End With
End Sub

'------------------------------------------
'데이터 삭제
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strHIGHCUSTCODE
	Dim strHIGHCUSTCODE2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht_CUST.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",i)
				If strHIGHCUSTCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					vntData = mobjSCCOCUSTLIST.SelectRtn_MPPCountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODE) 
					If mlngRowCnt > 0 Then
						strMSG = ""
						strMSG = strMSG & " 케이블: " & vntData(0,0) & "건" 
						gErrorMsgBox i & "행의 코드는 " & strMSG & " 이 청약데이터로 저장되어있습니다.","삭제안내!"
						Exit Sub
					End If
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		For i = .sprSht_CUST.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE2 = mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"HIGHCUSTCODE",i)
			
				If strHIGHCUSTCODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					intRtn = mobjSCCOCUSTLIST.DeleteRtn_EXE(gstrConfigXml, strHIGHCUSTCODE2)
					
					IF not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_CUST,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn") Then
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
		SelectRtn
	End With
	err.clear
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF"
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
												<td class="TITLE">MPP 관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,'')"
													width="100">MPP명</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 200px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="28" name="txtCLIENTNAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
													width="100">사업자번호</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtBUSINO" title="코드조회" style="WIDTH: 168px; HEIGHT: 22px" type="text"
														maxLength="15" align="left" name="txtBUSINO">
													<asp:textbox id="txtSAPBUSINO" runat="server" Width="8px" Visible="False"></asp:textbox>
													<asp:textbox id="txtCNT" runat="server" Visible="false" Width="8px"></asp:textbox></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="자료를 인쇄합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_CUST" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="16378">
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
									</TD>
								</TR>
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 10px; HEIGHT: 10px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
