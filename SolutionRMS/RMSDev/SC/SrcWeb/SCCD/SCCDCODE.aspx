<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCDCODE.aspx.vb" Inherits="SC.SCCDCODE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공통코드 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
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
Dim mobjSCCCCODETR
Dim mobjSCCOGET
Dim mlngRowCnt,mlngColCnt

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
sub imgNew_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtCLASSCODESEARCH.focus
		.sprSht.focus
	End With 
end sub

'-----------------------------------
' 저장   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 닫기
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' 클래스코드 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------	
Sub imgCLASSSEARCH_onclick
	Call CLASSSEARCH_POP()
End Sub

Sub CLASSSEARCH_POP
	Dim vntRet, vntInParams
	with frmThis
		vntInParams = array(.txtCLASSCODESEARCH.value,.txtCLASSNAMESEARCH.value)
		vntRet = gShowModalWindow("SCCDCLass.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCLASSCODESEARCH.value = vntRet(0,0)	'Code값 저장
			.txtCLASSNAMESEARCH.value = vntRet(3,0)	'코드명 표시
			
		gSetChangeFlag .txtCLASSCODESEARCH
		end if
	end with
End Sub

Sub txtCLASSNAMESEARCH_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCCCODETR.GetCLASS(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLASSCODESEARCH.value,.txtCLASSNAMESEARCH.value)
			if not gDoErrorRtn ("txtCLASSNAMESEARCH_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLASSCODESEARCH.value = vntData(0,1)
					.txtCLASSNAMESEARCH.value = vntData(3,1)		
				Else
					Call CLASSSEARCH_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLASSSNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASS_CODE",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSSNAME",Row))

			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOCODETR.GetCLASS(gstrConfigXml,mlngRowCnt,mlngColCnt, strCode, strCodeName)

				If not gDoErrorRtn ("GetCLASS") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASS_CODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSLNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSMNAME",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSSNAME",Row, vntData(3,1)	
						
						.txtCLASSCODESEARCH.focus()
						.sprSht.focus()
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLASSSNAME"), Row
						.txtCLASSCODESEARCH.focus()
						.sprSht.focus()
					End If
   				End If
   			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"UPDATE_YN") Then
			if mobjSCGLSpr.GetTextBinding( .sprSht,"UPDATE_YN",Row) = "수정불가" then
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,.sprSht.ActiveRow,1,7,True
			else
				mobjSCGLSpr.SetCellsLock2 .sprSht,False,.sprSht.ActiveRow,6,7,True
			end if
		End If
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLASSSNAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASS_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSSNAME",Row)))
			
			vntRet = gShowModalWindow("SCCDCLass.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASS_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSLNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSMNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSSNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.txtCLASSCODESEARCH.focus
		.sprSht.Focus
	End With
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,False,frmThis.sprSht.ActiveRow,1,3,True
'		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,false,frmThis.sprSht.ActiveRow,2,2,true
'		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CUSTTYPE",frmThis.sprSht.ActiveRow, "비계열"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"FLAG",frmThis.sprSht.ActiveRow, "NEW"
'		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLASSCODESEARCH.focus
		frmThis.sprSht.focus
	End If
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASS_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSSNAME",Row)))
								
			vntRet = gShowModalWindow("SCCDCLass.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASS_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSLNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSMNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSSNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		.txtCLASSCODESEARCH.focus
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	End With
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjSCCCCODETR	= gCreateRemoteObject("cSCCO.ccSCCOCODETR")
	set mobjSCCOGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
	with frmThis
		'**************************************************
		'***Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 1, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "CLASS_CODE | BTN | CLASSSNAME | CLASSLNAME | CLASSMNAME | CODE | CODE_NAME | SORT_SEQ | USE_YN | UPDATE_YN | FLAG"
		mobjSCGLSpr.SetHeader .sprSht,		    "클래스코드|클래스명|대분류|중분류|공통코드|공통코드명|정렬순서|사용구분|수정여부|구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "        13|2|    18|    14|    17|      10|        18|      10|      10|      10|0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "USE_YN", -1, -1, "사용" & vbTab & "미사용" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "UPDATE_YN", -1, -1, "수정가능" & vbTab & "수정불가" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SORT_SEQ",-1,-1,2,2,false
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLASS_CODE | CLASSLNAME | CLASSMNAME | CLASSSNAME | CODE | CODE_NAME | SORT_SEQ | FLAG", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CLASS_CODE | BTN | CLASSSNAME | CLASSLNAME | CLASSMNAME | FLAG "
	
		pnlTab1.style.visibility = "visible" 

		'화면 초기값 설정
		InitPageData
	End with
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis

	'초기 데이터 설정
	With frmThis
		.sprSht.MaxRows = 0
	End With
End Sub

Sub EndPage()
	set mobjSCCCCODETR = Nothing
	set mobjSCCOGET = Nothing
	gEndPage	
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   	Dim strCODE, strCODENAME, strUSE_YN, strCLASSCODE
   	Dim intCnt, intCnt2
   	Dim strRows
	
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		strCODE		= .txtCODESEARCH.value
		strCODENAME = .txtCODENAMESEARCH.value
		strUSE_YN	= .cmbUSE_YNSEARCH.value
		strCLASSCODE= .txtCLASSCODESEARCH.value
		
		vntData = mobjSCCCCODETR.SelectRtn_CODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE,strCODENAME,strUSE_YN,strCLASSCODE)
		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True

   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			
   			For intCnt = 1 To .sprSht.MaxRows
   				if mobjSCGLSpr.GetTextBinding( .sprSht,"UPDATE_YN",intCnt) = "수정불가" then
					If intCnt2 = 1 Then
						strRows = intCnt
					Else
						strRows = strRows & "|" & intCnt
					End If
					intCnt2 = intCnt2 + 1
				end if
				If mobjSCGLSpr.GetTextBinding(.sprSht,"USE_YN",intCnt) = "미사용" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
				End If
			Next
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,1,7,True
   		end if
   	end with
End Sub

'여기까지 쉬트 버튼 클릭
Sub ProcessRtn()
   	Dim intRtn
   	Dim vntData
	Dim lngCol, lngRow
	Dim strDataCHK
	With frmThis
		
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "CLASS_CODE | CODE | CODE_NAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 클래스코드/클래스코드명/공통코드/공통코드명은 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CLASS_CODE | BTN | CLASSSNAME | CLASSLNAME | CLASSMNAME | CODE | CODE_NAME | SORT_SEQ | USE_YN | UPDATE_YN | FLAG")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If
		
		intRtn = mobjSCCCCODETR.ProcessRtn_CODE(gstrConfigXml,vntData)
		
		If not gDoErrorRtn ("ProcessRtn_CODE") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			SelectRtn
   		End If
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
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="78" background="../../../images/back_p.gIF"
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
											<td class="TITLE">공통코드관리&nbsp;</td>
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
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 119px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCODESEARCH,'')">&nbsp;공통코드
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 99px"><INPUT class="INPUT_L" id="txtCODESEARCH" style="WIDTH: 96px; HEIGHT: 22px" type="text"
													maxLength="8" size="10" name="txtCODESEARCH"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCODENAMESEARCH,'')">공통코드명
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 104px"><INPUT class="INPUT_L" id="txtCODENAMESEARCH" style="WIDTH: 119px; HEIGHT: 22px" type="text"
													maxLength="255" size="13" name="txtCODENAMESEARCH"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 80px">사용구분
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 108px"><SELECT id="cmbUSE_YNSEARCH" title="사용구분" style="WIDTH: 104px" name="cmbUSE_YNSEARCH">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="Y">사용</OPTION>
													<OPTION value="N">미사용</OPTION>
												</SELECT>
											</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 106px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLASSNAMESEARCH,txtCLASSCODESEARCH)">클래스코드&nbsp;
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLASSNAMESEARCH" title="담당부서명" style="WIDTH: 168px; HEIGHT: 22px"
													type="text" maxLength="100" size="22" name="txtCLASSNAMESEARCH"> <IMG id="imgCLASSSEARCH" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="imgCLASSSEARCH"> <INPUT class="INPUT_L" id="txtCLASSCODESEARCH" title="담당부서코드" style="WIDTH: 65px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="20" size="5" name="txtCLASSCODESEARCH">
											<td class="SEARCHDATA" width="50">
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
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
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
											<TD align="left" width="400" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left">
															<TABLE cellSpacing="0" cellPadding="0" width="130" background="../../../images/back_p.gIF"
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
														<td class="TITLE">공통코드 등록 및 변경&nbsp;</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
																src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
												<!--Common Button End-->
											</TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="14579">
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
