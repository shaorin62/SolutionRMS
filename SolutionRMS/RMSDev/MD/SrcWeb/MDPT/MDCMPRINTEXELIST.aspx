<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTEXELIST.aspx.vb" Inherits="MD.MDCMPRINTEXELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>일자별 세부검색</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 일별/광고주별 조회
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/04 By Kim Tae Yub
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
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjEXECUTE'공통코드, 클래스
Dim mcomecalender
mcomecalender = FALSE
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
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtFPUB_DATE.value = "" AND frmThis.txtTPUB_DATE.value = "" then
		gErrorMsgBox "게재일을 입력하시오","조회안내"
		exit Sub
	end if

	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'****************************************************************************************
' 게재일 달력
'****************************************************************************************
Sub imgCalFrom_onclick
	'CalEndar를 화면에 표시
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtFPUB_DATE,frmThis.imgCalFrom,"txtFPUB_DATE_onchange()"
	mcomecalender = false
End Sub

Sub imgCalTo_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtTPUB_DATE,frmThis.imgCalTo,"txtTPUB_DATE_onchange()"
End Sub

Sub txtFPUB_DATE_onchange
	gSetChange
End Sub

Sub txtTPUB_DATE_onchange
	gSetChange
End Sub

Sub txtFPUB_DATE_onblur
	Dim strdate 
	Dim strPUB_DATE
	strdate = ""
	strPUB_DATE =""
	With frmThis
		strdate=.txtFPUB_DATE.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender Then
			strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
		else
			If len(strdate) = 4 Then
				strPUB_DATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
			elseif len(strdate) = 10 Then
				strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			elseif len(strdate) = 3 Then
				strPUB_DATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
			else
				strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
			End If
		End If
		
		If .txtFPUB_DATE.value <> "" Then 
			DateClean_Change strPUB_DATE
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("MDCMPRINTEXECUSTLISTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtFPUB_DATE.value = trim(vntRet(0,0))
			.txtTPUB_DATE.value = trim(vntRet(1,0))
			.txtCLIENTCODE.value = trim(vntRet(2,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			selectRtn
     	end if
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjEXECUTE.GetPRINTCLIENT_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt,TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetPRINTCLIENT_LIST") then
				If mlngRowCnt = 1 Then
					.txtFPUB_DATE.value = trim(vntData(0,1))
					.txtTPUB_DATE.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))
					.txtCLIENTNAME.value = trim(vntData(3,1))
					selectRtn
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 팀코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
Sub ImgTIMCODE_onclick
	Call TIMCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), _
							TRIM(.txtCLIENTCODE.value), TRIM(.txtCLIENTNAME.value), _
							TRIM(.txtTIMCODE.value), TRIM(.txtTIMNAME.value)) '<< 받아오는경우
	    
	    vntRet = gShowModalWindow("MDCMPRINTEXETIMLISTPOP.aspx",vntInParams , 413,455)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtFPUB_DATE.value = trim(vntRet(0,0))
			.txtTPUB_DATE.value = trim(vntRet(1,0))
			.txtTIMCODE.value = trim(vntRet(2,0))
			.txtTIMNAME.value = trim(vntRet(3,0))
			.txtCLIENTCODE.value = trim(vntRet(4,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(5,0))  ' 코드명 표시
			selectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjEXECUTE.GetPRINTTIM_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, _ 
												   TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), _
												   TRIM(.txtCLIENTCODE.value),TRIM(.txtCLIENTNAME.value), _
												   TRIM(.txtTIMCODE.value),TRIM(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetPRINTTIM_LIST") Then
				If mlngRowCnt = 1 Then
					.txtFPUB_DATE.value = trim(vntData(0,1))
					.txtTPUB_DATE.value = trim(vntData(1,1))
					.txtTIMCODE.value = trim(vntData(2,1))
					.txtTIMNAME.value = trim(vntData(3,1))
					.txtCLIENTCODE.value = trim(vntData(4,1))  ' Code값 저장
					.txtCLIENTNAME.value = trim(vntData(5,1))  ' 코드명 표시
					selectRtn
				Else
					Call TIMCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' 매체사코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgMEDCODE_onclick
	Call MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub MED_CODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtMEDCODE.value), trim(.txtMEDNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("MDCMPRINTEXEMEDLISTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtFPUB_DATE.value = trim(vntRet(0,0))
			.txtTPUB_DATE.value = trim(vntRet(1,0))
			.txtMEDCODE.value = trim(vntRet(2,0))
			.txtMEDNAME.value = trim(vntRet(3,0))
			selectRtn
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjEXECUTE.GetPRINTMED_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt,TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtMEDCODE.value), trim(.txtMEDNAME.value))
			if not gDoErrorRtn ("GetPRINTMED_LIST") then
				If mlngRowCnt = 1 Then
					.txtFPUB_DATE.value = trim(vntData(0,1))
					.txtTPUB_DATE.value = trim(vntData(1,1))
					.txtMEDCODE.value = trim(vntData(2,1))
					.txtMEDNAME.value = trim(vntData(3,1))
					selectRtn
				Else
					Call MED_CODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjEXECUTE	= gCreateRemoteObject("cMDSC.ccMDSCEXECUTE")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "PUB_DATE | CLIENTNAME | MEDNAME | MATTERNAME | COL_DEG | EXECUTE_FACE | STD | PRICE | AMT | MEMO | SORTGBN | MEDNAME2"
		
		mobjSCGLSpr.SetHeader .sprSht,		  "게재일|광고주|매체명|소재명|색도|집행면|사이즈|단가|금액|비고|SORTGBN|MEDNAME2"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   10|    15|    15|    18|   5|    11|    10|  10|  12|  15|      0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRICE | AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME | MEDNAME | MATTERNAME | COL_DEG | EXECUTE_FACE | STD | MEMO | SORTGBN | MEDNAME2", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PUB_DATE | CLIENTNAME | MEDNAME | MATTERNAME | COL_DEG | EXECUTE_FACE | STD | PRICE | AMT | MEMO | SORTGBN | MEDNAME2"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "COL_DEG|STD",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "SORTGBN | MEDNAME2 ", true
		mobjSCGLSpr.CellGroupingEach .sprSht, "PUB_DATE | CLIENTNAME | MEDNAME"
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjEXECUTE = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		DateClean Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtFPUB_DATE.focus()
	End with
End Sub

Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		.txtFPUB_DATE.value = date1
		.txtTPUB_DATE.value = date2
	End With
End Sub

Sub DateClean_Change (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	frmThis.txtTPUB_DATE.value = date2
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjEXECUTE.SelectRtn_OneAndOne(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtFPUB_DATE.value, .txtTPUB_DATE.value, .txtCLIENTCODE.value, .txtMEDCODE.value, .txtTIMCODE.value)

		if not gDoErrorRtn ("SelectRtn_OneAndOne") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If mobjSCGLSpr.GetTextBinding(.sprSht,"STD",intCnt) = "소계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		elseif mobjSCGLSpr.GetTextBinding(.sprSht,"STD",intCnt) = "합계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		End If
	Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
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
													<TABLE cellSpacing="0" cellPadding="0" width="205" background="../../../images/back_p.gIF"
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
												<td class="TITLE">광고비 집행내역 - 일자별 세부검색</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
														height="20" alt="창을 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
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
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFPUB_DATE, txtTPUB_DATE)"
														width="80">게재일</TD>
													<TD class="SEARCHDATA" width="440"><INPUT dataFld="FPUB_DATE" class="INPUT" id="txtFPUB_DATE" title="게재일" style="WIDTH: 96px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtFPUB_DATE">&nbsp;<IMG id="imgCalFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"  align="absMiddle" border="0" name="imgCalFrom">&nbsp;~
														<INPUT dataFld="TPUB_DATE" class="INPUT" id="txtTPUB_DATE" title="게재일" style="WIDTH: 96px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtTPUB_DATE">&nbsp;<IMG id="imgCalTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"  align="absMiddle" border="0" name="imgCalTo">&nbsp; 
														위수탁 거래명세서 발행 기준
													</TD>
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)"
														width="80">광고주</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 168px; HEIGHT: 22px"
															type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
															border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
															type="text" maxLength="6" align="left" name="txtCLIENTCODE">
													</TD>
												</TR>
												<TR>
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDCODE, txtMEDNAME)">매체사</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMEDNAME" title="코드명" style="WIDTH: 173px; HEIGHT: 22px" type="text"
															maxLength="100" align="left" name="txtMEDNAME"> <IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
															align="absMiddle" border="0" name="ImgMEDCODE"> <INPUT class="INPUT_L" id="txtMEDCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px" type="text"
															maxLength="6" align="left" size="5" name="txtMEDCODE">
													</TD>
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME, txtTIMCODE)">팀명</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME" title="팀명" style="WIDTH: 168px; HEIGHT: 22px" type="text"
															maxLength="100" align="left" size="22" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
															align="absMiddle" border="0" name="ImgTIMCODE"> <INPUT class="INPUT_L" id="txtTIMCODE" title="팀코드" style="WIDTH: 64px; HEIGHT: 22px" type="text"
															maxLength="6" align="left" name="txtTIMCODE">
													</TD>
												</TR>
											</TABLE>
										
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="17066">
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
					<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
