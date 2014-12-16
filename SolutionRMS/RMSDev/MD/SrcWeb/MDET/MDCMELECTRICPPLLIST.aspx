<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRICPPLLIST.aspx.vb" Inherits="MD.MDCMELECTRICPPLLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 간접광고 현황 및 정산 계획</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 공중파 간접광고 현황 및 정산 계획
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 거래처 대한 MAIN 정보를 조회/저장/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/12/05 By KTY
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
Dim mobjMDETELECTRICPPLLIST '공통코드, 클래스
Dim mobjMDCOGET
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
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


'--------------------------------------
'하단 버튼 이벤트
'-----------------------------------
 '하단 추가
'----------------------------------- 

SUB ImgConfirmRequest_onclick ()
	With frmThis
		ProcessRtn_CONFIRM
	End With 
END SUB



sub ImgAddRowDTR_onclick ()
	With frmThis
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "상단의 데이터가 없을경우 추가 하실 수 없습니다..","저장안내"
			exit Sub
		End If
		
		call sprSht_DTL_Keydown(meINS_ROW, 0)
		.txtYEARMON.focus
		.sprSht_DTL.focus
	End With 
end sub

'-----------------------------------
' 저장   
'-----------------------------------
Sub imgSaveDTL_onclick ()
	IF frmThis.sprSht_DTL.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'삭제
'-----------------------------------
Sub imgDelete_DTL_onclick ()
	Dim i
	If frmThis.sprSht_DTL.MaxRows = 0 Then
		gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn_DTL
	gFlowWait meWAIT_OFF
End Sub



'-----------------------------------
'인쇄
'-----------------------------------
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
	Dim strYEARMON,strCLIENTCODE,strDEPT_CD,strGUBUN
	
	Dim Con1,Con2,Con3,Con4
	
	with frmThis
		Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" 
		
		if frmThis.sprSht_DTL.MaxRows = 0 then
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
			Exit Sub
		end if
		
		for i = 1 to .sprSht_DTL.maxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",i) > 0 then
				if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CONFIRM_USER",i)  = "" then
					gErrorMsgBox "승인 되지 않은 데이터는 인쇄 하실수 없습니다.!.","인쇄 안내!"
					Exit Sub
				end if
			end if 
		next
		
		ModuleDir = "MD"
		
		ReportName = "MDELECTRICPPL.rpt"

		
		strYEARMON		= mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",.sprSht.ActiveRow)
		strCLIENTCODE	= mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		strDEPT_CD		= mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_CD",.sprSht.ActiveRow)
		strGUBUN		= mobjSCGLSpr.GetTextBinding( .sprSht,"GUBUN",.sprSht.ActiveRow)
		
		If strYEARMON <> ""	Then Con1  = " AND (YEARMON = '" & strYEARMON & "') "
		If strCLIENTCODE <> ""	Then Con2  = " AND (CLIENTCODE = '" & strCLIENTCODE & "') "
		If strDEPT_CD <> ""	Then Con3  = " AND (DEPT_CD = '" & strDEPT_CD & "') "
		If strGUBUN <> ""	Then Con4  = " AND (ISNULL(ATTR01,'') = '" & strGUBUN & "') "

		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & Con4 
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
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
		vntRet = gShowModalWindow("../../../MD/SrcWeb/MDCO/MDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
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
			vntData = mobjMDCOGET.GetMDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetPDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
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

'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTCODE") or Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						
						.txtYEARMON.focus()
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_DTL_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME"), Row
						.sprSht_DTL.focus 
						mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTCODE") or Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)
		
				If not gDoErrorRtn ("Get_EXCLIENT") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"EXCLIENTCODE",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"EXCLIENTNAME",Row, trim(vntData(3,1))
						.txtYEARMON.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_DTL_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME"), Row
						.txtYEARMON.focus
						.sprSht_DTL.focus 
						mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		'----------------------------------------
		'금액이나 횟수 변경시 자동 계산 로직
		'----------------------------------------
		Dim intCNT
		Dim intCNT_AMT
		Dim intEXSUSU
		
		Dim intTOT_CNT
		Dim intCHARGE_CNT
		
		intEXSUSU = 0
		intCNT_AMT = 0
		intCNT = 0
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CNT") Then
			intCNT = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CNT",Row)
			intCNT_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CNT_AMT",Row)
			
			intEXSUSU = intCNT * intCNT_AMT
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"EXSUSU",row, intEXSUSU
			
		end if
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CNT_AMT") Then
			intCNT_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CNT_AMT",Row)
			intCNT = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CNT",Row)
			
			intEXSUSU = intCNT * intCNT_AMT
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"EXSUSU",row, intEXSUSU
		end if 
		
	'총 횟수 변경시 당월횟수는 0 으로 입력유도하고 잔여 횟수는 총횟수로 세팅 
	'일단 횟수나 자동 계산 로직은 뺏음..20111208_ SH
	'	If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"TOT_CNT") Then
	'		intTOT_CNT = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"TOT_CNT",Row)
	'		
	'		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CNT",row, 0
	'		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CHARGE_CNT",row, intTOT_CNT
	'	END IF 
		

	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub

Sub mobjSCGLSpr_DTL_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End IF
		end if
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, vntRet(1,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",Row, vntRet(2,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End IF
		end if
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.txtYEARMON.focus
		.sprSht_DTL.Focus
	end with
End Sub


'-----------------------------------
'쉬트 클릭
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_DTLBinding Col, Row
		End If
	End With
End Sub

'-----------------------------------
'쉬트 더블클릭
'-----------------------------------
sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		end if
	end with
end sub



'--------------------------------------------------
'쉬트 키업
'--------------------------------------------------
Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_DTLBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		
	End If
End Sub


'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1) 'TBRDDAY
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"YEARMON",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CLIENTCODE",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CLIENTNAME",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"DEPT_CD",frmThis.sprSht_DTL.ActiveRow,  mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"DEPT_NAME",frmThis.sprSht_DTL.ActiveRow,  mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow) 
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"TBRDDAY",frmThis.sprSht_DTL.ActiveRow, "월"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"TBRDFDATE",frmThis.sprSht_DTL.ActiveRow, gNowDate
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"TBRDTDATE",frmThis.sprSht_DTL.ActiveRow, DATEADD("D",-1,DATEADD("m",1,MID(gNowDate,1,7) & "-01"))
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"TOT_CNT",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CNT",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CHARGE_CNT",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"PRICE",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"AMT",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"COMMISSION",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"CNT_AMT",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"EXSUSU",frmThis.sprSht_DTL.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"ATTR01",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow)
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht_DTL, 1,frmThis.sprSht_DTL.MaxRows
	End if
End Sub
		'

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_DTL_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXBTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMEXEPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, vntRet(1,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",Row, vntRet(2,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
		End If	
		.sprSht_DTL.Focus
		mobjSCGLSpr.ActiveCell .sprSht_DTL, Col, Row
	End With
End Sub

Sub txtYEARMON_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDETELECTRICPPLLIST = gCreateRemoteObject("cMDET.ccMDETELECTRICPPLLIST")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjSCCOGET				= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    
    With frmThis
		'상위 가상간접광고 내역 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht	
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTCODE | CLIENTNAME | PROGRAM | DEPT_CD | DEPT_NAME | BILLING | PPLAMT | PPLSUSU | GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		  "년월|광고주코드|광고주명|프로그램|부서코드|담당부서명|취급액|PPL인정금액|PPL인정수수료|구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|         0|      25|      10|       0|        20|    12|         12|           12|   8"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "BILLING | PPLAMT | PPLSUSU", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, " YEARMON | CLIENTCODE | CLIENTNAME | PROGRAM | DEPT_CD | DEPT_NAME | GUBUN", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTCODE | CLIENTNAME | PROGRAM | DEPT_CD | DEPT_NAME | BILLING | PPLAMT | PPLSUSU | GUBUN"
		mobjSCGLSpr.colhidden .sprSht, "CLIENTCODE | DEPT_CD",true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN" ,-1,-1,2,2,false
		'mobjSCGLSpr.CellGroupingEach .sprSht,"BILLING | PPLAMT | PPLSUSU"
		
		'하위 가상간접광고 추가 입력 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 29, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 4, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 20, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | YEARMON | SEQ | CLIENTCODE | BTN | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXBTN |  EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO | ATTR01 "
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "선택|년월|순번|광고주코드|광고주명|담당부서코드|담당부서명|채널|프로그램|요일|청약방송시작일|청약방송종료일|총횟수|당월횟수|잔여횟수|매체비단가|월총매체비|월총수수료|파트너코드|파트너명|파트너회당매체수익|월총매체청구비|전표번호|승인자|승일일|비고|구분"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 4|   0|   4|         8|2|    18|           0|        10|	8|      15|   5|            12|            12|     7|       7|       7|        10|        10|        10|         8|2|    15|                15|            13|       0|     8|     0|  20|   0"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_DTL, "TBRDDAY", -1, -1, "월" & vbTab & "화" & vbTab & "수" & vbTab & "목" & vbTab & "금" & vbTab & "토" & vbTab & "일"  , 10, 40, False, False
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_DTL,"..", "BTN | EXBTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | CNT_AMT | EXSUSU ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "TBRDFDATE | TBRDTDATE | CONFIRM_DATE ", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, " CLIENTCODE | CLIENTNAME | DEPT_CD | MEDNAME | PROGRAM | EXCLIENTCODE | EXCLIENTNAME | VOCHNO | CONFIRM_USER ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, True, " DEPT_NAME | VOCHNO | CONFIRM_USER | CONFIRM_DATE"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "YEARMON | DEPT_CD | VOCHNO | CONFIRM_DATE | ATTR01", True
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CHK | YEARMON | SEQ | PROGRAM | CONFIRM_USER",-1,-1,2,2,False  '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "MEMO",-1,-1,0,2,false
		
		.sprSht.style.visibility = "visible"	
		.sprSht_DTL.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjMDETELECTRICPPLLIST = Nothing
	set mobjMDCOGET = Nothing
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
	
		.sprSht.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
	
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	End With
End Sub

'------------------------------------------
' HDR 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strYEARMON

	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'변수 초기화
		strYEARMON = ""
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON	= .txtYEARMON.value
		
		vntData = mobjMDETELECTRICPPLLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON)

		If not gDoErrorRtn ("SelectRtn") Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
   			
   			Call SelectRtn_DTLBinding(1,1)
   		End if
   	End With
End Sub

'------------------------------------------
' DTL 데이터 조회
'------------------------------------------
Sub SelectRtn_DTLBinding(ByVal Col, ByVal Row)
	Dim strYEARMON, strCLIENTCODE, strDEPT_CD, strGUBUN
	Dim vntData
	Dim i, strCols
	Dim strRows
	Dim intCnt, intCnt2
	
	With frmThis
		'sprSht_DTL초기화
		intCnt2 = 1
		.sprSht_DTL.MaxRows = 0
		
		If mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",Row) <> "" Then
			strYEARMON = "" : strCLIENTCODE = "" : strDEPT_CD = ""
		
			strYEARMON		= mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",Row)
			strCLIENTCODE	= mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row)
			strDEPT_CD		= mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_CD",Row)
			strGUBUN		= mobjSCGLSpr.GetTextBinding( .sprSht,"GUBUN",Row)
				
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			vntData = mobjMDETELECTRICPPLLIST.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strDEPT_CD, strGUBUN)

			If not gDoErrorRtn ("SelectRtn_DTL") Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				'전표번호가 있으면 락을 건다
				FOR intCnt = 1 TO .sprSht_DTL.MaxRows
   					 
   					'SEQ 가 0인건은 TOTAL 라인 이다. TOTAL 라인은 셀을 합치고 색을 변경하며 CHK 박스를 제거한다..
   					IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",intCnt) = 0 THEN
   						mobjSCGLSpr.AddCellSpan .sprSht_DTL,2, intCnt, 21,1,false
   						mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "CHK",intCnt,intCnt,4,,,,,False
   						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CHK",intCnt, " "
   						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
   						mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,intCnt,1,-1,true
   					END IF 
   					
   					'승인 요청된 건은 락을 걸고 색을 변경한다.
   					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CONFIRM_USER",intCnt) <> "" then
   						mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,intCnt,1,-1,true
   						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HAAE8EE, &H000000,False
   																					  
   					END IF 
   					'전표가 생성된 건은 색을 변경하고 락을 건다.
   					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"VOCHNO",intCnt) <> "" then
   						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HEFE9EA, &H000000,False
   						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
   					end if
   					
   				next 
   				mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,True,strRows,2,26,True
			End if	
		
   			gWriteText lblStatusDTR, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
   		End If
	End With
End Sub

'------------------------------------------
' DTL 데이터 저장
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
   	Dim strRow
	Dim lngCol, lngRow
	Dim strDataCHK
	With frmThis
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_DTL, "CLIENTCODE | CLIENTNAME | EXCLIENTCODE | EXCLIENTNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 광고주명/파트너명 은 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | YEARMON | SEQ | CLIENTCODE | BTN | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXBTN |  EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO | ATTR01 ")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If
		
		intRtn = mobjMDETELECTRICPPLLIST.ProcessRtn(gstrConfigXml,vntData)
	
		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			SelectRtn_DTLBinding .sprsht.ActiveCol, .sprsht.ActiveRow
			mobjSCGLSpr.ActiveCell .sprSht_DTL, 1, strRow
   		End If
   	End With
End Sub

'------------------------------------------
'데이터 삭제 디테일
'------------------------------------------
Sub DeleteRtn_DTL ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '실제데이터여부 플레
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		
		for i = 1 to .sprSht_DTL.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) <> " " Then
				If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 Then
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"VOCHNO",i) <> "" Then
						gErrorMsgBox "선택하신 " & i & "행의 자료는 전표가 생성된 내역입니다..삭제 하실 수 없습니다.","삭제안내!"
						exit Sub
					else 
						If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CONFIRM_USER",i) <> "" Then
							gErrorMsgBox "선택하신 " & i & "행의 자료는 승인된 자료입니다." & vbcrlf & "먼저 승인취소처리 하십시오!","삭제안내!"
							exit Sub
						End If
						lngchkCnt = lngchkCnt +1
					End If
				End If
			end if
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_DTL.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) <> " " Then
				If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 Then
					dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",i)
					strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"YEARMON",i)
					
					If dblSEQ = "" Then
						mobjSCGLSpr.DeleteRow .sprSht_DTL,i
					else
						intRtn = mobjMDETELECTRICPPLLIST.DeleteRtn(gstrConfigXml,dblSEQ, strYEARMON)
						
						If not gDoErrorRtn ("DeleteRtn") Then
							mobjSCGLSpr.DeleteRow .sprSht_DTL,i
   						End If
   						strSEQFLAG = True
					End If				
   					intCnt = intCnt + 1
   				End If
   			end if
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_DTL
		'저장이 되어있던 데이터 삭제시 재조회 단순 추가후 삭제는 로우만 삭제
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
	err.clear	
End Sub


'------------------------------------------
'하단 데이터 승인 요청
'------------------------------------------
Sub ProcessRtn_confirm ()
    Dim intRtn
   	Dim vntData
   	Dim i , intCnt, intCnt2, intchkCnt
   	Dim intSaveRtn
   	Dim strYEARMON , strSEQ
   	
   	'SMS 정보
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim vntData_info
	Dim strMsg ,  strMstMsg
   	
	With frmThis
		intCnt = 0

		for i = 1 to  .sprSht_DTL.maxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CHK",i) = "1" then
				intCnt = intCnt + 1
			end if
			
			IF mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"SEQ",i) = "" THEN
				gErrorMsgBox "저장되지 않은 데이터는 승인요청 하실 수 없습니다.","승인요청안내"
				exit sub
			END IF 
		next
		
		if intCnt = 0 then 
			gErrorMsgBox "체크 된 데이터가 없습니다.","승인요청안내"
			exit sub
		end if
		
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "승인권자를 선택 하십시오.","승인요청안내"
			Exit Sub
		End If
		
		'승인권자 를 그리드에 탑재
		For intCnt2 = 1 To .sprSht_DTL.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CONFIRM_USER",intCnt2,Trim(.txtEMPNO.value)
			'그리드의 프로그램명 을 가져온다
			If intCnt2 = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PROGRAM",intCnt2)
			End If
		Next
		
		If intCnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] 승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 승인요청이있습니다"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] 외" & intCnt-1 & "건의승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 외" & intCnt-1 & "건의승인요청이있습니다"
			End If
		End If
		
		if DataValidation =false then exit sub 	
		
		intSaveRtn = gYesNoMsgbox("해당데이터를 승인요청 하시겠습니까?","승인요청 확인")
		IF intSaveRtn <> vbYes then 
			exit Sub
		END IF
		
		For intchkCnt = 1 To .sprSht_DTL.MaxRows
				strYEARMON = "" : strSEQ  = ""
			if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CHK",intchkCnt) = "1" then
				strYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"YEARMON",intchkCnt)
				strSEQ	   = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"SEQ",intchkCnt)
				
				intRtn = mobjMDETELECTRICPPLLIST.ProcessRtn_confirm(gstrConfigXml, strYEARMON, strSEQ, .txtEMPNO.value)
			end if
		Next

		If not gDoErrorRtn ("ProcessRtn_confirm") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 승인요청 되었습니다.!"," 저장안내!"
			
			'승인을 수락하였으므로 SMS 발송
			'보내는 사람의 정보 가져오기
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
			
			'보내는사람정보
			strFromUserName		= vntData_info(0,2)
			strFromUserEmail	= vntData_info(1,2)
			strFromUserPhone	= vntData_info(2,2)
			
			'받는사람 정보
			strToUserName		=  vntData_info(0,1)
			strToUserEmail		=  vntData_info(1,1)
			strToUserPhone		=  vntData_info(2,1)

			call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
	
			SelectRtn_DTLBinding .sprsht.ActiveCol, .sprsht.ActiveRow
   		End If
   	End With
End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
   	Dim intCnt
	'On error resume next
	with frmThis
   		for intCnt = 1 to .sprSht_DTL.MaxRows
   			'Sheet 필수 입력사항
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CONFIRM_USER",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 행의 승인권자 입력에 문제가 있습니다" & vbcrlf & "운영팀 에게 문의 하십시오.","승인요청안내"
				Exit Function
			End if
		next
   	End with
   	
	DataValidation = true
End Function

-->
		</script>
		<script language="javascript">
		//SMS 발송
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){	
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
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
												<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
											<td class="TITLE">공중파 간접광고 현황 및 정산 계획</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
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
											<TD class="SEARCHLABEL" width="60">년월</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="년월조회" style="WIDTH: 112px; HEIGHT: 22px" maxLength="6"
													size="13" name="txtYEARMON" accessKey="NUM"></TD>
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
											<TD vAlign="middle" align="right" height="20"></TD>
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
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="6959">
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
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="23" height="20" style="WIDTH: 23px"></TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td style="FONT-WEIGHT: bold; FONT-SIZE: 12px" align="right" width="600">
															<INPUT class="NOINPUTB_R" id="txtCOLORCONTRACT" title="전표생성됨" style="WIDTH: 20px; HEIGHT: 22px; BACKGROUND-COLOR: #eae9ef"
																accessKey="NUM" readOnly maxLength="100" size="13" name="txtCOLORCONTRACT"> 
															전표생성 <INPUT class="NOINPUTB_R" id="txtCOLORGUESS" title="승인요청&amp;승인" style="WIDTH: 20px; HEIGHT: 22px; BACKGROUND-COLOR: #eee8aa"
																accessKey="NUM" readOnly maxLength="100" size="13" name="txtCOLORGUESS"> 승인요청&amp;승인&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
															<span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">
																승인자:</span> <INPUT class="NOINPUTB_L" id="txtEMPNAME" title="승인권자" style="WIDTH: 96px; HEIGHT: 20px"
																maxLength="100" size="10" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
																border="0" name="ImgEMPNO" title="승인권자선택"> <INPUT class="NOINPUTB" id="txtEMPNO" title="승인권자사번" style="WIDTH: 58px; HEIGHT: 20px"
																maxLength="100" size="4" name="txtEMPNO"></td>
														<td><IMG id="ImgConfirmRequest" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'"
																height="20" alt="선택하신 데이터를 승인 요청 합니다." src="../../../images/ImgConfirmRequest.gIF"
																border="0" name="ImgConfirmRequest"></td>
														<TD><IMG id="ImgAddRowDTR" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRowDTR"></TD>
														<TD><IMG id="imgSaveDTL" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'"
																height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" border="0" name="imgSaveDTL"></TD>
														<TD><IMG id="imgDelete_DTL" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete_DTL"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
														<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="6985">
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
								<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
			</TABLE>
		</FORM>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 0px;HEIGHT: 0px" name="frmSMS"></iframe> <!--DISPLAY: none; -->
	</body>
</HTML>
