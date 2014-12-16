<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPROJECTIONLIST.aspx.vb" Inherits="MD.MDCMPROJECTIONLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>월별 AGENCY 공중파 광고정산</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/그룹광고 분담금 입력/조회 화면(MDCMGROUP)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMGROUP.aspx.aspx
'기      능 : 그룹광고 분담금 을 조회/입력 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/01/09 By Kim Tae Yub
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
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET, mobjMDPTPROJECTIONLIST'공통코드, 클래스
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

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
sub imgNewReg_onclick ()
	if frmThis.txtCLIENTCODE1.value = "" OR frmThis.txtCLIENTNAME1.value = ""   then
		gErrorMsgBox "광고주를 입력하시오","추가안내"
		frmThis.txtCLIENTNAME1.focus()
		exit Sub
	end if
	
	if TRIM(mobjSCGLSpr.GetTextBinding( frmThis.sprSht_SUM,"AMT", 2)) = "0"   then
		gErrorMsgBox "돌출 광고비가 0입니다. 추가할 수 없습니다.","추가안내"
		frmThis.txtCLIENTNAME1.focus()
		exit Sub
	end if
	
	With frmThis
		Call sprSht_Keydown(meINS_ROW, 0)
	End With 
End sub

Sub imgQuery_onclick
	if (frmThis.txtFYEARMON.value = "" AND frmThis.txtTYEARMON.value = "") or frmThis.txtCLIENTCODE1.value = ""   then
		ImgCLIENTCODE1_onclick
		'gErrorMsgBox "년월 또는 광고주를 입력하시오","조회안내"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	if frmThis.txtCLIENTCODE1.value = "" OR frmThis.txtCLIENTNAME1.value = ""   then
		gErrorMsgBox "광고주를 입력하시오","조회안내"
		frmThis.txtCLIENTNAME1.focus()
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick
	if frmThis.txtCLIENTCODE1.value = "" OR frmThis.txtCLIENTNAME1.value = ""   then
		gErrorMsgBox "광고주를 입력하시오","조회안내"
		frmThis.txtCLIENTNAME1.focus()
		exit Sub
	end if
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_EXCEL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
	Dim strYEARMON
	Dim strCLIENTNAME
	Dim strCLIENTCODE
	dim chkflag
   	dim strLIST
   	Dim strClientLIST
   	Dim intSUBRow
	
	with frmThis
		
		strLIST = ""
		chkflag = 1
				gErrorMsgBox "출력물은 개발중입니다..",""
			Exit Sub
			
'		if frmThis.sprSht.MaxRows = 0 then
'			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
'			Exit Sub
'		end if
		
'		strClientLIST = split(mstrClientcode, "|")
		
'		intSUBRow = UBound(strClientLIST, 1)
'		FOR i = 0 to intSUBRow
'			IF chkflag = 1 then
'				strLIST = "'" & strClientLIST(i) & "'"
'				chkflag = 2
'			else
'				strLIST = strLIST & ",'" & strClientLIST(i) & "'"
'			end if 
'		Next
		
'		ModuleDir = "MD"
'		ReportName = "MDCMCLIENTSUBSEQMEDLIST.rpt"
'		
'		strYEARMON		= .txtYEARMON.value
'		strCLIENTNAME	= .txtCLIENTNAME.value
'		
'		Params = strYEARMON & ":" & strLIST & ":" & strCLIENTNAME
'		
'		Opt = "A"
'		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

Sub txtFYEARMON_onchange
	with frmThis
		.txtTYEARMON.value = .txtFYEARMON.value
	end With
End Sub

'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	On error resume next
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------
' SpreadSheet sprSht 이벤트
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,false,frmThis.sprSht.ActiveRow,3,3,true
		
		frmThis.txtFYEARMON.focus
		frmThis.sprSht.focus
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 스프레드 버튼클릭 이벤트
'-----------------------------------------------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUBSEQ") Then '브랜드
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)), _
								TRIM(.txtCLIENTCODE1.value), TRIM(.txtCLIENTNAME1.value))
								
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_TIMCODE.aspx",vntInParams , 640,445)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col+1,Row
			End If
			.txtFYEARMON.focus
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		ElseIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNEX") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			.txtFYEARMON.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	Dim intColor
   	intColor = ""
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON") Then
			IF Row > 1 THEN
				IF mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",1) <> "" THEN
					IF mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",Row) <> mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",1) THEN
						gErrorMsgBox "첫행과 같은 년월을 입력하세요.","조회안내"
						mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",Row, ""
						exit Sub
					END IF
				END IF
			END IF
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_BrandInfo_TIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
								"",TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)),  _
								TRIM(.txtCLIENTCODE1.value), TRIM(.txtCLIENTNAME1.value))

				If not gDoErrorRtn ("Get_BrandInfo_TIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntData(1,1)
						
						.txtFYEARMON.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME"), Row
						.txtFYEARMON.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(2,1)			
						.txtFYEARMON.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME"), Row
						.txtFYEARMON.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet, vntInParams
	Dim strGUBUN
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)), _
								TRIM(.txtCLIENTCODE1.value), TRIM(.txtCLIENTNAME1.value))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,445)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
					
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtFYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtFYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
		
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.txtFYEARMON.focus
		.sprSht.Focus
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
	set mobjMDPTPROJECTIONLIST	= gCreateRemoteObject("cMDPT.ccMDPTPROJECTIONLIST")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 8, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | MATTERNAME | SUBSEQ | BTNSUBSEQ | SUBSEQNAME | EXCLIENTCODE | BTNEX | EXCLIENTNAME | AMT"
		mobjSCGLSpr.SetHeader .sprSht,        "선택|년월|번호|소재명|코드|브랜드|코드|제작사|집행누계"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  4|   9|   0|    25|   7|2|  15|   7|2|  15|      13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTNSUBSEQ | BTNEX"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | SEQ | MATTERNAME | SUBSEQ | SUBSEQNAME | EXCLIENTCODE | EXCLIENTNAME", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1,0
		mobjSCGLSpr.ColHidden .sprSht, "SEQ", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,false		
		
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 6, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "GBN | EXCLIENTCODE | EXCLIENTNAME | AMT | GBN2 | SAVEFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUM,       "구분|제작사코드|제작사|금액/비율|원/%|저장플레그"
		mobjSCGLSpr.SetColWidth .sprSht_SUM, "-1", "20|         0|    20|       15|  10|        0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUM, "GBN | EXCLIENTCODE | EXCLIENTNAME | GBN2", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT", -1, -1,0
		mobjSCGLSpr.ColHidden .sprSht_SUM, "SAVEFLAG", True
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUM, true, "GBN | EXCLIENTCODE | EXCLIENTNAME | GBN2"
		mobjSCGLSpr.CellGroupingEach .sprSht_SUM, "GBN"
		
		
		gSetSheetColor mobjSCGLSpr, .sprSht_EXCEL
		mobjSCGLSpr.SpreadLayout .sprSht_EXCEL, 5, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_EXCEL, "YEARMON | MATTERNAME | SUBSEQNAME | EXCLIENTNAME | AMT"
		mobjSCGLSpr.SetHeader .sprSht_EXCEL,       "년월|소재명|브랜드|제작사|집행누계"
		mobjSCGLSpr.SetColWidth .sprSht_EXCEL, "-1", "10|    25|    20|    15|     15"
		mobjSCGLSpr.SetRowHeight .sprSht_EXCEL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_EXCEL, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EXCEL, "YEARMON | MATTERNAME | SUBSEQNAME | EXCLIENTNAME | AMT", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EXCEL, "AMT", -1, -1,0
		mobjSCGLSpr.CellGroupingEach .sprSht_EXCEL, "YEARMON"
		
    End With

	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDPTPROJECTIONLIST = Nothing
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
		.txtFYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		'.txtFYEARMON.value = "200912"
		.txtTYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.sprSht_SUM.MaxRows = 0
		
		
		.txtFYEARMON.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	dim chkflag
   	dim strSUBLIST
   	Dim strCLIENTSUBLIST
   	Dim intSUBRow
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDPTPROJECTIONLIST.SelectRtn_PROJECTIONLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtFYEARMON.value, .txtTYEARMON.value, .txtCLIENTCODE1.value)

		if not gDoErrorRtn ("SelectRtn_PROJECTIONLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			
   			CALL SelectRtn_SUM ()
   			CALL SelectRtn_EXCEL ()
   		end if
   	end with
End Sub

Sub SelectRtn_SUM ()
	Dim vntData
   	Dim i, strCols
   	Dim strRows, strRows2
	Dim intCnt, intCnt2, intCnt3
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_SUM.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intCnt2 = 1
		intCnt3 = 1
		
		vntData = mobjMDPTPROJECTIONLIST.SelectRtn_SUM(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtFYEARMON.value, .txtTYEARMON.value, .txtCLIENTCODE1.value)

		if not gDoErrorRtn ("SelectRtn_PROJECTIONLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht_SUM, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			For intCnt = 1 To .sprSht_SUM.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"SAVEFLAG",intCnt) <> "Y" Then
					If intCnt2 = 1 Then
						strRows = intCnt
					Else
						strRows = strRows & "|" & intCnt
					End If
					intCnt2 = intCnt2 + 1
				End If
				IF mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",intCnt) = "%" Then
					If intCnt3 = 1 Then
						strRows2 = intCnt
					Else
						strRows2 = strRows2 & "|" & intCnt
					End If
					intCnt3 = intCnt3 + 1
				END IF
			Next
			
			mobjSCGLSpr.SetCellsLock2 .sprSht_SUM,true,strRows,4,4,true
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT", -1, -1,0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, strRows2, 4, 4, 2,"-99999999999999.99","99999999999999.99",FALSE, TRUE, "\",1,2,TRUE
   		end if
   		Layout_change
   	end with
End Sub

Sub SelectRtn_EXCEL ()
	Dim vntData
   	Dim i, strCols
   	Dim intRtn
   	Dim intSprshtcnt, intSprsht_Sumcnt, intSumcnt
   	Dim intCnt, intCnt2, intCnt3
   	Dim strRows, strRows2
	'On error resume next
	with frmThis
		.sprSht_EXCEL.MaxRows = 0
		
		intSprshtcnt = .sprSht.MaxRows
		intSprsht_Sumcnt = .sprSht_SUM.MaxRows
		
		intSumcnt = intSprshtcnt + intSprsht_Sumcnt
		
		for i=1 to intSprshtcnt
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_EXCEL, meINS_ROW, 0, -1, 1)
			
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"YEARMON",		.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"MATTERNAME",	.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"SUBSEQNAME",	.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"EXCLIENTNAME",.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"AMT",			.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",i)
		Next
		
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht_EXCEL, meINS_ROW, 0, -1, 1)
		intCnt2 = 1
		intCnt3 = 1
		
		for i=1 to intSprsht_Sumcnt
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_EXCEL, meINS_ROW, 0, -1, 1)
			
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"YEARMON",		.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"MATTERNAME",	.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"EXCLIENTNAME",i)
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",i) = "" THEN
				mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"SUBSEQNAME",	.sprSht_EXCEL.ActiveRow, ""
			ELSE
				mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"SUBSEQNAME",	.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"AMT",i)
			END IF
			
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"EXCLIENTNAME",.sprSht_EXCEL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",i)
			mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"AMT",			.sprSht_EXCEL.ActiveRow, ""
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",i) = "원" Then
				If intCnt2 = 1 Then
					strRows = intSprshtcnt + 1 + i
				Else
					strRows = strRows & "|" & intSprshtcnt + 1 + i
				End If
				intCnt2 = intCnt2 + 1
			END IF
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",i) = "%" Then
				If intCnt3 = 1 Then
					strRows2 = intSprshtcnt + 1 + i
				Else
					strRows2 = strRows2 & "|" & intSprshtcnt + 1 + i
				End If
				intCnt3 = intCnt3 + 1
			END IF
		Next
		
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EXCEL, "SUBSEQNAME", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EXCEL, strRows, 3, 3, 0,"-99999999999999","99999999999999",FALSE, TRUE, "\",1,2,TRUE
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EXCEL, strRows2, 3, 3, 2,"-99999999999999.99","99999999999999.99",FALSE, TRUE, "\",1,2,TRUE
		
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	With frmThis
		For intCnt = 1 To .sprSht_SUM.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"GBN2",intCnt) = "" Then
				mobjSCGLSpr.SetCellShadow .sprSht_SUM, -1, -1, intCnt, intCnt,&HD3FED7, &H000000,False
			Else
				mobjSCGLSpr.SetCellShadow .sprSht_SUM, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
			End If
		Next 
	End With
End Sub


'****************************************************************************************
' 저장로직
'****************************************************************************************
Sub ProcessRtn()
	Dim intRtn, intRtn2
	Dim lngCol, lngRow
   	Dim vntData, vntData2, vntData_Src
   	Dim strDataCHK
   	Dim intCnt
   	
	With frmThis
		intRtn2 = gYesNoMsgbox("입력한 년월/소재명/브랜드/제작사를 확인하셨습니까? " & VBCRLF & " 년월,소재명입력없는 로우는 자동 삭제됩니다.","저장안내")
		If intRtn2 <> vbYes Then exit Sub
		
		For intCnt = 1 to .sprSht.MaxRows
			If Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intCnt)) = "" OR Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",intCnt)) = ""  Then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End If
		Next
	
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "YEARMON | MATTERNAME | SUBSEQ | EXCLIENTCODE",lngCol, lngRow, False) 
		
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 년월/소재명/브랜드/제작사는 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If
		 
		 mobjSCGLSpr.SetFlag .sprSht, meINS_TRANS
		 mobjSCGLSpr.SetFlag .sprSht_SUM, meINS_TRANS
		 
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | MATTERNAME | SUBSEQ | BTNSUBSEQ | SUBSEQNAME | EXCLIENTCODE | BTNEX | EXCLIENTNAME | AMT")
		
		vntData2 = mobjSCGLSpr.GetDataRows(.sprSht_SUM,"GBN | EXCLIENTCODE | EXCLIENTNAME | AMT | GBN2 | SAVEFLAG")
		
		If Not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"	
			Exit Sub 
		End If
		
		'처리 업무객체 호출
		intRtn = mobjMDPTPROJECTIONLIST.ProcessRtn(gstrConfigXml,vntData, vntData2, .txtCLIENTCODE1.value)
		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox intRtn & " 건 이 저장되었습니다.","저장안내"
			SelectRtn
   		End If
   	End With
End Sub


'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
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
				mobjSCGLSpr.DeleteRow .sprSht,i
   				intCnt = intCnt + 1
   			End If
		Next
		
		call ProcessRtn()
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
	End With
	err.clear	
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
												<TABLE cellSpacing="0" cellPadding="0" width="90" background="../../../images/back_p.gIF"
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
											<td class="TITLE">인쇄 돌출 조회&nbsp;</td>
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
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table Start-->
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
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="년월을삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFYEARMON,txtTYEARMON)">년&nbsp;&nbsp;월
											</TD>
											<TD class="SEARCHDATA" width="250"><INPUT class="INPUT" id="txtFYEARMON" title="년도을입력하세요" style="WIDTH: 100px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="14" name="txtFYEARMON"> ~ <INPUT class="INPUT" id="txtTYEARMON" title="년도을입력하세요" style="WIDTH: 100px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="14" name="txtTYEARMON">
											</TD>
											<TD class="SEARCHLABEL" title="광고주를삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1,txtCLIENTCODE1)">광고주
											</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 192px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="26" name="txtCLIENTNAME1">&nbsp;<IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">&nbsp;<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1" size="5">
											</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<td><IMG id="imgNewReg" onmouseover="JavaScript:this.src='../../../images/imgNewRegOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgNewReg.gif'"
																height="20" alt="신규자료를 등록합니다." src="../../../images/imgNewReg.gIF" border="0" name="imgNewReg"></td>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
																name="imgQuery"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="8414">
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
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_SUM" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht_SUM">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="8414">
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
										<OBJECT id="sprSht_EXCEL" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 0%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht_EXCEL">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="0">
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
											<PARAM NAME="ReDraw" VALUE="-1">
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
							<!--List End-->
							<!--Bottom Split Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
								</TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
			</TABLE>
		</FORM>
		</TR></TABLE>
	</body>
</HTML>
