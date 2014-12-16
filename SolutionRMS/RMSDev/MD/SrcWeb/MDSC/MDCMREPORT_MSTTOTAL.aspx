<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMREPORT_MSTTOTAL.aspx.vb" Inherits="MD.MDCMREPORT_MSTTOTAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Comm.BU 광고 매출 실적 및 비용 집계</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987"> <!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET"> <!-- 공통으로 사용될 클라이언트 스크립트를 Include--> <!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->  <!-- UI 공통 ActiveX COM --> <!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->  <!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET,mobjMDSCREPORT_MST'공통코드, 클래스
Dim mClientsubcode

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub


'===================================
' 이벤트 프로시져 
'===================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'저장 버튼 숨기기
Sub Set_MR(byVal strmode)
	With frmThis
		IF .rdMR.checked = TRUE then 
			document.getElementById("imgSave").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgSave").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub    
    
    
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			
			if .txtYEARMON.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			end if
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
					
					if .txtYEARMON.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					end if
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub rdA_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdB_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdAO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdOU_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdMC_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub rdMR_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************

Sub sprSht_Change(ByVal Col, ByVal Row)
   	Dim intsil,intBP,intCHA
	
	With frmThis
		if mobjSCGLSpr.GetTextBinding(.sprSht, "AMTGBN", Row) = "경영계획" then
			intsil= mobjSCGLSpr.GetTextBinding(.sprSht, Col, Row-1)
			intCHA = intsil - mobjSCGLSpr.GetTextBinding(.sprSht, Col ,Row)
			IF mobjSCGLSpr.GetTextBinding(.sprSht, "AMTGBN", Row+1) = "차이" THEN
				mobjSCGLSpr.SetTextBinding .sprSht, Col, Row+1, intCHA  
			END IF
			
			mobjSCGLSpr.SetTextBinding .sprSht, "SUMAMT", Row, mobjSCGLSpr.GetTextBinding(.sprSht, "A" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "B" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "O" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "D" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "R" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "P" ,Row) + _
															   mobjSCGLSpr.GetTextBinding(.sprSht, "E" ,Row)
															   
			mobjSCGLSpr.SetTextBinding .sprSht, "SUMAMT", Row+1, mobjSCGLSpr.GetTextBinding(.sprSht, "A" ,Row+1) + _
															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "B" ,Row+1) + _
															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "O" ,Row+1) + _
															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "D" ,Row+1) + _
															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "R" ,Row+1) + _
															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "P" ,Row+1) + _
															  	 mobjSCGLSpr.GetTextBinding(.sprSht, "E" ,Row+1)

'추후 차이를 수정해서 경영계획을 값을 계산하자고 그럴때 주석풀면됨			
'		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht, "AMTGBN", Row) = "차이" then
'			intsil= mobjSCGLSpr.GetTextBinding(.sprSht, Col, Row-2)
'			intCHA = intsil - mobjSCGLSpr.GetTextBinding(.sprSht, Col ,Row)
'			IF mobjSCGLSpr.GetTextBinding(.sprSht, "AMTGBN", Row-1) = "경영계획" THEN
'				mobjSCGLSpr.SetTextBinding .sprSht, Col, Row-1, intCHA  
'			END IF
'			
'			mobjSCGLSpr.SetTextBinding .sprSht, "SUMAMT", Row, mobjSCGLSpr.GetTextBinding(.sprSht, "A" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "B" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "O" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "D" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "R" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "P" ,Row) + _
'															   mobjSCGLSpr.GetTextBinding(.sprSht, "E" ,Row)
'															   
'			mobjSCGLSpr.SetTextBinding .sprSht, "SUMAMT", Row-1, mobjSCGLSpr.GetTextBinding(.sprSht, "A" ,Row-1) + _
'															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "B" ,Row-1) + _
'															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "O" ,Row-1) + _
'															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "D" ,Row-1) + _
'															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "R" ,Row-1) + _
'															 	 mobjSCGLSpr.GetTextBinding(.sprSht, "P" ,Row-1) + _
'															  	 mobjSCGLSpr.GetTextBinding(.sprSht, "E" ,Row-1)
		END IF 
	END WITH
END SUB

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDSCREPORT_MST	= gCreateRemoteObject("cMDSC.ccMDSCREPORT_MST")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GUBUN", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "GUBUN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN",-1,-1,2,2,false
		
		.sprSht.style.visibility = "visible"
    End With
		
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSCREPORT_MST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		
		'Sheet초기화
		.sprSht.MaxRows = 0
		.rdA.checked = TRUE
		'.txtCLIENTNAME1.focus()
		
		SetChangeLayout
	End with	
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
												'  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   												'1|
		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GUBUN", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "GUBUN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN",-1,-1,2,2,false
	End With
End Sub

Sub SetChangeLayout () 
	With frmThis
		gInitComParams mobjSCGLCtl,"MC"
		mobjSCGLCtl.DoEventQueue
		gSetSheetDefaultColor()
		
		Call Grid_init()
		'년월 광고주코드 광고주명  공중파	CATV	인쇄	옥외	온라인Ⅰ	온라인Ⅱ	프로모션	S.C.수익	제작	GBS/사보	 계 
		if .rdA.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTCODE | CLIENTNAME | A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT"
			mobjSCGLSpr.SetHeader .sprSht,        "년월|광고주코드|광고주명|공중파|CATV|인쇄|옥외|온라인Ⅰ|온라인Ⅱ|프로모션|S.C.수익|제작|GBS/사보|총 계"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  9|         6|      15|    10|       15|  10|      10|      10|       7|      10|  10|      10|10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTCODE | CLIENTNAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT", -1, -1,0
			'mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1,2
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTCODE | CLIENTNAME | A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,false 
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON"
		elseif .rdB.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTCODE | CLIENTNAME | A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT"
			mobjSCGLSpr.SetHeader .sprSht,        "년월|광고주코드|광고주명|공중파|CATV|인쇄|옥외|온라인Ⅰ|온라인Ⅱ|프로모션|S.C.수익|제작|GBS/사보|총 계"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  9|         6|      15|    10|       15|  10|      10|      10|       7|      10|  10|      10|10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTCODE | CLIENTNAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTCODE | CLIENTNAME | A | A2 | B | D | O1 | O2 | R | S | P | E | SUMAMT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,false 
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON"
		elseif .rdAO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON"
			mobjSCGLSpr.SetHeader .sprSht,		 ""
													'  1|
			mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   													'1|
			
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON", -1, -1, 20
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,false
		elseif .rdOU.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | O2 | R | S | P | E | SUMOUTAMT"
			mobjSCGLSpr.SetHeader .sprSht,        "년월|광고주명|온라인|프로모션|S.C.수익|제작|GBS/사보|총 계"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  9|      15|    10|      10|       7|  10|      10|   10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "O2 | R | S | P | E | SUMOUTAMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | O2 | R | S | P | E | SUMOUTAMT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,false 
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON"
		elseif .rdMC.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "CLIENTCODE | CLIENTNAME | A1 | A2 | A3 | A4 | A5 | A6 | A7 | A8 | A9 | A10 | A11 | A12 | AMTSUM"
			mobjSCGLSpr.SetHeader .sprSht,        "광고주코드|광고주명|1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월|합계"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "        0|      12| 12| 12| 12| 12| 12| 12| 12| 12| 12|  12|  12|  12|  15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME ", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " A1 | A2 | A3 | A4 | A5 | A6 | A7 | A8 | A9 | A10 | A11 | A12 | AMTSUM", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CLIENTNAME | A1 | A2 | A3 | A4 | A5 | A6 | A7 | A8 | A9 | A10 | A11 | A12 | AMTSUM"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME",-1,-1,2,2,false 
			mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", True
		elseif .rdMR.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | GBN | AMTGBN | A | B | O | D | R | P | E | SUMAMT"
			mobjSCGLSpr.SetHeader .sprSht,        "년월|구분|실적|전파|인쇄|온라인|옥외|프로모션|제작|GBS/사보|총계"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  9|   8|   8|  12|  12|    12|  12|      12|  12|      12|  15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | GBN | AMTGBN ", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " A | B | O | D | R | P | E | SUMAMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | GBN | AMTGBN | SUMAMT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GBN | AMTGBN",-1,-1,2,2,false 
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | GBN | AMTGBN "
		else
			Call Grid_init()
		end if
		
   	End With
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, j, strCols
   	Dim strYEARMON
   	Dim strGUBUN
   	Dim intCnt2
   	Dim strRows
   	
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		.txtYEARMON_SRC.value = strYEARMON
		
		IF .rdA.checked THEN
			strGUBUN = .rdA.value
		ELSEIF .rdB.checked THEN
			strGUBUN = .rdB.value
		ELSEIF .rdAO.checked THEN
			strGUBUN = .rdAO.value
		ELSEIF .rdOU.checked THEN
			strGUBUN = .rdOU.value
		ELSEIF .rdMC.checked THEN
			strGUBUN = .rdMC.value
		ELSEIF .rdMR.checked THEN
			strGUBUN = .rdMR.value
		end if
		
		vntData = mobjMDSCREPORT_MST.SelectRtn_REPORT_MST_TOTAL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strGUBUN)

		if not gDoErrorRtn ("SelectRtn_CLIENTYEARCUSTTIMNAMELIST") then
			IF .rdAO.checked THEN
				IF mlngRowCnt > 0 THEN
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, 0, mlngColCnt, mlngRowCnt, True
					for i=3 to .sprSht.MaxCols
						If i = 3 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
					Next
					mobjSCGLSpr.SetColWidth .sprSht, "-1", "13"
					mobjSCGLSpr.SetColWidth .sprSht, "1", "0"
					mobjSCGLSpr.SetColWidth .sprSht, "2", "18"
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht, strRows, -1, -1,0
				END IF
			else
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			end if
			
			
			if .rdMR.checked then
				For i = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "실적" OR _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" OR _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) = "누계"  Then
						If intCnt2 = 1 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
						intCnt2 = intCnt2 + 1
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" THEN 
						mobjSCGLSpr.SetTextBinding .sprSht,"A",i, mobjSCGLSpr.GetTextBinding(.sprSht,"A",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"A",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"B",i, mobjSCGLSpr.GetTextBinding(.sprSht,"B",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"B",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"O",i, mobjSCGLSpr.GetTextBinding(.sprSht,"O",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"O",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"D",i, mobjSCGLSpr.GetTextBinding(.sprSht,"D",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"D",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"R",i, mobjSCGLSpr.GetTextBinding(.sprSht,"R",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"R",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"P",i, mobjSCGLSpr.GetTextBinding(.sprSht,"P",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"P",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"E",i, mobjSCGLSpr.GetTextBinding(.sprSht,"E",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"E",i-1)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUMAMT",i, mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT",i-2) - mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT",i-1)
					End if
					
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출이익" THEN 
						mobjSCGLSpr.SetTextBinding .sprSht,"A",i, mobjSCGLSpr.GetTextBinding(.sprSht,"A",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"A",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"B",i, mobjSCGLSpr.GetTextBinding(.sprSht,"B",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"B",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"O",i, mobjSCGLSpr.GetTextBinding(.sprSht,"O",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"O",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"D",i, mobjSCGLSpr.GetTextBinding(.sprSht,"D",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"D",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"R",i, mobjSCGLSpr.GetTextBinding(.sprSht,"R",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"R",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"P",i, mobjSCGLSpr.GetTextBinding(.sprSht,"P",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"P",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"E",i, mobjSCGLSpr.GetTextBinding(.sprSht,"E",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"E",i-3)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUMAMT",i, mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT",i-6) - mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT",i-3)
					End if
					
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "취급액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "실적" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 11, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 11) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "취급액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "경영계획" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 10, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 10) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "취급액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 9, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 9) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "실적" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 8, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 8) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
						
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "경영계획" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 7, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 7) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출액" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 6, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 6) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출원가" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "실적" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 5, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 5) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
						
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출원가" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "경영계획" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 4, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 4) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출원가" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 3, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 3) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출이익" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "실적" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 2, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 2) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
						
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출이익" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "경영계획" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows - 1, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows - 1) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i) <> "누계" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i) = "매출이익" AND _
					   mobjSCGLSpr.GetTextBinding(.sprSht,"AMTGBN",i) = "차이" THEN 
						
						for j=4 to 11
							mobjSCGLSpr.SetTextBinding .sprSht,j,.sprSht.MaxRows, mobjSCGLSpr.GetTextBinding(.sprSht,j,.sprSht.MaxRows) + mobjSCGLSpr.GetTextBinding(.sprSht,j,i)
						Next
					END IF
				Next
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,4,10,True
			END IF
			
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
			If mobjSCGLSpr.GetTextBinding(.sprSht,3,intCnt) = "총 계" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
			End If
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,2,intCnt) = "총 계" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
			End If
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,3,intCnt) = "경영계획" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
			End If
		Next 
	End With
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim lngCol, lngRow
	Dim strYEARMON
	
	With frmThis
   		if  .sprSht.MaxRows = 0 then 
   			gErrorMsgBox "저장할 테이터가 없습니다.","저장안내"
   			exit sub
   		End if
   		
   		
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
   	
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | GBN | AMTGBN | A | B | O | D | R | P | E | SUMAMT ")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		strYEARMON = .txtYEARMON_SRC.value
		intRtn = mobjMDSCREPORT_MST.ProcessRtn(gstrConfigXml,vntData, strYEARMON)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
			.sprSht.focus()
   		End If
   	end With
End Sub



-->
		</script>
	</HEAD>
	<body class="base">
		<FORM id="frmThis" method="post" runat="server"> <!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--Top TR Start-->
				<TR>
					<TD> <!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="225" background="../../../images/back_p.gIF"
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
											<td class="TITLE">Comm.BU 광고 매출 실적 및 비용 집계</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, '')"
												width="60">년월</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="년월을입력하세요" style="WIDTH: 89px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="12" name="txtYEARMON"><INPUT id="txtYEARMON_SRC" style="WIDTH: 8px; HEIGHT: 21px" type="hidden" name="txtYEARMON_SRC"></TD>
											<TD class="SEARCHDATA" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="DISPLAY: none; CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<TABLE class="SEARCHDATA" id="tblKey" style="BORDER-TOP-STYLE: none" cellSpacing="1" cellPadding="0"
										width="100%" border="0">
										<tr>
											<TD class="SEARCHLABEL" width="60">구분
											</TD>
											<TD class="SEARCHDATA"><INPUT id="rdA" type="radio" CHECKED value="A" name="chkGBN" onclick="vbscript:Call Set_MR('imgSave')">&nbsp;취급액&nbsp;&nbsp;
												<INPUT id="rdB" type="radio" value="B" name="chkGBN" onclick="vbscript:Call Set_MR('imgSave')">&nbsp;매출액&nbsp;&nbsp;
												<INPUT id="rdAO" type="radio" value="AO" name="chkGBN">&nbsp;AOR&nbsp;및 
												대행수수료&nbsp; <INPUT id="rdOU" type="radio" value="OU" name="chkGBN" onclick="vbscript:Call Set_MR('imgSave')">&nbsp;외주비&nbsp;
												<INPUT id="rdMC" type="radio" value="MC" name="chkGBN" onclick="vbscript:Call Set_MR('imgSave')">&nbsp;월별/광고주별 
												매출이익&nbsp;&nbsp; <INPUT id="rdMR" type="radio" value="MR" name="chkGBN" onclick="vbscript:Call Set_MR('imgSave')">&nbsp;월별/매체별 
												매출이익&nbsp;&nbsp;
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31856">
										<PARAM NAME="_ExtentY" VALUE="17119">
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
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></FORM>
	</body>
</HTML>
