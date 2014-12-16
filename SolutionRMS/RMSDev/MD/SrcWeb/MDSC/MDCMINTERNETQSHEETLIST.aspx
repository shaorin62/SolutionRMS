<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNETQSHEETLIST.aspx.vb" Inherits="MD.MDCMINTERNETQSHEETLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>월별 매체별 광고비</title>
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
Dim mobjMDCOGET, mobjEXECUTE, mobjMDSRREPORTLIST'공통코드, 클래스
Dim mClientsubcode

Dim mintCnt
Dim mintCnt2
Dim mintCnt3
Dim mvntData3
Dim mstrField
Dim mintCntExist
Dim mstrFieldExist
Dim mvntDataCust
Dim mvntDataMed
Dim mvntDataBRANDCNT
Dim mvntDataBRAND
Dim mvntDataCAMPAIGN
Dim mvntDataCAMPAIGNCNT
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
	
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	if frmThis.txtCLIENTCODE.value = ""  then
		gErrorMsgBox "광고주코드를 입력하시오","조회안내"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
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
	Dim strSUBLIST
	Dim strCLIENTSUBLIST
	Dim intSUBRow
	Dim chkflag
	Dim strCLIENTCODE
	
	Dim Con1 
	Dim Con2
	Dim Con3
	
	with frmThis
		Con1 = ""
		Con2 = ""
		Con3 = ""
		gErrorMsgBox "출력물은 개발 중입니다..",""
		EXIT SUB
		if frmThis.sprSht.MaxRows = 0 then
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
			Exit Sub
		end if
		
		
		ModuleDir = "MD"
		ReportName = "MDCMMONAMTLIST.rpt"
		
		strFYEARMON		= .txtFYEARMON.value
		strTYEARMON		= .txtTYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		
		If strYEARMON <> "" Then Con1 = " AND (YEARMON = '" & strYEARMON & "')"
		If strCLIENTCODE <> "" Then Con2 = " AND (CLIENTCODE = '" & strCLIENTCODE & "')"
		
		strCLIENTSUBLIST=""
		strSUBLIST = ""
		chkflag = 1
		strCLIENTSUBLIST = 	split(mClientsubcode,"♥")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
		
		if strSUBLIST <> "" then Con3 = " AND (CLIENTSUBCODE IN(" & strSUBLIST & "))"
		strCLIENTNAME = .txtCLIENTNAME.value
        
		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & strCLIENTNAME & ":" & strYEARMON
		
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			if not gDoErrorRtn ("GetHIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
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
' 팀코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgTIMCODE_onclick
	Call TIMCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value), _
							trim(.txtTIMCODE.value), trim(.txtTIMNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtTIMNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' 코드명 표시
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' 코드명 표시
			SELECTRTN
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
			
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
											trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME.value = trim(vntData(1,1))       ' 코드명 표시
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					SELECTRTN
				Else
					Call TIMCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST")
	set mobjEXECUTE	= gCreateRemoteObject("cMDCO.ccMDCOEXECUTE")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "MEDNAME"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
												'  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   												'1|
		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MEDNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME",-1,-1,2,2,false
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSRREPORTLIST = Nothing
	set mobjMDCOGET = Nothing
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
		.txtYEARMON.value = mid(gNowDate2,1,4) & mid(gNowDate2,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtYEARMON.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim intLayOutCnt
   	
	'On error resume next
	with frmThis
		If .txtYEARMON.value = ""  Then
			gErrorMsgbox "조회년월을 선택하세요","조회안내"
			Exit Sub
		End If
		'그리드 재생성 
		SetChangeLayout
		'EXIT SUB
		IF mvntDataBRANDCNT = 0 THEN
			EXIT SUB
		END IF
		
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
			
		vntData = mobjMDSRREPORTLIST.SelectRtn_INTERNETQSHEETLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, .txtCLIENTCODE.value, .txtTIMCODE.value, mvntDataBRAND,mvntDataBRANDCNT, mvntDataCAMPAIGN, mvntDataCAMPAIGNCNT)

		if not gDoErrorRtn ("SelectRtn_INTERNETQSHEETLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub SetChangeLayout () 
	Dim strYEARMON
	Dim strCLIENTCODE
	Dim intAddCnt,intAddHeadCnt,intAddWith,intFieldSetting,intHide,intFloat,intAddCnt2 'For 문 Count변수
	Dim vntData
	Dim strAddHead
	Dim lngRowReal
	Dim lngColReal
	Dim strStartHead
	Dim strEndHead
	Dim strCLIENTSUBLIST
	Dim strSUBLIST
	Dim intSUBRow
	Dim chkflag
	
	Dim strClientAndMed
	Dim i
	
	mvntDataBRAND = ""
	mvntDataBRANDCNT = 0
	gInitComParams mobjSCGLCtl,"MC"
	With frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		lngRowReal=clng(0)
		lngColReal=clng(0)
		
		strYEARMON = .txtYEARMON.value
		
		
		mvntDataCAMPAIGN = mobjMDSRREPORTLIST.GetINTERNETCAMPAIGN(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, .txtCLIENTCODE.value, .txtTIMCODE.value)
		mvntDataCAMPAIGNCNT = mlngRowCnt
		
		mvntDataBRAND = mobjMDSRREPORTLIST.GetINTERNETBRAND(gstrConfigXml,lngRowReal,lngColReal,strYEARMON, .txtCLIENTCODE.value, .txtTIMCODE.value)
		
		mvntDataBRANDCNT = lngRowReal
		
		If mvntDataBRANDCNT > 0 Then 
			'필드 고정값세팅
			Dim strField
			strField = "MEDNAME"
			
			'필드 증가값세팅 [광고주코드]
			Dim strAddField
			strAddField = ""
			For intAddCnt = 1 To mvntDataBRANDCNT
				strAddField = strAddField & "|A" & intAddCnt
			Next
			
			'필드 증가값 [값]
			mstrField = strField & strAddField & "|SUMAMT|REAL_MED_NAME"
			
			'헤더 고정값세팅
			Dim strHead
			strHead = "사이트명"
			'헤더 증가값세팅
			Dim strHeadCLIENT
			Dim strHeadMED
			Dim lngSUBCNT
			Dim strHeadCLIENTSUB
			Dim strENDSUBHead
			
			lngSUBCNT =1
			strHeadCLIENT = ""
			strHeadCLIENTSUB = ""
			strHeadMED = ""
			strStartHead = ""
			strEndHead = ""
			strENDSUBHead = ""
			strHeadCLIENTSUB = ""
			
			Dim cntzzz
			cntzzz = 0
			FOR intAddHeadCnt = 1 TO mvntDataCAMPAIGNCNT
				for i = 1 to mvntDataCAMPAIGN(2,intAddHeadCnt)
					if i = 1 then
						strHeadMED = strHeadMED & "|" & mvntDataCAMPAIGN(1,intAddHeadCnt)
					else 
						strHeadMED = strHeadMED & "|"
					end if
				next
			NEXT 
			
			For intAddHeadCnt = 1 To  mvntDataBRANDCNT
				IF intAddHeadCnt MOD (mvntDataBRANDCNT+1) = 0 THEN
					'strHeadCLIENT   = strHeadCLIENT & "|소계" 
				ELSE 
					strHeadCLIENT   = strHeadCLIENT & "|" & TRIM(mvntDataBRAND(2,intAddHeadCnt MOD (mvntDataBRANDCNT+1)))
				END IF
				
				IF intAddHeadCnt MOD (mvntDataBRANDCNT+1) = 0 THEN
					'strHeadCLIENT   = strHeadCLIENT & "|소계" 
				ELSE 
					strHeadCLIENTSUB   = strHeadCLIENTSUB & "|" & mid(TRIM(mvntDataBRAND(3,intAddHeadCnt MOD (mvntDataBRANDCNT+1))),5,2) & "/" & mid(TRIM(mvntDataBRAND(3,intAddHeadCnt MOD (mvntDataBRANDCNT+1))),7,2) & "~" &	mid(TRIM(mvntDataBRAND(4,intAddHeadCnt MOD (mvntDataBRANDCNT+1))),5,2) & "/" & mid(TRIM(mvntDataBRAND(4,intAddHeadCnt MOD (mvntDataBRANDCNT+1))),7,2)
				END IF
			NEXT
			
			strStartHead = strHead & strHeadMED & "|소계|청구지"
			strEndHead =  strHeadCLIENT & "||"
			strENDSUBHead = strHeadCLIENTSUB & "||"
			
			'넓이 고정값세팅
			Dim strWith
			strWith = "20"
			'넓이 증가값세팅
			Dim strAddWith
			Dim strEndWith
			strAddWith = ""
			For intAddWith = 1 To mvntDataBRANDCNT
				strAddWith = strAddWith & "|15"
			Next
			strEndWith = strWith & strAddWith & "|20|20"
			
			
			'총컬럼갯수
			Dim intLayOutCnt
			intLayOutCnt = 1 + mvntDataBRANDCNT + 2
			'여기까지 괜찮음
			
			gSetSheetColor mobjSCGLSpr, .sprSht
			
			'그리드 초기화(셀합칠때 문제가 생기기 때문에 넣어놓음)	
			Call Grid_init()
			'Sheet Layout 디자인
			mobjSCGLSpr.SpreadLayout .sprSht, intLayOutCnt, 0, 1, 0, , 3, 0, , , True
			mobjSCGLSpr.SpreadDataField .sprSht, mstrField 
			mobjSCGLSpr.SetHeader .sprSht,       strStartHead ,0,1,true
			mobjSCGLSpr.SetHeader .sprSht,       strEndHead ,SPREAD_HEADER + 1,1,true
			mobjSCGLSpr.SetHeader .sprSht,       strENDSUBHead ,SPREAD_HEADER + 2, 1,true
			mobjSCGLSpr.AddCellSpan .sprSht, 1, SPREAD_HEADER + 0, 1    , 3      , 1 , true
			Dim cntaa
			cntaa = 2
			for i = 1 to mvntDataCAMPAIGNCNT
				mobjSCGLSpr.AddCellSpan .sprSht, cntaa, SPREAD_HEADER + 0, mvntDataCAMPAIGN(2,i)    , 1, -1, true
				cntaa = cntaa + mvntDataCAMPAIGN(2,i) 
			NEXT 
			'                                 20번째 부터            하위6개를 1개로 3번단위로 나눠서
			mobjSCGLSpr.AddCellSpan .sprSht, intLayOutCnt-1, SPREAD_HEADER + 0, 1    , 3      , -1 , true
			'                                 마지막 풀리는곳 은 44번째이고 2개로 합쳐라 -1 전체
			mobjSCGLSpr.SetColWidth .sprSht, "-1", strEndWith
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, mstrField, -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME|REAL_MED_NAME", , , 50, , ,0
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "18"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,strField
			mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|REAL_MED_NAME",-1,-1,2,2,false
		ELSE
			'Sheet 기본Color 지정
			gSetSheetDefaultColor() 
			
			With frmThis
				gSetSheetColor mobjSCGLSpr, .sprSht
				mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
				mobjSCGLSpr.SpreadDataField .sprSht, "MEDNAME"
				mobjSCGLSpr.SetHeader .sprSht,		 ""
														'  1|
				mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   														'1|
				
				mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
				mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME", -1, -1, 20
				mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MEDNAME"
				mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME",-1,-1,2,2,false
				
			End With
		End If
   	End With
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "MEDNAME"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
												'  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   												'1|
		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MEDNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME",-1,-1,2,2,false
	End With
End Sub


Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",intCnt) = "소계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",intCnt) = "브랜드별 소계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		end if
		
		If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",intCnt),4) = "청구금액" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",intCnt) = "계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		End If
	Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
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
												<TABLE cellSpacing="0" cellPadding="0" width="125" background="../../../images/back_p.gIF"
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
											<td class="TITLE">광고주/부문별 큐시트</td>
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
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="110" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="95%" border="0"> <!--TopSplit Start->
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TBODY>
											<TR>
												<TD class="SEARCHLABEL" title="년도을삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">년&nbsp;&nbsp;월
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 110px" width="110"><INPUT class="INPUT" id="txtYEARMON" title="년도을입력하세요" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="9" name="txtYEARMON">
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
													width="80">광고주
												</TD>
												<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 207px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME, txtTIMCODE)"
													width="80">팀&nbsp;
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME" title="팀명" style="WIDTH: 207px; HEIGHT: 22px" type="text"
														maxLength="100" size="20" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
														align="absMiddle" border="0" name="ImgTIMCODE"> <INPUT class="INPUT_L" id="txtTIMCODE" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTIMCODE"></TD>
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
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
								width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31591">
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
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
