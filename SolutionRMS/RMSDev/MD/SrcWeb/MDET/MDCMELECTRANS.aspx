<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANS.aspx.vb" Inherits="MD.MDCMELECTRANS" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 위수탁 거래명세표 전체생성</title>
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECTRANS, mobjMDCMGET
Dim mstrCheck
mstrCheck=True
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
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgNew_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgSaveProc_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_BatchProc
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'----------------------------
'위수탁 현황 TAB BUTTON CLICK
'----------------------------
Sub btnTab1_onclick
	
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	
	pnltab1.style.visibility = "visible" 
	
	mobjSCGLCtl.DoEventQueue
End Sub

'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------
Sub imgCalDemandday_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'청구년월
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'발행일
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 검색조건 년월에 MON 형식을 맞춰주기 위함
'-----------------------------------------------------------------------------------------
Sub txtYEARMON_onblur
	With frmThis
		If .txtYEARMON.value <> "" AND Len(.txtYEARMON.value) = 6 Then DateClean
	End With
End Sub

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim intAMT,intADJAMT,intBALANCE,intCalCul	
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	'서버업무객체 생성	
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "126px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 30, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht,   "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetHeader .sprSht,		   "YEARMON|SEQ|광고주|MEDNAME| 매체사|INPUT_MEDFLAG|매체구분|PROGRAM|ADLOCALFLAG|WEEKDAY|대행금액|부가세|계|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|소재명|GFLAG|SUBSEQ|브랜드명|CLIENTSUBCODE|사업부명"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "	  0|  0|    20|      0|     20|            0|       8|      0|          0|      0|      10|    10|10|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |19    |0    |0     |13      |0            |12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " CLIENTNAME|REAL_MED_NAME|ATTR02|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE|TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true 'GFLAG 앞에 TRANSRANK 추가
		
		
		'합계 표시 그리드 기본화면 구성
		gSetSheetColor mobjSCGLSpr, .sprSht_TRANSSUM
		mobjSCGLSpr.SpreadLayout .sprSht_TRANSSUM, 30, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_TRANSSUM, "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetText .sprSht_TRANSSUM, 3, 1, "           합       계"
	    mobjSCGLSpr.SetScrollBar .sprSht_TRANSSUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_TRANSSUM,"1|3",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_TRANSSUM, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_TRANSSUM, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_TRANSSUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_TRANSSUM
    End With    
    
	pnlTab1.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'기본값 설정
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
	.txtYEARMON.value =  Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtYEARMON.value = vntInParam(i)	
	'			case 1 : mstrFields = vntInParam(i)
	'			case 2 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
	'			case 3 : mstrUseDate = vntInParam(i)		'코드 사용 시점
	'			case 4 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
	'		end select
	'	next
	end with
	DateClean
	'SelectRtn
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECTRANS = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
	.txtPRINTDAY.value  = gNowDate
	.sprSht.MaxRows = 0	

	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON.value,1,4) & "-" & MID(frmThis.txtYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub
Sub ProcessRtn_BatchProc
	
		Dim intSaveChkRtn	
		Dim intRtn
   		Dim vntData
		Dim strMasterData
		Dim strTRANSYEARMON
		Dim intTRANSNO
		Dim intRANKTRANS
		Dim intCnt,bsdiv
		Dim intColFlag
		Dim strDESCRIPTION
		intSaveChkRtn = gYesNoMsgbox("광고주별 거래명세서를 전체 생성 하시겠습니까?","자료삭제 확인")
		IF intSaveChkRtn <> vbYes then exit Sub
		
		
		
		If SelectRtn_Proc = False Then Exit Sub	
		with frmThis
			strDESCRIPTION = ""
			'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
			If .txtDEMANDDAY.value = "" Then
				msgbox "청구일은 필수 입력 사항 입니다."
				Exit Sub
			End If
			
			If .txtPRINTDAY.value = "" Then
				msgbox "발행일은 필수 입력 사항 입니다."
				Exit Sub
			End If

			'저장플레그 설정
			mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
			gXMLSetFlag xmlBind, meINS_TRANS

			'그룹 최대값 설정
			intColFlag = 0
			For intCnt = 1 To .sprSht.MaxRows
			'최대값
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			Next
			
   			'데이터 Validation
   			If .sprSht.MaxRows = 0 Then
   				msgbox "상세항목 이 없습니다."
   				Exit Sub
   			End If
			if DataValidation =false then exit sub
			'On error resume next
			'쉬트의 변경된 데이터만 가져온다.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
			
			'마스터 데이터를 가져 온다.
			strMasterData = gXMLGetBindingData (xmlBind)
			
			'처리 업무객체 호출
			intTRANSNO = 0
			strTRANSYEARMON = .txtYEARMON.value
			
			intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

			if not gDoErrorRtn ("ProcessRtn") then
				'모든 플래그 클리어
				
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				InitPageData
				gOkMsgBox "거래명세서가 생성되었습니다.","확인"
				gEndPage
   			end if
   		end with

End Sub
'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
	Dim intSaveChkRtn
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
	
	intSaveChkRtn = gYesNoMsgbox("사업부별 거래명세서를 전체 생성 하시겠습니까?","자료삭제 확인")
	IF intSaveChkRtn <> vbYes then exit Sub
	with frmThis
		strDESCRIPTION = ""
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .txtDEMANDDAY.value = "" Then
			msgbox "청구일은 필수 입력 사항 입니다."
			Exit Sub
		End If
		
		If .txtPRINTDAY.value = "" Then
			msgbox "발행일은 필수 입력 사항 입니다."
			Exit Sub
		End If

		 '저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'그룹 최대값 설정
		intColFlag = 0
		For intCnt = 1 To .sprSht.MaxRows
		'최대값
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
   		'데이터 Validation
   		If .sprSht.MaxRows = 0 Then
   			msgbox "상세항목 이 없습니다."
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intTRANSNO = 0
		strTRANSYEARMON = .txtYEARMON.value
		
		intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			InitPageData
			gOkMsgBox "거래명세서가 생성되었습니다.","확인"
			gEndPage
   		end if
   	end with
End Sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	'On error resume next
	with frmThis
'		intColSum = 0
' 		for intCnt = 1 to .sprSht.MaxRows
'			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1  Then 
'				intColSum = intColSum + 1
'			End if
'		next
'		If intColSum = 0 Then exit Function
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	'쿼리튜닝 필요 변수
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
	'On error resume next
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		'Sheet초기화
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtYEARMON.value
		
		vntPreData = mobjMDCMELECTRANS.SelectRtn_PreCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
			if not gDoErrorRtn ("SelectRtn_PreCnt") then
				lngCnt = vntPreData(0,0)
				If lngCnt < 100 Then
					lngCnt = 1
				Else
					lngCnt = int(lngCnt/100)
					lngCnt = lngCnt+1
				End If
			End if
		
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		vntDataConfirm = mobjMDCMELECTRANS.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		
		If IngCOMMITRowCnt = 0 Then
			gErrorMsgBox strYEARMON & "월은 승인처리되지 않았습니다.",""
			EXIT SUB	
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		'For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			end if
   			'strST = strST + 100
   			'strED = strED + 100
   		'Next
   		'for intDelCnt = .sprSht.MaxRows to 1 step -1				
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intDelCnt) = "" Then
		'		mobjSCGLSpr.DeleteRow .sprSht,intDelCnt
		'	End If		
		'next
   		
   		
   		AMT_SUM
   		PreSearchFiledValue strYEARMON	
   		gWriteText lblStatus, "위수탁 " & mlngRowCnt & " 건 의 자료가 검색" & mePROC_DONE
   	end with
End Sub
Function SelectRtn_Proc ()
SelectRtn_Proc = False
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	'쿼리튜닝 필요 변수
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
	'On error resume next
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "년월은 반드시 넣어야 합니다.",""
			Exit Function
		End If 
		'Sheet초기화
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtYEARMON.value
		
		vntPreData = mobjMDCMELECTRANS.SelectRtn_PreCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
			if not gDoErrorRtn ("SelectRtn_PreCnt") then
				lngCnt = vntPreData(0,0)
				If lngCnt < 100 Then
					lngCnt = 1
				Else
					lngCnt = int(lngCnt/100)
					lngCnt = lngCnt+1
				End If
			End if
		
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		vntDataConfirm = mobjMDCMELECTRANS.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		
		If IngCOMMITRowCnt = 0 Then
			gErrorMsgBox strYEARMON & "월은 승인처리되지 않았습니다.",""
			EXIT Function	
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		'For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn_Proc(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			end if
   		'	strST = strST + 100
   		'	strED = strED + 100
   		'Next
   		'for intDelCnt = .sprSht.MaxRows to 1 step -1				
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intDelCnt) = "" Then
		'		mobjSCGLSpr.DeleteRow .sprSht,intDelCnt
		'	End If		
		'next
   		
   		
   		AMT_SUM
   		'PreSearchFiledValue strYEARMON	
   		'gWriteText lblStatus, "위수탁 " & mlngRowCnt & " 건 의 자료가 검색" & mePROC_DONE
   	end with
SelectRtn_Proc = True
End Function

Sub PreSearchFiledValue (strYEARMON)
	frmThis.txtYEARMON.value = strYEARMON
End Sub

'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntVAT, IntVATSUM, IntSUMATMVAT, IntSUMATMVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMATMVATSUM = 0
		
		'위수탁 그리드 합계그리드 값넣기
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntVAT = 0
			IntSUMATMVAT = 0
			
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
			IntSUMATMVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMATMVAT", lngCnt)
			
			IntAMTSUM = IntAMTSUM + IntAMT
			IntVATSUM = IntVATSUM + IntVAT
			IntSUMATMVATSUM = IntSUMATMVATSUM + IntSUMATMVAT
		Next
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"SUMATMVAT",1, IntSUMATMVATSUM
		end if
	End With
End Sub


'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_TRANSSUM	
	End with
end sub

'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_TRANSSUM, NewTop, NewLeft
End Sub

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON, dblSEQ

	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",vntData(i))
			
				intRtn = mobjMDCMELECTRANS.DeleteRtn(gstrConfigXml,dblSEQ, strYEARMON)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			End IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intSelCnt & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">
													&nbsp;위수탁 거래명세서 관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0"> <!--TopSplit Start->
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="top" align="center">
										<TABLE class="DATA" id="tblDATA1" style="WIDTH: 1040px" cellSpacing="1" cellPadding="0"
											width="1040" border="0">
											<TR>
												<TD class="SEARCHLABEL" title="삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">년월</TD>
												<TD class="SEARCHDATA" width="180"><INPUT class="INPUT" id="txtYEARMON" title="년월조회" style="WIDTH: 89px; HEIGHT: 22px" type="text"
														maxLength="6" size="9" name="txtYEARMON" accessKey="MON"></TD>
												<TD class="SEARCHLABEL" title="삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')">청구일자</TD>
												<TD class="SEARCHDATA" width="180"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="청구일자" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></TD>
												<TD class="SEARCHLABEL" title="삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')">발행일자</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="발행일자" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="12" name="txtPRINTDAY"><IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalPrintday">
												</TD>
												<TD class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px;HEIGHT: 25px"></TD>
					</TR>
					<TR>
						<TD class="KEYFRAME" vAlign="middle" align="center">
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">
													&nbsp;거래명세서 전체생성</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/ImgTRANSALLSUBOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTRANSALLSUB.gIF'"
														height="20" alt="해당월의 사업부별 거래명세서 전체를 생성합니다. " src="../../../images/ImgTRANSALLSUB.gIF"
														border="0" name="imgSave"></TD>
												<TD><IMG id="imgSaveProc" onmouseover="JavaScript:this.src='../../../images/ImgTRANSALLOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTRANSALL.gIF'"
														height="20" alt="해당월의 광고주별 거래명세서 전체를 생성합니다." src="../../../images/ImgTRANSALL.gIF"
														border="0" name="imgSaveProc"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gIF'"
														height="20" alt="자료를 저장합니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 714px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 1040px; HEIGHT: 690px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
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
								<OBJECT id="sprSht_TRANSSUM" style="WIDTH: 1040px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="635">
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
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
