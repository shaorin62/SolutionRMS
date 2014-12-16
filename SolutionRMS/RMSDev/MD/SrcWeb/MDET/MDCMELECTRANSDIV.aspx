<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANSDIV.aspx.vb" Inherits="MD.MDCMELECTRANSDIV" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 거래명세서 분할생성 및 발행</title>
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
Dim mobjMELECSPONTRANS, mobjMDCMGET, mobjMDCMELECTRANSLIST
Dim mobjMDCMCODETR	
Dim mstrCheck
Dim mALLCHECK
Dim mbuttonchk
Dim mobjMDCMELECTRANS
mALLCHECK = TRUE
mstrCheck=True
mbuttonchk = true
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

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgAllSet_onclick()
	'Dim intRtn
	'intRtn = gYesNoMsgbox("전체 거래명세서를 생성 하시겠습니까?","저장안내!")
	'IF intRtn <> vbYes then exit Sub
	'gFlowWait meWAIT_ON
	'If ProcessRtn_All = True Then
	'ProcessAllRtn
	'End If
	'gFlowWait meWAIT_OFF ''이상 예전 버튼 로직 
	'변경 20081031
	
	ProcessRtn_BatchProc
End Sub
Sub ProcessRtn_BatchProc
	dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		
			
		vntInParams = ""
		vntRet = gShowModalWindow("MDCMELECTRANS.aspx",vntInParams , 1062,900)
		SelectRtn
	End with
	'gSetChange
End Sub


Function ProcessRtn_All ()

	ProcessRtn_All = False
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
   	Dim intRtn
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
	Dim vntData2
	
	with frmThis
	If .txtTRANSYEARMON.value = "" Then
		gErrorMsgBox "년월을 입력하십시오.","저장안내!"
		Exit Function
	End If
	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'Sheet초기화
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtTRANSYEARMON.value
		
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
		For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprShtAll, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprShtAll,meCLS_FLAG
   			end if
   			strST = strST + 100
   			strED = strED + 100
   		Next
   		for intDelCnt = .sprShtAll.MaxRows to 1 step -1				
			If mobjSCGLSpr.GetTextBinding(.sprShtAll,"YEARMON",intDelCnt) = "" Then
				mobjSCGLSpr.DeleteRow .sprShtAll,intDelCnt
			End If		
		next
   		'조회 완료 시점
   		End with
   ProcessRtn_All = True
End Function

Sub ProcessAllRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
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
		mobjSCGLSpr.SetFlag  .sprShtAll,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'그룹 최대값 설정
		intColFlag = 0
		For intCnt = 1 To .sprShtAll.MaxRows
		'최대값
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprShtAll,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
   		'데이터 Validation
   		If .sprShtAll.MaxRows = 0 Then
   			msgbox "생성할 내역이 없습니다."
   			Exit Sub
   		End If
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprShtAll,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intTRANSNO = 0
		strTRANSYEARMON = .txtTRANSYEARMON.value
		
		intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'InitPageData
			gOkMsgBox "거래명세서가 생성되었습니다.","확인"
   		end if
   	end with
End Sub


Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
		
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		vntData = mobjMDCMELECTRANSLIST.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_ELETRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt
			strUSERID = ""
			vntDataTemp = mobjMDCMELECTRANSLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
		End IF
	
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "printSetTimeout", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub



Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP ()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		strSPONSOR = ""
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELEC", strSPONSOR) 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			IF vntRet(3,0) = "완료" THEN
				.txtTRANSYEARMON.value = vntRet(0,0)
				.txtTRANSNO.value = vntRet(1,0)
				.txtCLIENTCODE.value = vntRet(4,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			ELSE
				.txtTRANSYEARMON.value = vntRet(0,0)
				.txtTRANSNO.value = ""
				.txtCLIENTCODE.value = vntRet(1,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			END IF
			gSetChangeFlag .txtCLIENTCODE             ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strSPONSOR
   		
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strSPONSOR = ""
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtCLIENTCODE.value,.txtCLIENTNAME1.value,"","trans", "ELEC", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,1)
					.txtTRANSNO.value = ""
					.txtCLIENTCODE.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
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
' 거래처번호팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------

'이미지버튼 클릭시
Sub ImgTRU_onclick
	Call TRU_POP()
End Sub

Sub txtTRANSNO_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strTRANSYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strTRANSYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELEC", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)  ' Code값 저장
					.txtTRANSNO.value = vntData(1,0)  ' 코드명 표시
					.txtCLIENTCODE.value = vntData(2,0)  ' 코드명 표시
					.txtCLIENTNAME1.value = vntData(3,0)  ' 코드명 표시
					'Call SelectRtn ()
				Else
					Call TRU_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRU_POP
	dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
			strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
			
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "ELEC") '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code값 저장
			.txtTRANSNO.value = vntRet(1,0)  ' 코드명 표시
			.txtCLIENTCODE.value = vntRet(2,0)  ' 코드명 표시
			.txtCLIENTNAME1.value = vntRet(3,0)  ' 코드명 표시
			'Call SelectRtn ()
		end if
	End with
	gSetChange
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
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'단가
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'금액
Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

'수수료
Sub txtSUMAMTVAT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMTVAT,0,true)
	end with
End Sub

Sub txtTRANSYEARMON_onblur
	With frmThis
		if .txtTRANSNO.value ="" then
			If .txtTRANSYEARMON.value <> "" AND Len(.txtTRANSYEARMON.value) = 6 Then DateClean
		end if
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
	dim amt1, vat1, sumamtvat
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
	if col = 15 then
		amt1 = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"AMT",Row) 
		vat1 = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"VAT",Row) 
		sumamtvat = amt1 + vat1
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",Row, sumamtvat
		AMT_SUM1
		
	end if	
End Sub

sub sprSht_ButtonClicked(Col,Row,ButtonDown)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	with frmThis
	IF mbuttonchk THEN
		if ButtonDown then
			lngAMT = 0
			lngSUMAMT = 0
			lngTOT = 0
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) 
			lngSUMAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"REAL_MED_NAME",1)
			lngTOT = lngAMT + lngSUMAMT
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			mobjSCGLSpr.DeselectBlock .sprSht
		else
			lngAMT = 0
			lngSUMAMT = 0
			lngTOT = 0
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) 
			lngSUMAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"REAL_MED_NAME",1)
			lngTOT = lngSUMAMT - lngAMT
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			mobjSCGLSpr.DeselectBlock .sprSht
		end if
	END IF
	end with
end sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
	if Col = 1 and Row = 0 then
		mbuttonchk = false
		mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
		if mstrCheck = True then 
			mstrCheck = False
			lngTOT = 0
			For i=1 to .sprSht.MaxRows
				lngAMT = 0
				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",i) 
				lngTOT = lngTOT + lngAMT
				mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			Next
		elseif mstrCheck = False then 
			mstrCheck = True
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, 0
		end if
	elseif Col = 1 and Row > 0 then
		mbuttonchk = true
	end if 
	End With
End Sub  

'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM	
	End with
end sub

'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_SUM, NewTop, NewLeft
End Sub


'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt
	Dim IntAMT, IntVAT, IntSUMAMTVAT, IntAMTSUM, IntVATSUM, IntSUMAMTVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMAMTVATSUM = 0
		'위수탁 그리드 합계그리드 값넣기
		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntVAT = 0
				IntSUMAMTVAT = 0
				
				IntAMT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntVAT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
				IntSUMAMTVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt)
				
				IntAMTSUM		= IntAMTSUM + IntAMT
				IntVATSUM		= IntVATSUM + IntVAT
				IntSUMAMTVATSUM	= IntSUMAMTVATSUM + IntSUMAMTVAT
			Next
		END IF
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, IntSUMAMTVATSUM
		end if
	End With
End Sub

'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub AMT_SUM1
	Dim lngCnt
	Dim IntAMT, IntVAT, IntSUMAMTVAT, IntAMTSUM, IntVATSUM, IntSUMAMTVATSUM
	With frmThis
		IntVATSUM = 0
		IntSUMAMTVATSUM = 0
		'위수탁 그리드 합계그리드 값넣기
		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntVAT = 0
				IntSUMAMTVAT = 0
				
				IntVAT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
				IntSUMAMTVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt)
				
				IntVATSUM		= IntVATSUM + IntVAT
				IntSUMAMTVATSUM	= IntSUMAMTVATSUM + IntSUMAMTVAT
			Next
		END IF
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, IntSUMAMTVATSUM
		end if
	End With
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMELECSPONTRANS	= gCreateRemoteObject("cMDET.ccMDETELECNOSPON")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECTRANSLIST = gCreateRemoteObject("cMDET.ccMDETELECTRANSLIST")
	set mobjMDCMCODETR	= gCreateRemoteObject("cMDCO.ccMDCOCODETR")
	set mobjMDCMELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "240px"
	pnlTab1.style.left= "7px"
	
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "240px"
	pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'거래명세서 생성 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 32, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,   "CHK |YEARMON| SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|년월|순번|광고주|매체명|그룹구분|매체사|소재명|매체구분코드|매체구분|프로그램|지역|요일|대행금액|부가세|계|COMMISSION|DEPTCD|단가|횟수|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|브랜드코드|브랜드명|사업부코드|사업부명"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|   0|   0|    20|     0|       0|    14|    28|           0|       8|       0|   0|   0|      12|    12|13|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |0    |0         |16      |0         |15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " CLIENTNAME|REAL_MED_NAME|GFLAGNAME|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME|ATTR02", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|CLIENTNAME|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		'합계 표시 그리드 기본화면 구성
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 32, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "CHK|YEARMON | SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetText .sprSht_SUM, 1, 1, "합계"
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"1|1",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|VAT|SUMAMTVAT|REAL_MED_NAME", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_SUM, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
	    
		'******************************************************************
		'거래명세서 조회 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 28, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht1, "TRANSYEARMON | TRANSNO | SEQ |CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|PROGRAM |ADLOCALFLAG|WEEKDAY|DEPT_CD|DEMANDDAY|PRINTDAY| PRICE| CNT|AMT| VAT| MED_FLAG|ATTR02|TAXYEARMON|TAXNO|TRUST_SEQ|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME"
		mobjSCGLSpr.SetHeader .sprSht1,		  "TRANSYEARMON|TRANSNO|SEQ|CLIENTCODE|CLIENTNAME|MEDCODE|매체명|REAL_MED_CODE|매체사|프로그램명|지역|방송요일|DEPT_CD|DEMANDDAY|발행일자|공급가액|회수|광고대금|부가세|매체종류|소재명|세금계산서년월|세금계산서번호|신탁순번|브랜드코드|브랜드명|사업부코드|사업부명"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "         0|	  0|  0|         0|	        0|      0|    10|            0|    10|        10|   6|       8|      0|        0|       8|       9|   5|       9|     9|       9|    15|0             |0             |0       |0         |12      |0         |12"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "PRICE | AMT| VAT|CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "MEDNAME|REAL_MED_NAME|PROGRAM|ADLOCALFLAG|WEEKDAY|MED_FLAG|ATTR02|BRANDNAME|CLIENTSUBNAME", -1, -1, 50
		mobjSCGLSpr.ColHidden .sprSht1, "TRANSYEARMON|TRANSNO | SEQ|CLIENTCODE|CLIENTNAME|MEDCODE|REAL_MED_CODE|DEPT_CD|DEMANDDAY|TAXYEARMON|TAXNO|TRUST_SEQ|SUBSEQ|CLIENTSUBCODE", true
		'******************************************************************
		'전체 거래명세서 생성용 Hidden 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprShtAll
		mobjSCGLSpr.SpreadLayout .sprShtAll, 30, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprShtAll,   "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetHeader .sprShtAll,		   "YEARMON|SEQ|광고주|MEDNAME| 매체사|INPUT_MEDFLAG|매체구분|PROGRAM|ADLOCALFLAG|WEEKDAY|대행금액|부가세|계|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|소재명|GFLAG|SUBSEQ|브랜드명|CLIENTSUBCODE|사업부명"
		mobjSCGLSpr.SetColWidth .sprShtAll, "-1", "	  0|  0|    20|      0|     20|            0|       8|      0|          0|      0|      10|    10|10|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |19    |0    |0     |13      |0            |12"
		mobjSCGLSpr.SetRowHeight .sprShtAll, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprShtAll, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprShtAll, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprShtAll, " CLIENTNAME|REAL_MED_NAME|ATTR02|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprShtAll, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprShtAll, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
    End With    
	pnlTab1.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'기본값 설정
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
	'			case 1 : .txtCLIENTCODE.value = vntInParam(i)
	'			case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'조회추가필드
	'			case 3 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
	'			case 4 : mstrUseDate = vntInParam(i)		'코드 사용 시점
	'			case 5 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
	'		end select
	'	next
	.txtTRANSYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	end with
	'SelectRtn
End Sub

Sub EndPage()
	set mobjMELECSPONTRANS = Nothing
	set mobjMDCMGET = Nothing
	set mobjMDCMCODETR = Nothing
	Set mobjMDCMELECTRANS = Nothing
	
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean
		.txtPRINTDAY.value  = gNowDate
		.sprSht.MaxRows = 0	
		.sprSht1.MaxRows = 0	
		
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"PRICE",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"CNT",1, ""
	
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
	Clean_Hdr	
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim intRtnCf
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strPROGNAME
	Dim strTRANSNO
	
	chkcnt = 0
	with frmThis
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "발행일은 필수 입력 사항 입니다.",""
			Exit Sub
		End If

		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "거래명세서를 생성할 데이터를 체크 하십시오",""
			exit sub
		end if

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
   			gErrorMsgBox "상세항목 이 없습니다.",""
   			Exit Sub
   		End If
		'쉬트의 변경된 데이터만 가져온다.
		intRtnCf = gYesNoMsgbox("청구일을 확인하셨습니까?","저장안내!")
		IF intRtnCf <> vbYes then exit Sub
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON | SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) &  MID(.txtDEMANDDAY.value,6,2) 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strREAL_MED_CODE	= .cmbREAL_MED_CODE.value
		'strPROGNAME		= .txtPROGNAME.value
		strTRANSNO		= 0
		
		
		intRtn = mobjMELECSPONTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		'if not gDoErrorRtn ("ProcessRtn") then
		'	'모든 플래그 클리어
		'	mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		'	gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
		'	InitPageData
		'	PreSearchFiledValue strTRANSYEARMON, strTRANSNO, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strPROGNAME
   		'end if
   		
   		''''''''''''''''''''''''''
   		'intRtn = mobjMDCMINTERNETTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)
   		
   		if not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			If intRtn <> 0  Then
				.txtTRANSYEARMON.value = strTRANSYEARMON
				.txtTRANSNO.value = intTRANSNO
				.txtCLIENTCODE.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				selectRtn
			Else
				initpagedata
			End If
   		end if
   		
   		
   		'''''''''''''''''''''''''
   	end with
End Sub


'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData1
	Dim strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME, strTRANSNO
	Dim strPRINTDAY
	Dim strREAL_MED_CODE, strREAL_MED_NAME, strPROGNAME
   	Dim i, strCols
	Dim strMEDGUBUN
	'On error resume next
	with frmThis

		If .txtTRANSYEARMON.value = "" Then
			gErrorMsgBox "년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		
		strTRANSNO = ""
		strTRANSNO = .txtTRANSNO.value
		IF strTRANSNO = "" THEN
			IF  .txtCLIENTCODE.value = "" THEN
				gErrorMsgBox "조회시 광고주코드는 반드시 넣어야 합니다.",""
				Exit SUb
			END IF
			IF  .cmbREAL_MED_CODE.value = "" THEN
				gErrorMsgBox "조회시 매체사코드는 반드시 넣어야 합니다.",""
				Exit SUb
			END IF
		END IF
		
		'Sheet초기화
		.sprSht.MaxRows = 0
		.sprSht1.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON	= .txtTRANSYEARMON.value
		strTRANSNO		= .txtTRANSNO.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strREAL_MED_CODE	= .cmbREAL_MED_CODE.value
		strMEDGUBUN = .cmbMEDGUBUN.value
		'strPROGNAME = .txtPROGNAME.value
		
		
		Dim strSUBSEQ
		Dim strSUBSEQNAME
		Dim strCLIENTSUBCODE
		Dim strCLIENTSUBNAME
		Dim strMATTERCODE
		Dim strMATTERNAME
		
		strSUBSEQ = .txtSUBSEQ.value 
		strSUBSEQNAME = .txtSUBSEQNAME.value 
		strCLIENTSUBCODE = .txtCLIENTSUBCODE.value 
		strCLIENTSUBNAME = .txtCLIENTSUBNAME.value 
		strMATTERCODE = .txtMATTERCODE.value 
		strMATTERNAME = .txtMATTERNAME.value 
		
		'strSUBSEQ,strSUBSEQNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strMATTERCODE,strMATTERNAME
		'msgbox strTRANSYEARMON & "번호" & strTRANSNO & "광고주코드" & strCLIENTCODE & "매체사코드" & strREAL_MED_CODE
		'exit sub
		IF strTRANSNO <> "" THEN
			'InitPageData
			IF not SelectRtn_HDR (strTRANSYEARMON, strTRANSNO, strCLIENTCODE) Then Exit Sub
			
			pnlTab1.style.visibility = "HIDDEN"
			pnlTab2.style.visibility = "visible"
						
			.txtDEMANDDAY.readOnly = "TRUE"
			.txtDEMANDDAY.className = "NOINPUT"	
			.imgCalDemandday.disabled = True		

			'쉬트 조회
			Call SelectRtn_DTL (strTRANSYEARMON, strTRANSNO, strCLIENTCODE)
			
			PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
			.txtCLIENTSUBNAME.value = "" 
			.txtCLIENTSUBCODE.value = ""
			.txtMATTERNAME.value = ""
			.txtMATTERCODE.value = ""
			.txtSUBSEQNAME.value = ""
			.txtSUBSEQ.value = ""		
		ELSE
			'InitPageData
			'미생성 내역
			vntData = mobjMELECSPONTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, strCLIENTCODE, strREAL_MED_CODE, strSUBSEQ,strSUBSEQNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strMATTERCODE,strMATTERNAME, strMEDGUBUN)
			
			pnlTab1.style.visibility = "visible"
			pnlTab2.style.visibility = "HIDDEN"
			
			.txtDEMANDDAY.readOnly = "FALSE"
			.txtDEMANDDAY.className = "INPUT"
			.imgCalDemandday.disabled = FALSE
			
			if not gDoErrorRtn ("SelectRtn") then
				if mlngRowCnt > 0 then
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   					PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
   				else
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   					InitPageData
   					PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
   				end if
   				Clean_Hdr
   			end if
		END IF
		DateClean
		AMT_SUM
   	end with
End Sub
Sub Clean_Hdr
with frmThis
.txtCLIENTNAME.value = ""
.txtAMT.value = ""
.txtVAT.value = ""
.txtSUMAMTVAT.value = ""
End with
End Sub
Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strCLIENTCODE )
	dim vntData
	on error resume next

	'초기화
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMELECSPONTRANS.Get_ELECTRANS_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 거래명세번호에 대하여" & meNO_DATA, ""
			Clean_Hdr
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			gWriteText "", "선택한 거래명세번호에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			
			SelectRtn_HDR = True
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByVal strYEARMON,ByVal strTRANSNO, ByVal strCLIENTCODE)
	Dim vntData
	on error resume next

	'초기화
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMELECSPONTRANS.Get_ELECTRANS_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_LIST") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG

		SelectRtn_DTL = True
		
		gWriteText "", "선택한 거래명세번호건의 상세내역에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE)
' strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtTRANSNO.value = strTRANSNO
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
	frmThis.txtMATTERCODE.value =strMATTERCODE
	frmThis.txtMATTERNAME.value = strMATTERNAME
	frmThis.txtSUBSEQ.value = strSUBSEQ
	frmThis.txtSUBSEQNAME.value = strSUBSEQNAME
	frmThis.cmbREAL_MED_CODE.value = strREAL_MED_CODE
	
	
	'frmThis.txtPROGNAME.value = strPROGNAME
	'frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
	
	
End Sub

'****************************************************************************************
' 거래명세서 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	with frmThis
		IF .sprSht1.MaxRows = 0 THEN
			gErrorMsgBox "삭제할 건의 상세내역이 없습니다.","삭제안내!"
			Exit Sub
		END IF
		
		For intCnt2 = 1 To .sprSht1.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",intCnt2) & " 에 대하여" &vbcrlf & "세금계산서번호가 존재하는 내역은 삭제가 되지 않습니다.","삭제안내!"
				Exit Sub
			End If
		Next
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0
		
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		'mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|PROGRAM |ADLOCALFLAG|WEEKDAY|DEPT_CD|DEMANDDAY|PRINTDAY| PRICE| CNT|AMT| VAT| MED_FLAG|ATTR02|TAXYEARMON|TAXNO|TRUST_SEQ")
		
		'선택된 자료를 끝에서 부터 삭제
		strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
	
		intRtn = mobjMELECSPONTRANS.DeleteRtn(gstrConfigXml,vntData, strTRANSYEARMON, strTRANSNO)

		IF not gDoErrorRtn ("DeleteRtn") then
			for i = .sprSht1.MaxRows to 1 step -1
				mobjSCGLSpr.DeleteRow .sprSht1,i
			next
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", strTRANSYEARMON & "-" & strTRANSNO & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht1
		initpagedata
		.txtDEMANDDAY.readOnly = false
		.txtDEMANDDAY.className = "INPUT"
		'SelectRtn
	End with
	err.clear	
End Sub
'-----------------------------------------------------------------------------------------
' 사업부코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTSUBCODE_POP

	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../MDCO/MDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' 코드명 표시
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME1.value = trim(vntRet(6,0))
			.txtMATTERNAME.focus()					' 포커스 이동
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTSUBNAME_onkeydown

	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME1.value = trim(vntData(6,0))
					
					.txtMATTERNAME.focus()
					gSetChangeFlag .txtCLIENTSUBCODE
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 브랜드코드팝업 버튼
'-----------------------------------------------------------------------------------------
'광고주 시퀀스가져오기
Sub ImgSUBSEQ_onclick
	Call SUBSEQCODE_POP()
End Sub


Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,440)
		if isArray(vntRet) then
			if .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			.txtSUBSEQ.value = trim(vntRet(1,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(2,0))	' 브랜드명 표시
			.txtCLIENTCODE.value = trim(vntRet(3,0))	' 광고주코드 표시
			.txtCLIENTNAME1.value = trim(vntRet(4,0))	' 광고주명 표시
			.txtCLIENTSUBCODE.value = trim(vntRet(7,0))	' 사업부코드 표시
			.txtCLIENTSUBNAME.value = trim(vntRet(8,0))	' 사업부명 표시
			
			.txtMATTERNAME.focus()					' 포커스 이동
			gSetChangeFlag .txtSUBSEQ		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value))
			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(1,0))
					.txtSUBSEQNAME.value = trim(vntData(2,0))
					.txtCLIENTCODE.value = trim(vntRet(3,0))		' 광고주코드 표시
					.txtCLIENTNAME1.value = trim(vntRet(4,0))	' 광고주명 표시
					.txtCLIENTSUBCODE.value = trim(vntRet(7,0))		' 사업부코드 표시
					.txtCLIENTSUBNAME.value = trim(vntRet(8,0))	' 사업부명 표시
					.txtMATTERNAME.focus()
				Else
					Call SUBSEQCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 소재코드팝업 버튼
'-----------------------------------------------------------------------------------------
'광고주 시퀀스가져오기
Sub ImgMATTER_onclick
	Call MATTERCODE_POP()
	
End Sub

Sub MATTERCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		
		vntInParams = array(trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("MDCMMATTERPOP.aspx",vntInParams , 783,473)
		if isArray(vntRet) then
			if .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			.txtMATTERCODE.value = trim(vntRet(0,0))	
			.txtMATTERNAME.value = trim(vntRet(1,0))	
			
			.txtCLIENTSUBCODE.value = trim(vntRet(4,0))
			.txtCLIENTSUBNAME.value =  trim(vntRet(5,0))
			.txtSUBSEQ.value = trim(vntRet(6,0))
			.txtSUBSEQNAME.value = trim(vntRet(7,0))
			
			'gSetChangeFlag .txtMATTERCODE
			
			.txtCLIENTCODE.value = trim(vntRet(2,0))
			.txtCLIENTNAME1.value = trim(vntRet(3,0))
     	end if
	End with
	'gSetChange
	
End Sub

Sub txtMATTERNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			'vntData = mobjMDCMCODETR.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtMATTERCODE.value),trim(.txtPROGNAME.value))
			vntData = mobjMDCMCODETR.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value),"", "")
			if not gDoErrorRtn ("GetMATTER") then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntData(0,1))		' 브랜드 표시
					.txtMATTERNAME.value = trim(vntData(1,1))	' 브랜드명 표시 2,3,6,7
					.txtCLIENTCODE.value = trim(vntData(2,1))
					.txtCLIENTNAME1.value = trim(vntData(3,1))
					.txtCLIENTSUBCODE.value = trim(vntData(4,1))
					.txtCLIENTSUBNAME.value =  trim(vntData(5,1))
					.txtSUBSEQ.value = trim(vntData(6,1))
					.txtSUBSEQNAME.value = trim(vntData(7,1))
				Else
					Call MATTERCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'번호를 클리어한다.
Sub CleanField (objField1, objField2, objField3)
	if isobject(objField1) then objField1.value = ""
	if isobject(objField2) then objField2.value = ""
	if isobject(objField3) then objField3.value = ""
	'InitPageData
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
				border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;위수탁 거래명세서 관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 342px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start--></TD>
							</TR>
							<!--Top Define Table End-->
							<!--Input Define Table End--></TABLE>
						<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)"
												width="90">년&nbsp;&nbsp;월</TD>
											<TD class="SEARCHDATA" style="WIDTH: 347px" width="347"><INPUT class="INPUT" id="txtTRANSYEARMON" title="거래명세년월" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="6" name="txtTRANSYEARMON"><IMG id="ImgTRU" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTRU"><INPUT class="INPUT" id="txtTRANSNO" title="거래명세번호" style="WIDTH: 72px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="6" name="txtTRANSNO"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="90">매체사</TD>
											<TD class="SEARCHDATA" style="WIDTH: 427px"><SELECT id="cmbREAL_MED_CODE" title="매체사" name="cmbREAL_MED_CODE">
													<OPTION value="B00140" selected>한국방송광고공사본사</OPTION>
													<OPTION value="B00144">한국방송광고공사부산지사</OPTION>
													<OPTION value="B00142">한국방송광고공사대구지사</OPTION>
													<OPTION value="B00143">한국방송광고공사대전지사</OPTION>
													<OPTION value="B00141">한국방송광고공사광주지사</OPTION>
													<OPTION value="B00145">한국방송광고공사전북지사</OPTION>
												</SELECT>
												<SELECT id="cmbMEDGUBUN" title="매체구분" style="WIDTH: 80px" name="cmbMEDGUBUN">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="TV">TV</OPTION>
													<OPTION value="RD">RADIO</OPTION>
													<OPTION value="DMB">DMB</OPTION>
												</SELECT>
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME1, txtCLIENTCODE, txtTRANSNO) ">광고주
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="광고주코드조회" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)">사업부
											</TD>
											<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="사업부코드명" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT" id="txtCLIENTSUBCODE" title="사업부명" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtCLIENTSUBCODE"></TD>
										</tr>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME, txtSUBSEQ)">브랜드&nbsp;
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="브랜드코드명" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQ" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgSUBSEQ"><INPUT class="INPUT" id="txtSUBSEQ" title="브랜드명조회" style="HEIGHT: 22px" type="text" maxLength="6"
													align="left" size="5" name="txtSUBSEQ"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME, txtMATTERCODE)">소재명
											</TD>
											<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtMATTERNAME" title="소재코드" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtMATTERNAME"><IMG id="ImgMATTER" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgMATTER"><INPUT class="INPUT" id="txtMATTERCODE" title="소재코드명" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtMATTERCODE"></TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"><FONT face="굴림"></FONT></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;거래명세서 생성</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="imgAllSet" onmouseover="JavaScript:this.src='../../../images/imgAllSetOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAllSet.gIF'"
																height="20" alt="해당년월 의 거래명세서를 전체 생성합니다." src="../../../images/imgAllSet.gIF" border="0"
																name="imgAllSet"></td>
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgTransCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTransCre.gIF'"
																height="20" alt="선택된 거래명세서 내역을 생성합니다." src="../../../images/imgTransCre.gIF" border="0"
																name="imgSave"></td>
														<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
									<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="right" border="0">
										<TR>
											<TD class="LABEL" width="90">광고주</TD>
											<TD class="DATA" width="256"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="37" name="txtCLIENTNAME">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"
												width="90">발행일자</TD>
											<TD class="DATA" width="257"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="담당부서명" style="WIDTH: 213px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="30" name="txtPRINTDAY"><IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
													name="imgCalPrintday">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')"
												width="90">청구일자</TD>
											<TD class="DATA" width="257"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="브랜드명" style="WIDTH: 210px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="29" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></TD>
										</TR>
										<TR>
											<TD class="LABEL">공급가액</TD>
											<TD class="DATA"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="광고금액" style="WIDTH: 257px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="37" name="txtAMT">
											</TD>
											<TD class="LABEL">부가세액</TD>
											<TD class="DATA"><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="부가세" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="37" name="txtVAT"></TD>
											<TD class="LABEL">계</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMTVAT" class="NOINPUT_R" id="txtSUMAMTVAT" title="계" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="37" name="txtSUMAMTVAT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start-->
				<TR>
					<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 604px" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 1040px; HEIGHT: 580px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
								<PARAM NAME="_ExtentY" VALUE="15346">
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
							<OBJECT id="sprSht_SUM" style="WIDTH: 1038px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27464">
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
						<DIV id="pnlTab2" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht1" style="WIDTH: 1040px; HEIGHT: 604px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
								<PARAM NAME="_ExtentY" VALUE="15981">
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
						<DIV id="pnlTab3" style="POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout"><!--VISIBILITY: hidden;-->
							<OBJECT id="sprShtAll" style="WIDTH: 1040px; HEIGHT: 0px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
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
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
