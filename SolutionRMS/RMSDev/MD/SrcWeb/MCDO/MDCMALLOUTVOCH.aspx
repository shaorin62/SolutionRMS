<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMALLOUTVOCH.aspx.vb" Inherits="MD.MDCMALLOUTVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 전표생성</title>
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
'HISTORY    :1) 2009/11/24 By 황덕수
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
Dim mlngRowCnt,mlngColCnt
Dim mobjMDCOALLOUTVOCH
Dim mobjMDCOGET
Dim mobjMDCOVOCH
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim mstrPROCESS
Dim mstrSTAY

mstrSTAY = TRUE

mstrGUBUN = "O"
mstrPROCESS = ""
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

'강제삭제 버튼 숨기기
Sub Set_delete(byVal strmode)
	With frmThis
		IF .rdT.checked = TRUE then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'조회버튼
Sub imgQuery_onclick
	If frmThis.txtYEARMON.value = "" Then
		gErrorMsgBox "조회년월을 입력하시오","조회안내"
		exit Sub
	End If

	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'엑셀버튼 클릭
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		
		if mstrGUBUN = "O"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_OUT
		end if
		
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'전표생성 클릭
Sub ImgvochCre_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'전표삭제 클릭
Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'오류전표삭제클릭
Sub ImgErrVochDel_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'전표강제 삭제 클릭
Sub imgVochDelco_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'--적용버튼 클릭
Sub ImgSUMMApp_onclick()
	Dim intRtn
	
	with frmThis
		if .cmbSETTING.value = "" then
			gErrorMsgBox "적용하실 컬럼 명이 없습니다. ","적용오류"
			exit sub
		end if 
		
		'취급액
		if mstrGUBUN = "O"  then 
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUT)
		end if
	End With
End Sub

sub settingRowChange(sprsht)
	Dim strSETTINGDATA
	Dim intCnt 
	Dim i ,j

	with frmThis
		intCnt = 0
		
		for j = 1 to sprsht.MaxRows
			If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",j) = "1" Then
				intCnt = intCnt + 1
			End if 
		next
		
		if intCnt = 0 Then
			gErrorMsgBox "체크된 데이터가 없습니다. 적용하실 데이터를 체크하세요. ","적용오류"
			EXIT SUB
		End if
		
		strSETTINGDATA = ""
		strSETTINGDATA = .txtSUMM.value
		
		for i = 1 to sprsht.MaxRows
			If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",i) = "1" Then
				mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
			End if
		next 
	End with
end sub


'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'이미지버튼 클릭시
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEXCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtEXCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEXCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(0,0))
					.txtEXCLIENTNAME.value = trim(vntData(1,0))
				Else
					Call EXCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'완료체크
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'미완료체크
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'에러체크
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUMM_onchange
	Dim blnByteCHk
	Dim intRtn
	blnByteCHk =  checkBytes(frmThis.txtSUMM.value)
	
	If blnByteCHk  > 23 Then
		intRtn = gYesNoMsgbox("적요의 크기는 23Byte 를 넘을수 없습니다. 초기화 하시겠습니까?","처리안내!")
		
		IF intRtn <> vbYes then exit Sub
		
		frmThis.txtSUMM.value = ""
	End If
End Sub

function checkBytes(expression)
	Dim VLength
	Dim temp
	Dim EscTemp
	Dim i
	
	VLength=0
	temp = expression
	
	if temp <> "" then
		for i=1 to len(temp) 
			if mid(temp,i,1) <> escape(mid(temp,i,1))  then
				EscTemp=escape(mid(temp,i,1))
				if (len(EscTemp)>=6) then
					VLength = VLength +2
				else
				VLength = VLength +1
				end if
			else
				VLength = VLength +1
			end if
		Next
	end if

	checkBytes = VLength
end function

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,false
			End If
		End if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"paycode") then
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUT,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUT)
		end if
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

'-----------------------------------
' SpreadSheet 클릭
'-----------------------------------
Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

'-----------------------------------
' SpreadSheet 더블 클릭
'-----------------------------------
sub sprSht_OUT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUT, ""
		end if
	end with
end sub

'----------------------------------------------------------
'시트 자동 계산 [시트 키업]
'----------------------------------------------------------
Sub sprSht_OUT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_OUT
	End With
End Sub

SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"VAT") Then
		
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"VAT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

'---------------------------------------------
'시트 마우스 업
'---------------------------------------------
'대대행 
Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_OUT
	end with
End Sub

'-----------------------------------
'시트에서 마우스를 금액합산 이벤트
'-----------------------------------
sub MouseUp_SumAmt(sprSht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i,j
	Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"AMT") or SprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"VAT") then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intGBN
	Dim strComboPREPAYMENT
	Dim strBMORDER
	
	'서버업무객체 생성	
	Set mobjMDCOALLOUTVOCH	= gCreateRemoteObject("cMDCO.ccMDCOALLOUTVOCH")
	Set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	Set mobjMDCOVOCH	 = gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
	
    with frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strBMORDER = "AD0110" & vbTab & "AD0120" & vbTab & "AD0130" & vbTab & "AD0140" & vbTab & "AD0150" & vbTab & "AD0160" & vbTab & "AD0190" _
					& vbTab & "AD0210" & vbTab & "AD0220" & vbTab & "AD0290" & vbTab & "AD0310" & vbTab & "AD0320" & vbTab & "AD0390" & vbTab & "AD0410" _ 
					& vbTab & "AD0420" & vbTab & "AD0430" & vbTab & "AD0440" & vbTab & "AD0450" & vbTab & "AD0510" & vbTab & "AD0610" & vbTab & ""
		
		'**************************************************
		'대대행 시트 디자인
		'**************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 42, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK | MEDFLAGNAME | CLIENTCODE | RMSNO | EXCLIENTCODE | INPUT_MEDFLAG | DEPT_CD | GUBUN"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "선택|전표일자|광고주|거래처코드|거래처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|지급기일|입금기일|상대VENDOR|구분|차변계정|계정|BMORDER|증빙일|지급방법|BANKTYPE|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|MEDFLAG|AMTGBN|TRANSRANK|매체구분|CLIENTCODE|RMSNO|EXCLIENTCODE|INPUT_MEDFLAG|DEPT_CD|GUBUN"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1",  "   4|       8|    15|        10|    15|  17|       5|         8|  10|    10|       6| 5|       8|       8|        10|   0|       7|   7|      7|     8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      0|     0|       10|       6|         0|    0|           0|            0|      0|    0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | SEMU | BP | VENDOR | GBN | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | GUBUN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR | ACCOUNT | DEBTOR | MEDFLAGNAME",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG|SUMMTEXT",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BP | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAGNAME | GUBUN"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN  | GFLAG | MEDFLAG | DUEDATE | ERRCODE | AMTGBN | CLIENTCODE | RMSNO | EXCLIENTCODE | INPUT_MEDFLAG | DEPT_CD | GUBUN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_OUT,"VOCHNO | ERRCODE | ERRMSG"
	End with
	
	pnlFLAG.style.visibility = "visible" 
	pnlTab_out.style.visibility = "visible" 
	'화면 초기값 설정
	InitPageData	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtEDYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet초기화

		.sprSht_OUT.MaxRows = 0
		
		.txtYEARMON.focus()
		
		Get_COMBO_VALUE
		
		'처음에 강제 삭제 감춤
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"		
	End with
End Sub

Sub EndPage()
	set mobjMDCOALLOUTVOCH = Nothing
	Set mobjMDCOGET = Nothing
	Set mobjMDCOVOCH = Nothing
	gEndPage	
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht_OUT.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCOALLOUTVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUT, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	

'-----------------------------------------------------------------------------------------
' 그리드 서브 콤보 설정
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strCODE, row, sprsht)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCODE = replace(strCODE,"-","")

       	vntData = mobjMDCOALLOUTVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
		If not gDoErrorRtn ("Get_SUBCOMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 sprsht, "BANKTYPE",Row,Row,vntData,,160 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
Sub SelectRtn (strVOCH_TYPE)
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
	
	with frmThis
		.sprSht_OUT.MaxRows = 0

		CALL SelectRtn_OUT()

		mstrSTAY = TRUE
   	end with
End Sub

Sub SelectRtn_OUT ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
    Dim strEDYEARMON
    Dim strTAXNO
	
	with frmThis
		.sprSht_OUT.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strEDYEARMON = .txtEDYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strEXCLIENTCODE = .txtEXCLIENTCODE.value
		strEXCLIENTNAME = .txtEXCLIENTNAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		vntData = mobjMDCOALLOUTVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, strEDYEARMON, _
												   strCLIENTCODE, strCLIENTNAME, _
												   strGBN, strEXCLIENTCODE, strEXCLIENTNAME)

		if not gDoErrorRtn ("SelectRtn_OUT") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				strTAXNO = ""
				
				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DUEDATE",intCnt,intCnt,false
						
						
						'코바코 데이터와 미디어 크리에트 데이터를 따로 나누지 않은 이전 데이터 중복을 막기위해 임시로 입력 완료된 데이터만 입력함
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GUBUN",intCnt) <> "" then
							if strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) then
								mobjSCGLSpr.DeleteRow .sprSht_OUT,intCnt
							end if 
							strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)
						
						end if
						
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DUEDATE",intCnt,intCnt,false
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,false
					End If
				Next
				
				AMT_SUM .sprSht_OUT
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

Function DataValidation_OUT ()
	DataValidation_OUT = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_OUT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUT = True
End Function

Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn
	
	with frmThis
		IF mstrPROCESS = "Create" THEN
			IF NOT .rdF.checked THEN
				gErrorMsgBox "미완료조회시 가능합니다.","생성및삭제"
				exit sub
			end IF 
		end if 
		IF mstrPROCESS = "Delete" THEN
			IF NOT .rdT.checked THEN
				gErrorMsgBox "완료조회시 가능합니다.","생성및삭제"
				exit sub
			end IF 
		end if 
		
		IF mstrSTAY THEN 
			mstrSTAY = FALSE
			if DataValidation_OUT =false then exit sub
			CALL ProcessRtn_OUT()
		ELSE
			gErrorMsgBox "전표처리 진행중입니다.","전표처리 안내"
		END IF
   	end with
End Sub

'대대행 저장 로직
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim strCUSTOMERCODE
	Dim intColFlag, bsdiv, intMaxCnt
	
	'전표 채번을 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK | MEDFLAGNAME | CLIENTCODE | RMSNO | EXCLIENTCODE | INPUT_MEDFLAG | DEPT_CD | GUBUN")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		
		if mstrPROCESS = "Create" then
			if not UpdateRtn_OUT_Medium (vntData_ProcesssRtn) then 
				gErrorMsgBox "청약데이터의 임시 번호가 저장되지 않았습니다. 관리자에게 문의 부탁 드립니다. " & meNO_DATA,"저장취소"
				exit Sub
			end if
		End if
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그
		
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0012"
		
		
		'최대값
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		IF .rdDIV.checked THEN
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
						
						'채번을 설정한다.
						'--------------------------------------------------------------------------------------

						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" :  strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'생성 저장한 RMS 채번 가져오기
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)
						'---------------------------------------------------------------------------------------
						
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "O"

						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt) 
						else
							
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt) 
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
			
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt) 
						else
							strHSEQ = strHSEQ + 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt) 
						end if
					end if 
				Next
			end if 
		ELSE
			if mstrPROCESS = "Create" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M" 
	                
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'청구합계
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								'청구합계,부가세합계,청구지는 헤더에 변수로 저장
								If intCnt2 = intCnt Then
								Else
								
									'채번을 설정한다.(합산전표의 채번 설정)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",intCnt)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'생성 저장한 RMS 채번 가져오기
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1

									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE = "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
									
									if strIF_CNT = "1" then
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									else
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									
									end if
												
												
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",j)
															
															
											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								End If
								'무한업데이트.
							End If
						End If
					Next
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
	                
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'청구합계
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								'청구합계,부가세합계,청구지는 헤더에 변수로 저장
								If intCnt2 = intCnt Then
								Else
									strIF_CNT = strIF_CNT + 1
									
									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE = "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
									
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
												
												
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",j)
															
															
											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								End If
								'무한업데이트.
								intCnt2 = intCnt
							End If
						End If
					Next
				Next
			end if 
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Function UpdateRtn_OUT_Medium (vntData)
	UpdateRtn_OUT_Medium = False
   	Dim intRtn
	With frmThis
	
		intRtn = mobjMDCOALLOUTVOCH.UpdateRtn_OUT_Medium(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			if intRtn = 0 then
				exit Function
			End if		
   		End If
   	end With
   	UpdateRtn_OUT_Medium = True
End Function

'****************************************************************************************
' 채번 설정처리
'****************************************************************************************
Function InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
	InsertRtn_VOCHNO = false
   	Dim strVOCHNO
	With frmThis
		
		'채번을 설정& 저장한다 (저장은 중복을 막고 SAP 쪽에서 에러가 날 경우에도 같은 번호로 생성되는 것을 막는다.).
		intRtn = mobjMDCOVOCH.InsertRtn_VOCHNO(gstrConfigXml,strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
		If not gDoErrorRtn ("InsertRtn_VOCHNO") Then
		
			If intRtn = 0 Then
				Exit Function
			End If		
   		End If
   	end With
   	InsertRtn_VOCHNO = true
End Function

'---------------------------------------------------
' 전표상태 및 전표번호 받아오기 및 실제 RMS업데이트
'---------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
	
		if mstrPROCESS ="Create" then
			IF mstrGUBUN = "O" THEN
				intRtn = mobjMDCOALLOUTVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "OUTALL")
			END IF

			if not gDoErrorRtn ("ProcessRtn") then
				'모든 플래그 클리어
				IF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "전표가 생성되었습니다.","저장안내"
				else
					gErrorMsgBox "에러가 발생했습니다.","저장안내"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   		elseif mstrPROCESS ="Delete" then
   			IF mstrGUBUN = "O" THEN
				intRtn = mobjMDCOALLOUTVOCH.VOCHDELL_BUY(gstrConfigXml, strRETURNLIST, mstrGUBUN, "OUT", "OUTALL")
			END IF
			
   			if not gDoErrorRtn ("VOCHDELL") then
				'모든 플래그 클리어
				IF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF
				
				gErrorMsgBox "전표가 삭제되었습니다.","저장안내"
				
				SelectRtn(mstrGUBUN)
   			end if
   		end if 
   		IF mstrGUBUN = "O" THEN
			.sprSht_OUT.focus()
		END IF
	End With
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   	
		IF NOT .rdE.checked THEN
			gErrorMsgBox "오류조회시 가능합니다.","생성및삭제"
			exit sub
		end if 
		
		IF mstrGUBUN = "O" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"삭제취소"
			exit sub
		End If
		
		intRtn = mobjMDCOALLOUTVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'모든 플래그 클리어
			IF mstrGUBUN = "O" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
			END IF
			
			if intRtn > 0 Then
				gErrorMsgBox "오류 전표가 삭제되었습니다.","저장안내"
			End If
			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

'-----------------------------------------
'전표 강제 삭제
'-----------------------------------------
Sub DeleteRtn (strGUBUN)
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strVOCHNO
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
	
		If mstrGUBUN = "O"  then
			If .sprSht_OUT.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUT.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if
		END IF
	
		intRtn = gYesNoMsgbox("강제삭제는 SAP에서 승인된 전표를 SAP에서 취소하여 RMS쪽에서 삭제할 수 없을때 RMS쪽 전표를 강제로 삭제할때 사용합니다. " & vbCrlf & "  " & vbCrlf & " 전표를 강제로 삭제하시겠습니까?","강제삭제 확인")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		If mstrGUBUN = "O"  then
			for i = .sprSht_OUT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
					
					intRtn = mobjMDCOALLOUTVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "OUT", "OUTALL")
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		END IF
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
			SelectRtn (strGUBUN)
	End With
	err.clear	
End Sub

		</script>
		<script language="javascript">
		//##########################################################################################################################################
		//******************************************주1) frmSapCon 아이 프레임 을 이용하여 Submit 하는 함수
		//##########################################################################################################################################

		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {
			//헤더
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			//dtl 
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			
			window.frames[0].document.forms[0].submit();
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
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
											<td class="TITLE">매체 대대행&nbsp;전표관리&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
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
						<TABLE id="tblBody" style="WIDTH: 100%" height="93%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
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
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')"
												width="60">&nbsp;년월
											</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtEDYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" size="9" name="txtEDYEARMON"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="60">&nbsp;광고주
											</TD>
											<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 142px; HEIGHT: 22px"
													type="text" maxLength="100" size="16" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="60">대대행사
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="코드명" style="WIDTH: 142px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="18" name="txtEXCLIENTNAME">&nbsp;<IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">&nbsp;<INPUT class="INPUT" id="txtEXCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtEXCLIENTCODE">
											</TD>
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
										<TR>
											<TD class="SEARCHLABEL">발행
											</TD>
											<TD class="SEARCHDATA" colSpan="6"><INPUT id="rdT" title="완료내역조회" type="radio" value="rdT" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;완료&nbsp;
												<INPUT id="rdF" title="미완료 내역조회" type="radio" value="rdF" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')"
													CHECKED>&nbsp;미완료&nbsp; <INPUT id="rdE" title="오류전표 내역조회" type="radio" value="rdE" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;오류&nbsp;
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<td class="DATA">
									합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 20px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
									<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 20px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="HEIGHT: 26px" align="left" width="100%">
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
													<TR>
														<td><IMG id="ImgvochCre" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																height="20" alt="전표를 저장합니다." src="../../../images/ImgvochCre.gIF" border="0" name="ImgvochCre"></td>
														<td><IMG id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'"
																height="20" alt="전표를 삭제합니다." src="../../../images/imgVochDel.gIF" border="0" name="imgVochDel"></td>
														<td><IMG id="ImgErrVochDel" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																height="20" alt="오류전표 를 삭제합니다." src="../../../images/ImgErrVochDel.gIF" border="0"
																name="ImgErrVochDel"></td>
														<td><IMG id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																height="20" alt="전표를 강제로 삭제합니다." src="../../../images/imgVochDelco.gIF" border="0"
																name="imgVochDelco" title="SAP에서 직접삭제하여 RMS에서 삭제할 수 없을때 RMS전표를 강제로 삭제한다."></td>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
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
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHDATA" width="90" onclick="vbscript:Call gCleanField(txtSUMM,'')">
												<select id="cmbSETTING" style="WIDTH: 90px">
													<OPTION value="" selected>선택</OPTION>
													<OPTION value="POSTINGDATE">전표일자</OPTION>
													<OPTION value="SUMM">적요</OPTION>
													<OPTION value="BA">사업영역</OPTION>
													<OPTION value="COSTCENTER">코스트센터</OPTION>
													<OPTION value="SEMU">세무코드</OPTION>
													<OPTION value="DEMANDDAY">지급기일</OPTION>
													<OPTION value="DUEDATE">입금기일</OPTION>
													<OPTION value="DEBTOR">차변계정</OPTION>
													<OPTION value="ACCOUNT">계정</OPTION>
													<OPTION value="PREPAYMENT">선수금구분</OPTION>
													<OPTION value="SUMMTEXT">본문TEXT</OPTION>
												</select></TD>
											<TD class="DATA"><INPUT id="txtSUMM" title="적요적용" style="WIDTH: 368px; HEIGHT: 21px" type="text" size="56"
													name="txtSUMM" class="INPUT_L"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													title="적요를 일괄 적용합니다" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20"
													alt="적요를 일괄 적용합니다" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0" name="ImgSUMMApp">
												<DIV id="pnlFLAG" align="center" style="VISIBILITY: hidden; WIDTH: 250px; POSITION: absolute; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdDIV" title="분할" type="radio" CHECKED value="rdDIV" name="rdDIVGUBUN">&nbsp;분할&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdSUM" title="합산" type="radio" value="rdSUM" name="rdDIVGUBUN">&nbsp;합산</DIV>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 80%" vAlign="top" align="center">
									<DIV id="pnlTab_out" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_OUT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="12700">
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
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
				</TR>
			</TABLE>
			<P>
				<!--Input Define Table End--> </TD></TR> 
				<!--Top TR End--> </TABLE> 
				<!--Main End--></P>
		</FORM>
		</TR></TABLE><iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 100%; HEIGHT: 300px" name="frmSapCon"
			src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
