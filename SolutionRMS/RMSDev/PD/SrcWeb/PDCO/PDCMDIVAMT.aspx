<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDIVAMT.aspx.vb" Inherits="PD.PDCMDIVAMT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>청구내역 분할처리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/제작관리번호 등록 화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMDIVAMT.aspx
'기      능 : 견적내역확정분에 대한 제작관리번호 분할 처리 될수 있도록 처리
'파라  메터 : 
'특이  사항 : 해당 하나의 청구처 를 다중으로 분할하며, 하나의 좝번호에 부번호를 부여한다.
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/19 By Kim Tae Ho
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
Dim mobjPDCMDIVAMT
Dim mobjPDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag

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

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'서버업무객체 생성	
	Set mobjPDCMDIVAMT = gCreateRemoteObject("cPDCO.ccPDCODIVAMT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'화면 초기값 설정
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	Dim strComboList 
	strComboList =  "사용" & vbTab & "미사용"
	gSetSheetDefaultColor
with frmThis
		
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht,    "PREESTNO|YEARMON|JOBNO|JOBNAME|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|DIVAMT|ADJAMT|INYN|CREDAY|INYNNM"
		mobjSCGLSpr.SetHeader .sprSht,		    "견적번호|합의월|JOBNO|JOB명|광고주코드|광고주명|사업부코드|사업부명|견적확정금액|청구금액|완료구분|견적합의일|입력구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "10      |10    |10   |22   |0         |18      |0         |18      |12          |12      |0       |0         |10"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|CLIENTSUBCODE|CREDAY|INYN", true
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PREESTNO|YEARMON|JOBNO|INYN|INYNNM",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|CLIENTSUBNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PREESTNO|YEARMON|JOBNO|JOBNAME|DIVAMT|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|INYNNM"
		.sprSht.MaxRows = 1
		
		'**************************************************
		'***상세내역 Sheet 디자인
		'**************************************************	
			
        gSetSheetColor mobjSCGLSpr, .sprSht1 
		mobjSCGLSpr.SpreadLayout .sprSht1, 18, 0
		mobjSCGLSpr.AddCellSpan  .sprSht1, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,10, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT|ATTR02"
		mobjSCGLSpr.SetHeader .sprSht1,         "견적번호|순번|제작번호|년월|견적합의일|브랜드|브랜드명|사업부|사업부명|광고주|광고주명|분할금액|JOB명|청구금액|청구년월"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "0       |0   |0       |0   |10        |6   |2|14      |6     |2|18    |6     |2|18    |10      |28   |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN0"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN2"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "CREDAY", -1, -1, 10
		mobjSCGLSpr.ColHidden .sprSht1, "PREESTNO|SEQ|JOBNO|YEARMON|ADJAMT|ATTR02", true
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|JOBNAME|SUBSEQ|SUBSEQNAME", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "DIVAMT|ADJAMT", -1, -1, 0
		'**************************************************
		'***상세내역 Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprShtSum
		mobjSCGLSpr.SpreadLayout .sprShtSum, 18, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT|ATTR02"
		mobjSCGLSpr.AddCellSpan  .sprShtSum, 2, 1, 2, 1
		mobjSCGLSpr.SetText .sprShtSum, 2, 1, "합 계"
		mobjSCGLSpr.SetScrollBar .sprShtSum, 0
		mobjSCGLSpr.SetBackColor .sprShtSum,"1|2",rgb(205,219,215),false
		mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "DIVAMT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON|ATTR02", true
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
		mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "15"
		.sprSht1.focus
			
		.txtPREESTNO.style.visibility = "hidden"
		.txtYEARMONPOP.style.visibility = "hidden"
		.txtCREDAY.style.visibility = "hidden"
		.txtJOBNOPOP.style.visibility = "hidden"
		.txtDIVAMT.style.visibility = "hidden"
		
	End with

	
	'pnlTab1.style.visibility = "visible" 
	
End Sub
'-----------------------------------------------------------------------------------------
' 상세내역 및 내역합계 그리드 Change 시 처리
'-----------------------------------------------------------------------------------------
'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht1_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
	End with
end sub
'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub


Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
				
		vntData = mobjPDCMDIVAMT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtJOBNAME.value,.txtJOBNO.value,.cmbYN.value)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
   				Call sprSht_Click(1,1)
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_NAME",Row)))
			vntRet = gShowModalWindow("MDCMDEPTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtDEPTCODE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		
		end if
	End with
End Sub
'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)

end sub

'여기까지 쉬트 버튼 클릭

'Validation
Function DataValidation ()
	DataValidation = false	
	With frmThis
		'IF not gDataValidation(frmThis) then exit Function	
	End With
	DataValidation = True
End Function
'저장로직


Sub EndPage()
	set mobjPDCMDIVAMT = Nothing
	set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0
	.txtYEARMON.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2)
	End with
	
End Sub

sub DeleteRtn

	
End Sub
'-----------------------------------------------------------------------------------------
' JOB 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'실제 데이터List 가져오기
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtJOBNO.value = trim(vntRet(0,0))  ' Code값 저장
			.txtJOBNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO.value = trim(vntData(0,0))
					.txtJOBNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master 입력 데이터 Validation : 필수 입력항목 검사
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  		If mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",1) = "" Then
  			gErrorMsgBox "첫번째 행의 제작건명은 반드시 입력하셔야 합니다.","입력오류"
  			Exit Function
  		End if
  		for intCnt = 1 to .sprSht1.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " 번째 행의 광고주코드를 확인하십시오","입력오류"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTSUBCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " 번째 행의 사업부코드를 확인하십시오","입력오류"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",intCnt) = 0 Then 
					gErrorMsgBox intCnt & " 번째 행의 분할금액을 확인하십시오","입력오류"
					Exit Function
			 End if
		next
		
   	End with
	DataValidation = true
End Function
Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, .txtJOBNOPOP.value 		
		mobjSCGLSpr.SetTextBinding .sprSht,"CREDAY",.sprSht.ActiveRow, .txtCREDAY.value  
	End with
End Sub
sub imgAddRow_onclick ()
	With frmThis
		call sprSht1_Keydown(meINS_ROW, 0)
	End With 
end sub
sub imgDelRow_onclick ()
	With frmThis
		call sprSht1_Keydown(meDEL_ROW, 0)
	End With 
end sub

Sub sprSht1_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 13 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub
'저장처리
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strJOBNO,strDEMANDAMT,strJOBYEARMON
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	
	with frmThis
   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		'회의결과 달라도 저장될수 있음.. 분담금액이 청구금액보다 크다면 에러,,
		'만약 작다면 바로저장 청구금액이 예산에서 삭제 또는 삭감 되면 기존 분담 PD_GROUP_DIVAMT 의 내역 삭제 
		If CDBL(.txtDIVAMT.value) < strSUMDEMANDAMT Then
   			msgbox "분할금액의 합은 청구금액을 넘을수 없습니다."
   			Exit Sub
   		End IF
		
		'제작건명 처음의 로우와 일치 시키기
		For intCnt2 = 2 To .sprSht1.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",intCnt2) = "" Then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",intCnt2, mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",1)  
			end if
		Next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|DIVAMT|JOBNAME")
		
		if .sprSht1.MaxRows = 0 Then
			MsgBox "디테일 데이터를 입력 하십시오"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		intRtn = mobjPDCMDIVAMT.ProcessRtn(gstrConfigXml,vntData,.txtCUSTCODEHRD.value )
	
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			strRow = .sprSht.ActiveRow
			SelectRtn
			Call sprSht_Click(1,strRow)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		end if
   		
   	end with
End Sub

'삭제로직
Sub DeleteRtn_DTL
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	dim strJOBNO,strCUST,strSEQ
	Dim lngSUMAMT,lngSUMAMT2
	Dim strPREESTNO
	Dim dblSEQ
	Dim strRow
	Dim strGUBN
	'On error resume next
	
	with frmThis
		'한 건씩 삭제할 경우
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)

		if gDoErrorRtn ("DeleteRtn_Dtl") then exit sub

		if intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		
		strJOBNO = ""
		strCUST = ""
		strSEQ = 0
		lngSUMAMT = 0
		lngSUMAMT2 = 0
		'합계가 맞는지 여부검사
		'현재저장되어 있는 금액
		
		strGUBN = ""
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			strJOBNO = Trim(.txtJOBNOPOP.value) 
			strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i))	
			dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))	
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			if cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "" AND cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "1" then
				If cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"ATTR02",vntData(i))) <> "" Then
					gErrorMsgBox "거래명세서 작성내역은 삭제될수 없습니다.","삭제오류"
					Exit Sub
				End If
				intRtn = mobjPDCMDIVAMT.DeleteRtn(gstrConfigXml,strJOBNO,strPREESTNO,dblSEQ)
				strGUBN = "T"
			Elseif cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) = "1" Then
				gErrorMsgBox "최초생성 견적내역은 삭제될수 없습니다.","삭제오류"
				Exit Sub
			Elseif cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "" Then
				strGUBN = "F"
			end if
			
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				'합계재계산
				
   			end if
		next
		'ProcessRtn
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht1
		mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
		'gWriteText lblStatus,"자료가 삭제" & mePROC_DONE
		If strGUBN = "T" Then
			strRow = .sprSht.ActiveRow
			SelectRtn
			Call sprSht_Click(1,strRow)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		End If
		
	end with
End Sub
sub SelectRtn_DTL ()
   	Dim vntData
   	Dim i, strCols
	Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMDIVAMT.SelectRtn_DIV(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNOPOP.value)

		if not gDoErrorRtn ("SelectRtn_DIV") then
			mobjSCGLSpr.SetClipBinding .sprSht1, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			If mlngRowCnt < 1 Then
			frmThis.sprSht1.MaxRows = 0 
			
			Else
				'거래명세서 작성분 에 한하여 수정 불가능 하도록 처리하였음
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht1,"ATTR02",intCnt) = "" Then
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
						mobjSCGLSpr.SetCellsLock2 .sprSht1,false,intCnt,-1,-1,true 
					Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						
						mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,-1,-1,true
					End If
				Next
			End If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht1
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			'Call SUM_AMT ()
   		end if
   	end with
end sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	with frmThis
		.txtPREESTNO.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row)
		.txtYEARMONPOP.value = mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",Row)
		.txtJOBNOPOP.value = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row)
		.txtCREDAY.value = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDAY",Row)
		.txtDIVAMT.value =mobjSCGLSpr.GetTextBinding( .sprSht,"DIVAMT",Row)
		SelectRtn_DTL
		SUM_AMT
	End with
End Sub
'------------------------------------------
' 상세내역 그리드 처리
'------------------------------------------
Sub sprSht1_change(ByVal Col,ByVal Row)
	
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName,strCodeName2
   	Dim strQTY,strPRICE,strAMT 
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		IF  Col = 11 Then
			
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)
			
			if not gDoErrorRtn ("GetCUSTNO_HIGHCUSTCODE") then
			
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(5,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(6,0)			
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+4,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 10, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
   		ElseIF  Col = 14 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(1,0)			
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 13, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
   		ElseIF  Col = 8 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)

			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntData(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntData(3,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntData(4,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(7,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(8,0)		
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+7,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 7, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
		end if
   	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col,Row
	SUM_AMT
End Sub	

Sub sprSht1_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 16 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub

Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht1,"PREESTNO",.sprSht1.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",.sprSht1.ActiveRow, .txtJOBNOPOP.value 		
		mobjSCGLSpr.SetTextBinding .sprSht1,"CREDAY",.sprSht1.ActiveRow, .txtCREDAY.value  
	End with
End Sub

Sub sprSht1_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""
		IF Col = 10 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(6,0)				
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+5,Row
			End IF
		elseIF Col = 13 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
			End IF
		elseIF Col = 7 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN0") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(8,0)		
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+8,Row
			End IF
		
		end if
		.txtYEARMON.focus
		.sprSht1.focus 

	End with
	
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht1, Col, Row)
dim vntRet, vntInParams
	with frmThis
		IF Col = 10 Then			
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN1") then exit Sub
			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(6,0)		
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+4,Row
			End IF
		elseIF Col = 13 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
			End IF
		elseIF Col = 7 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(8,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+7,Row
			End IF
		
		end if
		.txtYEARMON.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.sprSht1.Focus
	end with
End Sub


Sub SUM_AMT()
	Dim lngCnt
	Dim strSUMDEMANDAMT
	Dim strDIVAMT
	strSUMDEMANDAMT = 0
	With frmThis
		For lngCnt = 1 To .sprSht1.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		
		mobjSCGLSpr.SetTextBinding .sprShtSum,"DIVAMT",1, strSUMDEMANDAMT
	End With
End Sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD >
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">
												&nbsp;청구관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="LEFT" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" width="100%" height="100%"  cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="LEFT" colSpan="2">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0" align="LEFT">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
													width="90">년월
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="년월" style="WIDTH: 102px; HEIGHT: 22px" type="text"
														maxLength="6" size="11" name="txtYEARMON" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME,txtJOBNO)"
													width="90">JOB명
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 378px"><INPUT class="INPUT_L" id="txtJOBNAME" title="코드명" style="WIDTH: 256px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="37" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23"
														align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="코드조회" style="WIDTH: 65px; HEIGHT: 22px" type="text"
														maxLength="8" align="left" size="3" name="txtJOBNO"></TD>
												<TD class="SEARCHLABEL" width="90">완료구분
												</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbYN" title="사용구분" style="WIDTH: 104px" name="cmbYN">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="Y">완료</OPTION>
														<OPTION value="N">미완료</OPTION>
													</SELECT>
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 1040px; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;JOB 리스트</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
									<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
									<!--내용 및 그리드-->
								</TR>
								<TR>
									<!--내용-->
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 40%" vAlign="top" align="left">
										<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27517">
												<PARAM NAME="_ExtentY" VALUE="9604">
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
									<TD>
										<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px" id="lblstatus"><FONT face="굴림"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left"  height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="굴림"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">
																&nbsp;청구내역 분할</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD>
																<!--Hidden Control Start-->
																<INPUT class="NOINPUT" id="txtPREESTNO" style="WIDTH: 8px; HEIGHT: 22px" tabIndex="1" type="text"
																	size="1" name="txtPREESTNO"> <INPUT class="NOINPUT" id="txtYEARMONPOP" style="WIDTH: 16px; HEIGHT: 22px" tabIndex="1"
																	type="text" size="1" name="txtYEARMONPOP"> <INPUT class="NOINPUT" id="txtCREDAY" style="WIDTH: 13px; HEIGHT: 22px" tabIndex="1" type="text"
																	size="1" name="txtCREDAY"> <INPUT class="NOINPUT" id="txtJOBNOPOP" style="WIDTH: 32px; HEIGHT: 22px" readOnly type="text"
																	size="1" name="txtJOBNOPOP"> <INPUT class="NOINPUT" id="txtDIVAMT" style="WIDTH: 40px; HEIGHT: 22px" tabIndex="1" readOnly
																	type="text" size="1" name="txtDIVAMT"> 
																<!--Hidden Control End-->
															</TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
															<td><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></td>
															<TD><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'"
																	alt="한 행 삭제" src="../../../images/imgDelRow.gif" width="54" border="0" name="imgDelRow"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
										<TABLE id="tblBody1" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
											border="0">
											<TR>
												<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
										<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 90%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27517">
											<PARAM NAME="_ExtentY" VALUE="6376">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
											
										<OBJECT id="sprShtSum" style="WIDTH: 100%; HEIGHT: 8%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27517">
											<PARAM NAME="_ExtentY" VALUE="609">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
									</div>
										</FONT>
									</TD>
								</TR>
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR>
			<TR>
				<TD class="BOTTOMSPLIT" id="lblStatus2" style="WIDTH: 1040px"></TD>
			</TR>
			<!--Bottom Split End--> </TBODY></TABLE> 
			<!--Input Define Table End-->
			</TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
