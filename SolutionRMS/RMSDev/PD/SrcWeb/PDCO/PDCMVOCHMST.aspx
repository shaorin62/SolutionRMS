<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMVOCHMST.aspx.vb" Inherits="PD.PDCMVOCHMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>제작비 전표생성</title>
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
Dim mobjPDCMVOCHMST
Dim mobjMDCMMEDGet
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mobjPDCMGET
Dim mstrCheck
mstrCheck=True

CONST meTAB = 9
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	'EndPage
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
Sub imgFind_onclick()
	FILE_POP
End Sub
Sub imgManageApp_onclick()
gFlowWait meWAIT_ON
	with frmThis
	Dim strNO
	strNO = .txtFILENO.value
	Call Voch_Batch(strNO)
	window.setTimeout "SelectRtn", 3000	
	End With
gFlowWait meWAIT_OFF
End Sub
Sub ImgSUMMApp_onclick()
	SummApp
End Sub
Sub imgDelete_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn
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
	
	Set mobjPDCMVOCHMST = gCreateRemoteObject("cPDCO.ccPDCOVOCHMST")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'set mobjMDCMMEDGet = gCreateRemoteObject("cMDCM.ccMDCMCUSTGET")
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
	
	
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		'CC_CODE,CC_NAME,OC_CODE,OC_NAME,USE_YN,STDATE,EDATE
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 20, 0, 4
		'mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK|POSTINGDATE|CUSTOMERCODE|SUMM|BA|COSTCENTER|SUMAMT|VAT|SEMU|BP|DEMANDDAY|TAXYEARMON|TAXNO|GBN|VOCHNO|RMSNO|MEDFLAG|ERRCODE|ERRMSG|BTN"
		mobjSCGLSpr.SetHeader .sprSht,		    "선택|전표일자|거래처코드|적요|BA|COSTCENTER|금액|부가세|세무코드|BP|지급기일|RMS년월|RMS번호|구분|전표번호|파일번호|매체구분|에러코드|에러메세지|삭제"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "4   |8       |10        |20  |6 |10        |11  |11    |8       |6 |8       |8      |6      |6   |10      |10      |0       |0       |50        |8"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"전표삭제", "BTN"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "POSTINGDATE|DEMANDDAY"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "BA|SEMU|BP|TAXYEARMON|TAXNO|GBN|VOCHNO|RMSNO|CUSTOMERCODE|COSTCENTER",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SUMM|ERRMSG",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUMAMT|VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"POSTINGDATE|CUSTOMERCODE|SUMM|BA|SUMAMT|VAT|SEMU|BP|DEMANDDAY|COSTCENTER|TAXYEARMON|TAXNO|GBN|VOCHNO|RMSNO|ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht, "MEDFLAG|ERRCODE", true
	End with

	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
    Dim intCnt,intRowCnt
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
    
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		'strREAL_MED_CODE = .txtREAL_MED_CODE.value
		'strREAL_MED_NAME = .txtREAL_MED_NAME.value
		strGBN = .cmbGBN.value 
		'년월,광고주,광고주명,구분,파일명
		
		vntData = mobjPDCMVOCHMST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strCLIENTCODE,strCLIENTNAME,strGBN,.txtFILENO.value)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows
					If  mobjSCGLSpr.GetTextBinding( .sprSht,"RMSNO",intCnt) <> "" Then
						'스태틱
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					else
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
					End If
					If mobjSCGLSpr.GetTextBinding( .sprSht,"RMSNO",intCnt) <> "" And mobjSCGLSpr.GetTextBinding( .sprSht,"ERRCODE",intCnt) = "1" Then
						'체크
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false			
					End If
				Next
   			Else
   				'initpageData
   				.sprSht.MaxRows = 0
   				
   				'PreSearchFiledValue strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strGBN
   			end If
   			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

'****************************************************************************************
'이전 검색어를 담아 놓는다.
'****************************************************************************************
Sub PreSearchFiledValue (strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strGBN)
	frmThis.txtYEARMON.value = strYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME.value = strCLIENTNAME
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
	frmThis.cmbGBN.value  = strGBN
End Sub


Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	Dim intRtnDell
	Dim strYEAR
	Dim strSP
	Dim vntData
	Dim intRntYN,intRntYN_1,intRntYN_2,intRntYN_3,intRntYN_4,intRntYN_5
	with frmThis
	intRtnDell = 0
	
	
		IF Col = 20 Then
			If  mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",.sprSht.activeRow) = "" Then
				msgbox "전표내역이 없습니다."
				Exit Sub
			End If
			strYEAR = MID(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXYEARMON",Row)),1,4)
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
				intRntYN = gYesNoMsgbox("전표를 삭제하시겠습니까?","자료삭제 확인")
				IF intRntYN <> vbYes then exit Sub
					
			vntInParams = array(strYEAR, TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)))
			
			vntRet = gShowModalWindow("../../../MD/SrcWeb/MD/MDCMVOCHDEL.aspx",vntInParams , 413,126)
			If vntRet = "" Then exit Sub
			strSP = split(vntRet,"|")
			IF isArray(strSP) then
			'if strSP(3) = "" Then
			'Exit Sub
			'End If
				Select Case  strSP(3)
					Case "8"
					intRntYN_1 = gYesNoMsgbox("ERP에 존재하지 않는 전표입니다. RMS전표를 삭제 하겠습니까?","전표삭제 확인")
						IF intRntYN_1 <> vbYes then exit Sub

						intRtnDell = mobjPDCMVOCHMST.VOCHDELL(gstrConfigXml,strYEAR,TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXYEARMON",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",Row)))
						if not gDoErrorRtn ("VOCHDELL") then
							if intRtnDell > 0 Then
							gErrorMsgBox "전표번호" & TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)) & " 번 이 삭제되었습니다.","처리안내"
							End If
							SelectRtn
   						end if
					Case "0"
						intRtnDell = mobjPDCMVOCHMST.VOCHDELL(gstrConfigXml,strYEAR,TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXYEARMON",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",Row)))
						if not gDoErrorRtn ("VOCHDELL") then
							if intRtnDell > 0 Then
							gErrorMsgBox "전표번호" & TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)) & " 번 이 삭제되었습니다.","처리안내"
							End If
							SelectRtn
   						end if
					Case "4"
						gErrorMsgBox "ERP 에서 승인 완료된 전표입니다." & vbcrlf & "" & vbcrlf & "전표를 삭제 하시려면 ERP 에서 승인취소를 하십시오." ,"처리안내"
   					Case "9"
					intRntYN_2 = gYesNoMsgbox("ERP에서 승인이 취소된 전표입니다." & vbcrlf & "" & vbcrlf & "전표를 삭제 하시겠습니까?","전표삭제 확인")
					IF intRntYN_2 <> vbYes then exit Sub
					intRtnDell = mobjPDCMVOCHMST.VOCHDELL(gstrConfigXml,strYEAR,TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXYEARMON",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",Row)))
					if not gDoErrorRtn ("VOCHDELL") then
						if intRtnDell > 0 Then
						gErrorMsgBox "전표번호" & TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"VOCHNO",Row)) & " 번 이 삭제되었습니다.","처리안내"
						End If
						SelectRtn
   					end if
					Case "2"
						gErrorMsgBox "ERP전표삭제시 오류가 발생하였습니다." & vbcrlf & "전산담당자 에게 문의 하십시오."	,"처리안내"					
				End Select
			End If 
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			'vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_NAME",Row)))
			'vntRet = gShowModalWindow("MDCMDEPTPOP.aspx",vntInParams , 413,425)
			'IF isArray(vntRet) then
		'		mobjSCGLSpr.SetTextBinding .sprSht,"OC_CODE",Row, vntRet(0,0)
	'			mobjSCGLSpr.SetTextBinding .sprSht,"OC_NAME",Row, vntRet(1,0)			
	'			mobjSCGLSpr.CellChanged .sprSht, Col,Row
	'		End IF
	'		.txtDEPTCODE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
	'		.sprSht.Focus
	'		mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub
Sub SummApp
	Dim intCnt,intCnt2
	Dim intSumCnt
	Dim intRtn
	with frmThis
	intSumCnt = 0
		For intCnt = 1 To .sprSht.MaxRows
			If Trim(mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt)) = "" Then
			Else
				If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) = 1 Then 
				intSumCnt = intSumCnt +1
				End If
			end If
		Next
		If intSumCnt = 0  Then
		
		
			Exit Sub
		Elseif Trim(.txtSUMM.value) <> "" Then 
			intRtn = gYesNoMsgbox("적요를 변경하시겠습니까?","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			For intCnt2 = 1 To .sprSht.MaxRows
			If Trim(mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt2)) = "" Then
			Else
				If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt2) = 1 Then 
					mobjSCGLSpr.SetTextBinding .sprSht,"SUMM",intCnt2, .txtSUMM.value 
				End If
			End If
			Next
		End If
		
		
	End With
End Sub
Sub FILE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtYEARMON.value), trim(.txtFILENO.value))
		
			vntRet = gShowModalWindow("PDCMTRUVOCHFILEPOP.aspx",vntInParams , 413,425)
		
			
		if isArray(vntRet) then
			if .txtFILENO.value = vntRet(0,0) and .txtYEARMON.value = vntRet(4,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtFILENO.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtYEARMON.value = trim(vntRet(4,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	'selectRtn
End Sub
'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	Dim strUSEYN
	Dim vntData
	Dim strCC
	strUSEYN = ""
	strCC = ""
	with frmThis
	
	End With
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
End Sub

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

	dim VLength
	dim temp
	dim EscTemp
	dIM i
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





Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
	if Row > 0 and Col > 1 then		
			'sprShtToFieldBinding Col,Row
	elseif Col = 1 and Row = 0 then
	'msgbox mstrCheck
		mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
		if mstrCheck = True then 
			for intCnt = 1 To .sprSht.MaxRows
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
				'End If
			Next
				
			mstrCheck = False
		elseif mstrCheck = False then 
			mstrCheck = True
		end if
	
		For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"RMSNO",intCnt) <> "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If		
				If mobjSCGLSpr.GetTextBinding( .sprSht,"RMSNO",intCnt) <> "" And mobjSCGLSpr.GetTextBinding( .sprSht,"ERRCODE",intCnt) = "1" Then
							mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,mstrCheck								
				End If	
		Next
	end if 
	End With
End Sub 

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
		Dim strNO	
if Col = 16 Then
			strNO=mobjSCGLSpr.GetTextBinding( .sprSht,"RMSNO",Row) 
			Call ExcelDownLoad(strNO)
		End If
	end with
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

Sub ProcessRtn()
	Dim intRtn
   	dim vntData
   	Dim vntSelect
   	Dim strYEARMON
   	Dim strSAVEYEARMON
   	Dim strSAVESEQ
   	Dim strSAVERMSNO
	with frmThis
	
		
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		If .txtYEARMON.value = "" Then
			msgbox "조회년월을 기입하셔야 저장이 가능합니다."
			exit Sub
		Else 
		    strYEARMON = .txtYEARMON.value
		End If
		vntSelect = mobjPDCMVOCHMST.SelectRtn_SEQNO(strYEARMON)
		if  IsArray(vntSelect) then 
		
			strSAVEYEARMON = vntSelect(0,1)
			strSAVESEQ =vntSelect(1,1) 
			strSAVERMSNO =vntSelect(2,1)
			'msgbox strSAVEYEARMON & strSAVESEQ & strSAVERMSNO
		End If
		'Exit Sub
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|POSTINGDATE|CUSTOMERCODE|SUMM|BA|COSTCENTER|SUMAMT|VAT|SEMU|BP|DEMANDDAY|TAXYEARMON|TAXNO|GBN|VOCHNO|RMSNO|MEDFLAG")
		'처리 업무객체 호출
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		'exit Sub
		
		intRtn = mobjPDCMVOCHMST.ProcessRtn(gstrConfigXml,vntData,strYEARMON,strSAVEYEARMON,strSAVESEQ,strSAVERMSNO)
		if not gDoErrorRtn ("ProcessRtn") then
	
			Call Excel_save (strSAVERMSNO)
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 1 Then
			gErrorMsgBox intRtn & " 건 이 저장되었습니다.","저장안내"
			End If
			SelectRtn
			
		
   		end if
   	end with
End Sub
Sub EndPage()
	set mobjPDCMVOCHMST = Nothing
	'set mobjMDCMMEDGet = Nothing
	Set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtYEARMON.focus		
	End with
End Sub

sub DeleteRtn
	Dim intRtn
   	dim vntData
   	Dim intCnt
	with frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TAXYEARMON|TAXNO|ERRCODE")
		'처리 업무객체 호출
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"삭제취소"
			exit sub
		End If
		
		intRtn = mobjPDCMVOCHMST.DeleteRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("DeleteRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 0 Then
			gErrorMsgBox"에러내역 이 삭제되었습니다.","저장안내"
			
			End If
			initpageData
			SelectRtn
   		end if
   		
   		
   	end with
   	
End Sub
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_POP()
End Sub
'-----------------------------------------------------------------------------------------
' 광고주팝업(조회)
'-----------------------------------------------------------------------------------------
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)	
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
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
				vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
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
' 매체사팝업(조회) 'txtREAL_MED_NAME / txtREAL_MED_CODE
'-----------------------------------------------------------------------------------------
Sub REAL_MED_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
		
			vntRet = gShowModalWindow("MDCMREALMEDPOP.aspx",vntInParams, 413,425)
		
			
		if isArray(vntRet) then
			if .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
				vntData = mobjPDCMGET.GetREALMEDNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value))
			if not gDoErrorRtn ("txtREAL_MED_NAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,0))
					.txtREAL_MED_NAME.value = trim(vntData(1,0))
				Else
					Call REAL_MED_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
		</script>
		<script language="javascript">
		function Excel_save(strSAVERMSNO){
		location.href = "PDCMVOCHMSTSUB.asp?temp_filename="+ strSAVERMSNO; 
		}
		
		function Voch_Batch(strNO){
		ifrm.location.href = "PDCMVOCHMSTSUB2.asp?temp_filename="+ strNO;		
		}
		function ExcelDownLoad(strNO){
		ifrm.location.href = "PDCMEXCELDOWNLOAD.asp?temp_filename="+ strNO;		
		}
		
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%"  HEIGHT="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 100%">
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
											<td class="TITLE">&nbsp;청구관리</td>
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
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 95%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT"  colSpan="2"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD  style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="left"
										colSpan="2" class="SEARCHDATA">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="1042" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 75px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">&nbsp;년월
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 98px"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="10" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 76px; CURSOR: hand" width="76">상태구분
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 266px"><SELECT id="cmbGBN" title="완료구분선택" style="WIDTH: 144px; HEIGHT: 22px" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
														name="cmbGBN">
														<OPTION value="A" selected>전체</OPTION>
														<OPTION value="M">처리중</OPTION>
														<OPTION value="Y">완료</OPTION>
														<OPTION value="N">미완료</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 74px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,txtCLIENTCODE)"
													width="74">&nbsp;광고주
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 288px; HEIGHT: 22px"
														type="text" maxLength="100" size="42" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
														accessKey=",M" type="text" maxLength="6" size="3" name="txtCLIENTCODE">
												</TD>
												<td class="SEARCHDATA" ><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
											</TR>
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
												<TD align="right" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="굴림"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;전표생성</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" width="50" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
														<TR>
															<TD></TD>
															<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																	height="20" alt="자료를 저장합니다." src="../../../images/ImgvochCre.gIF" border="0" name="imgSave"></td>
															<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																	height="20" alt="오류전표 를 삭제합니다." src="../../../images/ImgErrVochDel.gIF" border="0"
																	name="imgDelete"></td>
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
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<TR>
									<TD>
										<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 73px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFILENO,'')">파일관리
												</TD>
												<TD class="DATA" style="WIDTH: 450px"><IMG id="imgFind" onmouseover="JavaScript:this.src='../../../images/imgFindOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgFind.gIF'" height="20" alt="자료를 검색합니다." src="../../../images/imgFind.gIF"
														align="absMiddle" border="0" name="imgFind"><INPUT id="txtFILENO" style="WIDTH: 142px; HEIGHT: 21px" type="text" size="18" name="txtFILENO"><IMG id="imgManageApp" onmouseover="JavaScript:this.src='../../../images/imgManageOn.gIF'"
														title="전표번호를 업데이트 합니다." style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgManage.gIF'" height="20" alt="전표번호를 업데이트 합니다." src="../../../images/imgManage.gif" width="54" align="absMiddle" border="0" name="imgManageApp">
												</TD>
												<TD class="LABEL" style="WIDTH: 75px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUMM,'')">적요적용
												</TD>
												<TD class="DATA"><INPUT id="txtSUMM" title="적요적용" style="WIDTH: 368px; HEIGHT: 21px" type="text" size="56"
														name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
														title="적요를 일괄 적용합니다" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
														height="20" alt="적요를 일괄 적용합니다" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0"
														name="ImgSUMMApp">
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px"></TD>
								</TR>
								<!--내용 및 그리드-->
								<TR vAlign="top" align="left">
									<!--내용-->
									<TD  class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 95%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 95%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; "
												height="95%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="18150">
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
							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
			<P>
				<!--Input Define Table End--> </TD></TR> 
				<!--Top TR End--> </TBODY></TABLE> 
				<!--Main End--></P>
		</FORM>
		</TR></TBODY></TABLE><iframe id="ifrm" frameBorder="0" width="0" height="0"></iframe>
	</body>
</HTML>
