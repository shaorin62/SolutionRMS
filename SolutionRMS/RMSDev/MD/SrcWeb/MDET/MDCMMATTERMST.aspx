<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMATTERMST.aspx.vb" Inherits="MD.MDCMMATTERMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>소재코드 관리</title>
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
Dim mobjMDCMMATTERMST
Dim mobjMDCMMEDGet

Dim mobjMDCMGET
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
Sub imgNew_onclick()
Call MATTERCODE_NEWPOP()
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
	Set mobjMDCMMATTERMST = gCreateRemoteObject("cMDCO.ccMDCOCODETR")
	Set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")

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
		'CC_CODE,CC_NAME,OC_CODE,OC_NAME,USE_YN,STDATE,EDATE
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1

		mobjSCGLSpr.SpreadDataField .sprSht,    "MATTERCODE|MATTER|CLIENTCODE|CLIENTNAME|SUBSEQ|SUBSEQNAME|EXCLIENTCODE|BTN|EXCLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|DEPTCD|DEPTNAME|ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,		    "코드|소재명|광고주코드|광고주명|브랜드코드|브랜드명|제작대행사코드|제작대행사명|사업부코드|사업부명|부서코드|부서명|입력"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "5   |22    |10        |16      |9         |19      |10          |2|          16|10        |19      |10      |13    |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht,"EXCLIENTCODE|EXCLIENTNAME" , -1,-1,200
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MATTERCODE|CLIENTCODE|SUBSEQ|EXCLIENTCODE",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MATTER|CLIENTNAME|SUBSEQNAME|EXCLIENTNAME",-1,-1,0,2,false
		
		
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"MATTERCODE|MATTER|CLIENTCODE|CLIENTNAME|SUBSEQ|SUBSEQNAME|CLIENTSUBCODE|CLIENTSUBNAME|DEPTCD|DEPTNAME"
		'mobjSCGLSpr.SetCellTypeComboBox .sprSht,6,6,,,strComboList
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|SUBSEQ|CLIENTSUBCODE|DEPTCD|ATTR01", true
	End with

	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
    Dim intMATTERCODE
	Dim strMATTER
	Dim strSEQCODE
	Dim strSEQNAME
	Dim strCUSTCODE
	Dim strCUSTNAME
	Dim strEXCLIENTCODE
	Dim strEXCLIENTNAME
	Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intMATTERCODE = .txtMATTERCODE.value
		strMATTER = .txtMATTERNAME.value
		strSEQCODE = .txtSUBSEQ.value
		strSEQNAME = .txtSUBSEQNAME.value
		strCUSTCODE = .txtCLIENTCODE.value 
		strCUSTNAME = .txtCLIENTNAME.value 
		strEXCLIENTCODE = .txtEXCLIENTCODE1.value
		strEXCLIENTNAME = .txtEXCLIENTNAME1.value 
		vntData = mobjMDCMMATTERMST.SelectRtn_MATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,intMATTERCODE,strMATTER,strSEQCODE,strSEQNAME,strCUSTCODE,strCUSTNAME,strEXCLIENTCODE,strEXCLIENTNAME)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
				for intCnt = 1 To .sprSht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",intCnt) = "1" Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						Else
							If intCnt Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False	
							Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
							End If
						End If
				Next
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
		IF Col = 8 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtMATTERNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
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
	Dim strUSEYN
	Dim vntData
	Dim strCC
	strUSEYN = ""
	strCC = ""
	with frmThis
	iF Col = 6 Then
		If mobjSCGLSpr.GetTextBinding(.sprSht,"USE_YN",.sprSht.ActiveRow) = "사용" Or mobjSCGLSpr.GetTextBinding(.sprSht,"USE_YN",.sprSht.ActiveRow) = "미사용"then
		'사용으로 바꿀수 없는경우 체크 (부서코드는 사용으로 들어가면 키값이 부서코드와 사용 이므로 중복일 경우가 있기 때문에)
			if mobjSCGLSpr.GetTextBinding(.sprSht,"USE_YN",.sprSht.ActiveRow) = "사용" Then
			strUSEYN = "Y"
			Else
			strUSEYN = "N"
			End IF 
			'msgbox strUSEYN
			strCC = mobjSCGLSpr.GetTextBinding(.sprSht,"CC_CODE",.sprSht.ActiveRow)
			mlngRowCnt1=clng(0)
			mlngColCnt1=clng(0)
			vntData = mobjMDCMMATTERMST.GetDup(gstrConfigXml,mlngRowCnt1,mlngColCnt1,strUSEYN,strCC)
			if not gDoErrorRtn ("GetDup") then
				If mlngRowCnt1 > 0 Then
   				gErrorMsgBox "이자료는사용구분을 바꿀수 없습니다.","선택안내!"
   					If strUSEYN = "Y" then
   						mobjSCGLSpr.SetTextBinding .sprSht,"USE_YN",.sprSht.ActiveRow, "미사용"
   					Else
   						mobjSCGLSpr.SetTextBinding .sprSht,"USE_YN",.sprSht.ActiveRow, "사용"
   					End If
   				End If
   			end if
		End If
	End If
	End With
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
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
	with frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"MATTERCODE|EXCLIENTCODE")
		'처리 업무객체 호출
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		intRtn = mobjMDCMMATTERMST.ProcessRtn_MATTER(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
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
	set mobjMDCMMATTERMST = Nothing
	set mobjMDCMMEDGet = Nothing
	Set mobjMDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.chkCHOICE.style.visibility = "hidden"
	.sprSht.maxrows = 0
	End with
End Sub

sub DeleteRtn
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
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			gSetChangeFlag .txtCLIENTCODE                 ' gSetChangeFlag objectID	 Flag 변경 알림
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
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
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
' 브랜드코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgSUBSEQCODE_onclick
	'with frmThis
	'	If .txtCLIENTCODE.value = "" Then
	'		gErrorMsgBox "브랜드검색은 광고주코드를 먼저 조회하십시오.","검색안내!"
	'		Exit Sub
	'	End If 
	'End with
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		if isArray(vntRet) then
			if .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtSUBSEQ.value = trim(vntRet(1,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(2,0))	' 브랜드명 표시
			.txtCLIENTCODE.value = trim(vntRet(3,0))		' 광고주 표시
			.txtCLIENTNAME.value = trim(vntRet(4,0))	' 광고주명 표시
			'.txtPUB_DATE.focus()					' 포커스 이동
			gSetChangeFlag .txtSUBSEQ		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
		
			if window.event.keyCode = meEnter then
				
			
					Dim vntData
   					Dim i, strCols
					'On error resume next
					with frmThis
						'Long Type의 ByRef 변수의 초기화
						mlngRowCnt=clng(0)
						mlngColCnt=clng(0)
						vntData = mobjMDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
						if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
							If mlngRowCnt = 1 Then
								.txtSUBSEQ.value = trim(vntData(1,0))
								.txtSUBSEQNAME.value = trim(vntData(2,0))
								.txtCLIENTCODE.value = trim(vntData(3,0))		' 광고주 표시
								.txtCLIENTNAME.value = trim(vntData(4,0))	' 광고주
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
' 대행사 코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
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
			'.txtMEDNAME.focus()					' 포커스 이동
			'gSetChangeFlag .txtEXCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
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
			vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(0,0))
					.txtEXCLIENTNAME.value = trim(vntData(1,0))
					'.txtMEDNAME.focus()
					'GetBrandAndDept'광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
				Else
					Call EXCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 대행사 코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEXCLIENTCODE1_onclick
	Call EXCLIENTCODE_POP1()
End Sub

'실제 데이터List 가져오기
Sub EXCLIENTCODE_POP1
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE1.value), trim(.txtEXCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtEXCLIENTCODE1.value = vntRet(0,0) and .txtEXCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEXCLIENTCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtEXCLIENTNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
			'.txtMEDNAME.focus()					' 포커스 이동
			'gSetChangeFlag .txtEXCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEXCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE1.value),trim(.txtEXCLIENTNAME1.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE1.value = trim(vntData(0,0))
					.txtEXCLIENTNAME1.value = trim(vntData(1,0))
					'.txtMEDNAME.focus()
					'GetBrandAndDept'광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
				Else
					Call EXCLIENTCODE_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
Sub ImgEXCLIENTApp_onclick
	Dim intCnt
	With frmThis
		If .chkCHOICE.checked = true Then
			For intCnt = 1 To .sprSht.MaxRows
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",intCnt,.txtEXCLIENTCODE.value 
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",intCnt,.txtEXCLIENTNAME.value 
					mobjSCGLSpr.CellChanged .sprSht, 7,intCnt
			Next
		Else 
			For intCnt = 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",intCnt,.txtEXCLIENTCODE.value 
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",intCnt,.txtEXCLIENTNAME.value 
					mobjSCGLSpr.CellChanged .sprSht, 7,intCnt
				End If
			Next
		End If
	End With
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
		vntInParams = array(trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("MDCMMATTERPOP.aspx",vntInParams , 783,473)
		if isArray(vntRet) then
			if .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMATTERCODE.value = trim(vntRet(0,0))		' 브랜드 표시
			.txtMATTERNAME.value = trim(vntRet(1,0))	' 브랜드명 표시 2,3,6,7
			.txtCLIENTCODE.value = trim(vntRet(2,0))
			.txtCLIENTNAME.value = trim(vntRet(3,0))
			.txtSUBSEQ.value = trim(vntRet(6,0))
			.txtSUBSEQNAME.value = trim(vntRet(7,0))
			.txtEXCLIENTCODE1.value = trim(vntRet(8,0))
			.txtEXCLIENTNAME1.value = trim(vntRet(9,0))
			'gSetChangeFlag .txtSEARCHSEQCODE
     	end if
	End with
	gSetChange
End Sub

Sub txtMATTERNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMMATTERMST.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value),"","")
			if not gDoErrorRtn ("GetMATTER") then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntData(0,1))
					.txtMATTERNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))
					.txtCLIENTNAME.value = trim(vntData(3,1))
					.txtSUBSEQ.value = trim(vntData(6,1))
					.txtSUBSEQNAME.value = trim(vntData(7,1))
					.txtEXCLIENTCODE1.value = trim(vntData(8,1))
					.txtEXCLIENTNAME1.value = trim(vntData(9,1))
				Else
					Call MATTERCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 소재등록 팝업 버튼
'-----------------------------------------------------------------------------------------
Sub MATTERCODE_NEWPOP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("MDCMMATTERNEWPOP.aspx",vntInParams , 783,483)
		'msgbox vntRet
		'if isArray(vntRet) then
		'	if .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		'	.txtMATTERCODE.value = trim(vntRet(0,0))		' 브랜드 표시
		'	.txtMATTERNAME.value = trim(vntRet(1,0))	' 브랜드명 표시 2,3,6,7
		'	.txtCLIENTCODE.value = trim(vntRet(2,0))
		'	.txtCLIENTNAME.value = trim(vntRet(3,0))
		'	.txtSUBSEQ.value = trim(vntRet(6,0))
		'	.txtSUBSEQNAME.value = trim(vntRet(7,0))
		'	.txtEXCLIENTCODE1.value = trim(vntRet(8,0))
		'	.txtEXCLIENTNAME1.value = trim(vntRet(9,0))
		'	'gSetChangeFlag .txtSEARCHSEQCODE
     	'end if
	End with
	SelectRtn
	gSetChange
	'msgbox "1111"
End Sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
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
											<td class="TITLE">&nbsp;매체 정보관리</td>
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
						<TABLE id="tblBody" style="WIDTH: 1040px; HEIGHT: 32px" cellSpacing="0" cellPadding="0"
							width="1040" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center"
									colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME,txtMATTERCODE)"
												width="85">소재명&nbsp;
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMATTERNAME" title="소재명" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" size="34" name="txtMATTERNAME"><IMG id="ImgMATTER" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgMATTER"><INPUT class="INPUT_L" id="txtMATTERCODE" title="소재코드" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="6" name="txtMATTERCODE"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="86">광고주</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" size="32" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTCODE1,txtEXCLIENTNAME1)"
												width="85">제작대행사&nbsp;
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtEXCLIENTNAME1" title="대대행사명" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" size="34" name="txtEXCLIENTNAME1"><IMG id="ImgEXCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgEXCLIENTCODE1"><INPUT class="INPUT_L" id="txtEXCLIENTCODE1" title="대대행사코드" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="6" name="txtEXCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME,txtSUBSEQ)"
												width="86">브랜드</TD>
											<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="브랜드명" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" size="32" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgSUBSEQCODE"><INPUT class="INPUT_L" id="txtSUBSEQ" title="브랜드코드" style="WIDTH: 53px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="6" name="txtSUBSEQ"></TD>
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
											<TD align="left" width="400" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;소재 정보관리</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
																src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
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
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 86px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTCODE,txtEXCLIENTNAME)">제작대행사적용
											</TD>
											<TD class="DATA"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="브랜드명적용값" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" size="34" name="txtEXCLIENTNAME"><IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgEXCLIENTCODE"><INPUT class="INPUT_L" id="txtEXCLIENTCODE" title="대대행사코드적용값" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="6" name="txtEXCLIENTCODE">&nbsp;<IMG id="ImgEXCLIENTApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="사업부를 일괄적용합니다" src="../../../images/ImgApp.gif" width="54" align="absMiddle"
													border="0" name="ImgEXCLIENTApp"><INPUT id="chkCHOICE" type="checkbox" name="chkCHOICE">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px"></TD>
							</TR>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD class="DATAFRAME">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 651px"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 651px"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27490">
											<PARAM NAME="_ExtentY" VALUE="17224">
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
							<!--BodySplit End-->
							<!--List Start--></TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
				</TR>
				<!--BodySplit End-->
				<!--Brench Start-->
				<TR>
					<TD class="BRANCHFRAME" style="WIDTH: 1040px">
						<!--<INPUT class="BUTTON" id="btn1" style="WIDTH: 123px; HEIGHT: 16pt" type="button" value="분기버튼"
											name="Button">--></TD>
				</TR>
				<!--Brench End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 1040px"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
