<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPREESTLIST.aspx.vb" Inherits="PD.PDCMPREESTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>가견적 관리</title>
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
Dim mobjPDCMPREESTLIST, mobjPDCMGET,mobjPDCMJOBNO
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
Dim mstrCHKROW
mstrCHKROW = false
Const meTab = 9
mALLCHECK = TRUE
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
Sub imgFind_onclick()
Dim vntRet
	vntRet = gShowModalWindow("PDCMCHARGELISTPOP.aspx","" , 1060,730)
End Sub
'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
mstrCHKROW = false
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
	
End Sub
Sub imgQuery1_onclick
	gFlowWait meWAIT_ON
	SelectRtn_HDR2
	gFlowWait meWAIT_OFF
	
End Sub

Sub imgNew_onclick
	Dim vntInParams
	Dim vntRet
	Dim strRow
	with frmThis
	vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow))
	vntRet = gShowModalWindow("PDCMPREESTDTLNEW.aspx",vntInParams , 1060,780)
	strRow = .sprSht.ActiveRow
	selectRtn
	mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
	Call sprSht_click(1,strRow)
	End with
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
Sub imgExcel1_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht1
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgRowAdd_onclick ()
	
	With frmThis
		
		call sprSht_Keydown(meINS_ROW, 0)
		intiSprValue
	End With 
End Sub

Sub imgRowDel_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDetail_onclick()
Dim vntInParams
Dim vntRet
Dim strJOBNO
Dim strRow
	with frmThis
		if mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow) = "" then
				gErrorMsgBox "가견적 List 를 우선저장하시고 입력 하십시오.","처리안내" 
				Exit Sub
			End if
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow))
			vntRet = gShowModalWindow("PDCMPREESTDTL.aspx",vntInParams , 1060,780)
			
			
			.txtCLIENTSUBNAME.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strRow = .sprSht.ActiveRow
			SelectRtn
			SelectRtn_DBLHDR(strJOBNO)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow	
	End with
End Sub

Sub Imgcopy_onclick()
Dim intRtn
Dim intRtnCopy
Dim strOLDCODE
Dim strJOBNO
Dim strCREDAY
Dim strCLIENTSUBCODE
Dim strCOMMITION
Dim strCLIENTCODE
Dim strSUBSEQ
Dim vntRet
Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			strJOBNO         = trim(vntRet(0,0))  ' Code값 저장
			strCREDAY        = trim(vntRet(6,0))  ' 코드명 표시
			strCLIENTSUBCODE = trim(vntRet(2,0)) 
			strCOMMITION     = trim(vntRet(3,0)) 
			strCLIENTCODE    = trim(vntRet(4,0)) 
			strSUBSEQ        = trim(vntRet(5,0)) 
		Else
			Exit Sub
     	end if

		strOLDCODE = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",.sprSht1.ActiveRow)
		'strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
		'strCREDAY = Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",.sprSht.ActiveRow),"-","")
		'strCLIENTSUBCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow)
		'strCOMMITION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMITION",.sprSht.ActiveRow)
		'strCOMMITION = CDBL(strCOMMITION)
		'strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		'strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
		
		'strCREDAY,strCLIENTSUBCODE,strCOMMITION,strCLIENTCODE
		intRtn = gYesNoMsgbox("[" & strOLDCODE & "] 자료를 복사하시겠습니까?","내역복사 확인")
		IF intRtn <> vbYes then exit Sub
		intRtnCopy = mobjPDCMPREESTLIST.ProcessRtn_COPY(gstrConfigXml,strOLDCODE,strJOBNO,strCREDAY,strCLIENTSUBCODE,strCOMMITION,strCLIENTCODE,strSUBSEQ)
		if not gDoErrorRtn ("ProcessRtn_COPY") then
			gErrorMsgBox " 자료가 복사" & mePROC_DONE,"저장안내" 
			.txtJOBNO.value = trim(vntRet(0,0)) 
			.txtJOBNAME.value = trim(vntRet(1,0)) 
			SelectRtn
  		end if
	End with 
End Sub
Sub intiSprValue
	Dim strJOBNAME
	Dim strJOBNO
	Dim strCREDAY
	Dim strCLIENTSUBCODE
	Dim strCOMMITION
	Dim strCLIENTCODE
	Dim strSUBSEQ
	with frmThis
		If .sprSht.MaxRows <> 0 Then
			strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow)
			strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
			strCREDAY = Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",.sprSht.ActiveRow),"-","")
			strCLIENTSUBCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow)
			strCOMMITION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMITION",.sprSht.ActiveRow)
			strCOMMITION = CDBL(strCOMMITION)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
			strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"JOBNO",.sprSht1.ActiveRow, strJOBNO
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"JOBNAME",.sprSht1.ActiveRow, strJOBNAME
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CREDAY",.sprSht1.ActiveRow, strCREDAY
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"SUSURATE",.sprSht1.ActiveRow, strCOMMITION
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CLIENTSUBCODE",.sprSht1.ActiveRow, strCLIENTSUBCODE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, strCLIENTCODE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"SUBSEQ",.sprSht1.ActiveRow, strSUBSEQ
		End If
	End with
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
	

End Sub

Sub sprSht1_Keydown(KeyCode, Shift)
Dim intRtn
if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
	
	
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 8 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(13), cint(Shift), -1, 1)
		'intiSprValue
		end if
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if

End sub
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
		
				
     		'GetBrandDefaultFind	
     			
			
			.txtCLIENTSUBNAME.focus()					' 포커스 이동
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
     	
	End with

	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	
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
' 사업부코드팝업 버튼[조회용]
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
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' 코드명 표시
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME.value = trim(vntRet(6,0))
			
		
			
			.txtJOBNAME.focus()					' 포커스 이동
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
			vntData = mobjPDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME.value = trim(vntData(6,0))
					
				
					.txtCLIENTSUBNAME.focus()
					gSetChangeFlag .txtCLIENTSUBCODE
					gSetChangeFlag .CLIENTCODE
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

'★★★★★★★★★★★★★ 하단조회시작 ★★★★★★★★★★★★★★★★★★★★★★★★★★
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP1()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP1
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
		
				
     		'GetBrandDefaultFind	
     			
			
			.txtJOBNAME1.focus()					' 포커스 이동
			'gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
     	
	End with

	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
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
					.txtCLIENTCODE1.value = trim(vntData(0,0))
					.txtCLIENTNAME1.value = trim(vntData(1,0))
				
				Else
					Call CLIENTCODE_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' JOB 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgJOBNO1_onclick
	Call SEARCHJOB_POP1()
End Sub

'실제 데이터List 가져오기
Sub SEARCHJOB_POP1
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO1.value), trim(.txtJOBNAME1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO1.value = vntRet(0,0) and .txtJOBNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtJOBNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtJOBNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtJOBNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO1.value),trim(.txtJOBNAME1.value))
			if not gDoErrorRtn ("txtJOBNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO1.value = trim(vntData(0,0))
					.txtJOBNAME1.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'★★★★★★★★★★★★★ 하단조회끝 ★★★★★★★★★★★★★★★★★★★★★★★★★★
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
	Dim datacnt
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	Dim intCnt2
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht1.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
'	For intCnt2 = 1 To frmThis.sprSht1.MaxRows
'		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXNO",intCnt2) <> "" THEN
'			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSNO",intCnt2) & " 에 대하여" &vbcrlf & "세금계산서번호가 존재하는 내역은 재출력할 수 없습니다.","인쇄안내!"
'			Exit Sub
'		End If
'	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 PD_CHARGE_TEMP테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다.
		'PD_CHARGE_TEMP삭제 시작
		intRtn = mobjPDCMPREESTLIST.DeleteRtn_temp(gstrConfigXml)
		'PD_CHARGE_TEMP삭제 끝
		
		ModuleDir = "PD"
		ReportName = "PDCMCHARGE.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",.sprSht1.activeRow)
		
		vntData = mobjPDCMPREESTLIST.Get_PREEST_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strPREESTNO)
	
		strcntsum = 0
		IF not gDoErrorRtn ("Get_PREEST_CNT") then
			datacnt = mlngRowCnt
			
			for i=1 to 3
				strUSERID = ""
				vntDataTemp = mobjPDCMPREESTLIST.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
				
		window.setTimeout "printSetTimeout", 10000
	
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMPREESTLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub



Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub






'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------




'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************

sub sprSht_DblClick (ByVal Col, ByVal Row)
Dim strJOBNO	
Dim vntInParams
Dim vntRet
Dim strRow
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row))
			vntRet = gShowModalWindow("PDCMESTDTL.aspx",vntInParams , 1060,780)
			strRow = Row
			'여기서 부터 실제 견적 화면 호출
			.txtCLIENTSUBNAME.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			
			SelectRtn
			SelectRtn_DBLHDR(strJOBNO)
			mobjSCGLSpr.ActiveCell .sprSht, Col, strRow			
		end if
	end with
end sub
Sub sprSht_Click(ByVal Col, ByVal Row)
Dim strJOBNO	
Dim vntInParams
Dim vntRet
Dim strRow
with frmThis
	mstrCHKROW = True
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			SelectRtn_DBLHDR(strJOBNO)
				
End with
End Sub
sub sprSht1_DblClick (ByVal Col, ByVal Row)
Dim strJOBNO	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		Else
			Dim vntInParams
			Dim vntRet
			Dim strRow
				with frmThis
					if mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow) = "" then
							gErrorMsgBox "가견적 List 를 우선저장하시고 입력 하십시오.","처리안내" 
							Exit Sub
						End if
						vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow))
						vntRet = gShowModalWindow("PDCMPREESTDTL.aspx",vntInParams , 1060,780)
						
						
						.txtCLIENTSUBNAME.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
						.sprSht1.Focus
						strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
						strRow = .sprSht.ActiveRow
						SelectRtn
						SelectRtn_DBLHDR(strJOBNO)
						mobjSCGLSpr.ActiveCell .sprSht, 1, strRow	
				End with
		end if
	end with
end sub

Sub sprSht1_Change(ByVal Col, ByVal Row)
Dim vntData
Dim i, strCols
Dim strCode, strCodeName

	with frmThis
				'Long Type의 ByRef 변수의 초기화
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				strCode = ""
				strCodeName = ""
				
				IF Col = 3 Then
						
					strCode = ""
					strCode		= mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow)
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNAME",.sprSht1.ActiveRow)
					
					vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)
					
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",.sprSht1.ActiveRow, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",.sprSht1.ActiveRow, vntData(1,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",.sprSht1.ActiveRow, vntData(2,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",.sprSht1.ActiveRow, vntData(3,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, vntData(4,0)
						mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",.sprSht1.ActiveRow, vntData(5,0)			
						mobjSCGLSpr.CellChanged .sprSht1, .sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					Else
						mobjSCGLSpr_ClickProc .sprSht1, Col, .sprSht1.ActiveRow
					End If
				END IF
				
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col, Row
End Sub
Sub mobjSCGLSpr_ClickProc(sprSht1, Col, Row)
	dim vntRet, vntInParams
	With frmThis
	
		IF Col = 3 Then
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( sprSht1,"JOBNAME",Row))
			vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(5,0)	
				mobjSCGLSpr.CellChanged sprSht1, Col,Row
				
				
			End IF
			.txtCLIENTSUBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus	
		
		end if
	End With
End Sub


Sub sprSht1_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strJOBNO
	Dim strRow
	with frmThis
		IF Col = 2 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") then exit Sub
			if mobjSCGLSpr.GetTextBinding( .sprSht1,"CONF",Row) = "Y" Then exit Sub	
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",Row), mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNAME",Row))
			vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(5,0)				
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtCLIENTSUBNAME.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus
			mobjSCGLSpr.ActiveCell .sprSht1, Col+3, Row
		
		
		end if
	end with
end SUB
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
	set mobjPDCMPREESTLIST	= gCreateRemoteObject("cPDCO.ccPDCOPREESTLIST")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
    set mobjPDCMJOBNO       = gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "160px"
	'pnlTab1.style.left= "7px"
	
	'pnlTab2.style.position = "absolute"
	'pnlTab2.style.top = "693px"
	'pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'가견적요청리스트
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 18, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "REQDAY|PROJECTNO|JOBNO|JOBNAME|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|ENDFLAG|JOBGUBN|CREPART|CREGUBN|COMMITION|CLIENTCODE|PREESTNO|AMT|DEMANDYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		   "의뢰일|프로젝트번호|JOBNO|JOB명|광고주|사업부|사업부명|브랜드|브랜드명|상태|매체부분|매체분류|신규|수수료율|광고주코드|확정견적코드|견적금액|청구일자"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "10    | 0          |7    |   19|13    |6     |12      |6     |13      |   6|12      |12      |6   |0       |0         |0           |11		|10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|DEMANDYEARMON", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PROJECTNO|JOBNO|JOBNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|JOBGUBN|CREPART|CREGUBN|REQDAY|ENDFLAG|CLIENTNAME|PREESTNO|AMT|DEMANDYEARMON"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME|MEDNAME|PROGRAM_NAME|PUB_FACENAME|COL_DEG ", -1, -1, 100
		mobjSCGLSpr.ColHidden .sprSht, "PROJECTNO|COMMITION|CLIENTCODE|PREESTNO|CLIENTSUBCODE|SUBSEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTSUBNAME|SUBSEQNAME|CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTSUBCODE|SUBSEQ|JOBGUBN|CREPART|CREGUBN|JOBNO|ENDFLAG|DEMANDYEARMON",-1,-1,2,2,false
		
		
	    
	    '******************************************************************
		'가견적리스트
		'******************************************************************
	    gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 13, 0, 0, 0,2
		mobjSCGLSpr.AddCellSpan  .sprSht1, 1, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1, "JOBNO|BTN|JOBNAME|PREESTNO|PREESTNAME|AMT|MEMO|CONF|CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ"
		mobjSCGLSpr.SetHeader .sprSht1,		"JOBNO|JOB명|가견적코드|가견적명|금액|비고|확정여부|의뢰일|사업부|커미션|광고주코드|브랜드코드"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "6|2|20|9|28|12|35|10|0|0|0|10|0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "SUSURATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "JOBNO|JOBNAME|PREESTNAME|MEMO", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht1, true, "PREESTNO|CONF|BTN|SUSURATE|JOBNO|JOBNAME|PREESTNAME|MEMO|CONF|AMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "PREESTNO|CONF",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht1, "CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ", true
		mobjSCGLSpr.SetScrollBar .sprSht1,2,False,0,-1
	
	    		
    End With    
	'pnlTab1.style.visibility = "visible"
	'pnlTab2.style.visibility = "visible"
	'화면 초기값 설정
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'기본값 설정
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'WITH frmThis
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
	'end with
	'SelectRtn
	Call SEARCHCOMBO_TYPE()
End Sub
'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
    With frmThis   

		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB종류 호출
		vntENDFLAG = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '제작상태 호출
		if not gDoErrorRtn ("COMBO_TYPE") then 
			 gLoadComboBox .cmbSEARCHENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbSEARCHJOBGUBN, vntJOBGUBN, False
   		end if    				   		
   	end with     	
End Sub

Sub EndPage()
	set mobjPDCMPREESTLIST = Nothing
	set mobjPDCMGET = Nothing
	set mobjPDCMJOBNO = Nothing
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
		
		'.txtCREDAY.value = gNowDate
		
		
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
	End with
	'DataNewClean
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub
Sub DateClean2
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM1.value = date1
		.txtTO1.value = date2
	End With
End Sub
'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
  	Dim vntData
  	Dim strRow
	with frmThis
	'On error resume next
  		'데이터 Validation
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"JOBNO|JOBNAME|PREESTNO|PREESTNAME|AMT|MEMO|CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		if DataValidation =false then exit sub
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .sprSht1.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
		
		intRtn = mobjPDCMPREESTLIST.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가" & intRtn & " 건 저장" & mePROC_DONE,"저장안내" 
			strRow = .sprSht.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
  		end if
 	end with
End Sub
'****************************************************************************************
' 일자조회 달력
'****************************************************************************************
'조회용
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM1,frmThis.imgCalEndarFROM1,"txtFROM1_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo1,frmThis.imgCalEndarTO1,"txtTo1_onchange()"
		gSetChange
	end with
End Sub
'조회용
Sub imgCalEndarFROM_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub

Sub txtFROM1_onchange
	gSetChange
End Sub


Sub txtTo1_onchange
	gSetChange
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
		for intCnt = 1 to .sprSht1.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNO",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNAME",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 제작번호 및 가견적명 을 확인하십시오","입력오류"
				Exit Function
			End if
		next	
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' 데이터 조회
'****************************************************************************************

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	
	On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		
		
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		'세금계산서 완료조회
		vntData = mobjPDCMPREESTLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),Trim(.txtJOBNO.value),Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
			ELSE
				
			End If
			Call SelectRtn_HDR()
			'gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			'sprShtToFieldBinding 1,1
			Call sprSht_Click(1,1)
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub
',Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value
Sub SelectRtn_HDR ()
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
	strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),Trim(.txtJOBNO.value),Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					Else
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB
Sub SelectRtn_HDR2 ()
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM1.value,1,4) &  MID(.txtFROM1.value,6,2) &  MID(.txtFROM1.value,9,2)
	strTO =  MID(.txtTO1.value,1,4) &  MID(.txtTO1.value,6,2) &  MID(.txtTO1.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR2(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME1.value),Trim(.txtJOBNO1.value),Trim(.txtCLIENTCODE1.value),Trim(.txtCLIENTNAME1.value))
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					
					Else
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB

Sub SelectRtn_DBLHDR (ByVal strJOBNO)
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
	strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),strJOBNO,Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0	
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					Else
					'mobjSCGLSpr.SetCellsLock2 .sprSht1, False, "JOBNO|JOBNAME"
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB





'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
	Dim strRow
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		
		'PREESTNO,ITEMCODESEQ
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",vntData(i)) = "Y" Then
				gErrorMsgBox "확정견적은 삭제하실수 없으며, 상세내역에서 확정을 취소후 삭제하십시오.","삭제안내"
				Exit Sub
			End if
			intRtn = gYesNoMsgbox("자료는 상세내역 과 함께 삭제 됩니다. " & vbcrlf & "자료를 삭제하시겠습니까?","자료삭제 확인")
			IF intRtn <> vbYes then exit Sub
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i)) <> "" Then
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i))
				intRtn = mobjPDCMPREESTLIST.DeleteRtn(gstrConfigXml,strPREESTNO)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				gWriteText "", "[" & strPREESTNO & "] 자료가 삭제되었습니다."
   			End IF
		next
		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		strRow = .sprSht.ActiveRow
		SelectRtn
		mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		Call sprSht_Click(1,strRow)
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
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<!--Top TR Start-->
				<TBODY>
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
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;견적관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
										width="85">광고주</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" size="20" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" size="5" name="txtCLIENTCODE"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
										width="85">JOB명</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><INPUT class="INPUT_L" id="txtJOBNAME" title="제작관리명 조회" style="WIDTH: 144px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="18" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="제작관리코드 조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="7" align="left" size="3" name="txtJOBNO">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="85">의뢰일자</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtFROM" title="기간검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
											accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
											border="0" name="imgCalEndarFROM">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="기간검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgCalEndarTO"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)"
										width="85">사업부</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="사업부명 조회" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="20" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT" id="txtCLIENTSUBCODE" title="사업부코드 조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtCLIENTSUBCODE"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
										width="85">매체부문</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><SELECT id="cmbSEARCHJOBGUBN" title="매체부문조회" style="WIDTH: 232px" name="cmbSEARCHJOBGUBN"></SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="85">제작진행상태</TD>
									<TD class="SEARCHDATA" colSpan="2"><SELECT id="cmbSEARCHENDFLAG" title="완료구분" style="WIDTH: 216px" name="cmbSEARCHENDFLAG"></SELECT></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
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
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
						
							<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="11721">
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
							</DIV>
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
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
												<td class="TITLE">&nbsp;청구 견적 리스트</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgDetail" onmouseover="JavaScript:this.src='../../../images/imgDetailOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDetail.gif'"
														height="20" alt="자료의 상세내역을 관리합니다." src="../../../images/imgDetail.gIF" border="0"
														name="imgDetail"></TD>
												<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
														src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
												<td><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gIF'"
														height="20" alt="자료를 복사합니다." src="../../../images/imglistcopy.gIF" width="77" border="0"
														name="Imgcopy"></td>
												<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="선택한 행을삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgRowDel"></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
												<TD><IMG id="imgExcel1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel1"></TD>
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
							<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
										width="85">광고주</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="광고주명" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" size="20" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE1"><INPUT class="INPUT" id="txtCLIENTCODE1" title="광고주코드" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" size="5" name="txtCLIENTCODE1"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME1, txtJOBNO1)"
										width="85">JOB명</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><INPUT class="INPUT_L" id="txtJOBNAME1" title="제작관리명 조회" style="WIDTH: 144px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="18" name="txtJOBNAME1"><IMG id="ImgJOBNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgJOBNO1"><INPUT class="INPUT" id="txtJOBNO1" title="제작관리코드 조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtJOBNO1">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean2()"
										width="85">의뢰일자</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtFROM1" title="기간검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
											accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM1"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
											border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO1" title="기간검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="7" name="txtTO1"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery1" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery1"></td>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
							<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									 VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="3889">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;합 
							계 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="금액" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
