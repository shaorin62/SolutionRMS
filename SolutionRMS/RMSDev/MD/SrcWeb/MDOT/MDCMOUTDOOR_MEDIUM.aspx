<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOOR_MEDIUM.aspx.vb" Inherits="MD.MDCMOUTDOOR_MEDIUM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>옥외 개별청약 등록/조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : MD/부킹 화면(MDCMBOOKING)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMBOOKING.aspx
'기      능 : 인쇄매체 Booking Process 처리
'파라  메터 : 
'특이  사항 : 복사처리(다중선택 Row Coyp)
'----------------------------------------------------------------------------------------
'HISTORY    :1) Old Ver. Kim Tae Yup
'			 2) 2008/08/14 By Kim Tae Ho
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjOUTDOOR_MEDIUM, mobjMDCOGET 
Dim mstrCheck
Dim mcomecalender, mcomecalender1, mcomecalender2
Dim mstrPROCESS	'신규이면 True 조회면 False
Dim mstrPROCESS2 '조회상태이면 True 신규상12태이면 False
Dim mstrHIDDEN

CONST meTAB = 9
mstrPROCESS = False
mstrPROCESS2 = True
mstrCheck = True
mcomecalender = FALSE
mcomecalender1 = FALSE
mcomecalender2 = FALSE
mstrHIDDEN = 0
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
'입력 필드 숨기기
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody").style.display = "inline"
			document.getElementById("tblSheet").style.height = "65%"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody").style.display = "none"
			document.getElementById("tblSheet").style.height = "82%"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
'조회버튼
Sub imgQuery_onclick
	If frmThis.txtYEARMON1.value = "" Then
		gErrorMsgBox "조회년월을 입력하시오","조회안내"
		exit Sub
	End If
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

'신규버튼
Sub imgREG_onclick ()
	Call sprSht_Keydown(meINS_ROW, 0)	
	mstrPROCESS = False
end Sub

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
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		'mobjSCGLSpr.ExportComboType = "2"
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' 내역복사한다.
'-----------------------------------------------------------------------------------------
Sub Imgcopy_onclick ()
	Dim intRtn
   	Dim vntData
	Dim intSelCnt,  i
	Dim strYEARMON, strGUBUN, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strREAL_MED_CODE, strREAL_MED_NAME, strREAL_MED_BISNO
	Dim strMED_FLAG, strDEMANDDAY, strTBRDSTDATE
	Dim strTBRDEDDATE, strGBN_FLAG, strTITLE, strMATTERNAME, strTOTALAMT, strAMT, strOUT_AMT
	Dim strCOMMI_RATE, strCOMMISSION, strMED_GBN, strLOCATION, strMEMO
	
	With frmThis
		intSelCnt = 0
		
		Dim strCNT, strCNT2
		strCNT2 = 0
		For i=1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				strCNT = i
				strCNT2 = strCNT2 +1
			End If
		Next
		If strCNT2 >1 Then
			gErrorMsgBox "내역복사는 한건만 가능합니다.",""
			Exit Sub
		elseif strCNT2 =0 Then
			gErrorMsgBox "내역복사할 로우를 선택하시오.",""
			Exit Sub
		elseif strCNT2 = 1 Then
			If mstrPROCESS Then
				for i = .sprSht.MaxRows to 1 step -1
					If strCNT = i Then
					else 
						mobjSCGLSpr.DeleteRow .sprSht,i
					End If
				Next
			End If
		End If
		
		strYEARMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",.sprSht.ActiveRow)
		strGUBUN			=	"미승인"
		strCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		strCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",.sprSht.ActiveRow)
		strTIMCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",.sprSht.ActiveRow)
		strTIMNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",.sprSht.ActiveRow)
		strREAL_MED_CODE	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",.sprSht.ActiveRow)
		strREAL_MED_NAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",.sprSht.ActiveRow)
		strREAL_MED_BISNO	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",.sprSht.ActiveRow)	
		strMED_FLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",.sprSht.ActiveRow)
		strDEMANDDAY		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",.sprSht.ActiveRow)
		strTBRDSTDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",.sprSht.ActiveRow)
		strTBRDEDDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",.sprSht.ActiveRow)
		strGBN_FLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"GBN_FLAG",.sprSht.ActiveRow)
		strTITLE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TITLE",.sprSht.ActiveRow)
		strMATTERNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",.sprSht.ActiveRow)
		strTOTALAMT			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TOTALAMT",.sprSht.ActiveRow)
		strAMT				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",.sprSht.ActiveRow)
		strOUT_AMT			=	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",.sprSht.ActiveRow)
		strCOMMI_RATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",.sprSht.ActiveRow)
		strCOMMISSION		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",.sprSht.ActiveRow)
		strMED_GBN			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_GBN",.sprSht.ActiveRow)
		strLOCATION			=	mobjSCGLSpr.GetTextBinding(.sprSht,"LOCATION",.sprSht.ActiveRow)
		strMEMO				=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",.sprSht.ActiveRow)
		
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, meINS_ROW, 0, -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strYEARMON
		mobjSCGLSpr.SetTextBinding .sprSht,"GUBUN",.sprSht.ActiveRow, strGUBUN
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, strCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, strCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, strTIMCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",.sprSht.ActiveRow, strTIMNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, strREAL_MED_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, strREAL_MED_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_BISNO",.sprSht.ActiveRow, strREAL_MED_BISNO
		mobjSCGLSpr.SetTextBinding .sprSht,"MED_FLAG",.sprSht.ActiveRow, strMED_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",.sprSht.ActiveRow, strTBRDEDDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"GBN_FLAG",.sprSht.ActiveRow, strGBN_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"TITLE",.sprSht.ActiveRow, strTITLE		
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",.sprSht.ActiveRow, strMATTERNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TOTALAMT",.sprSht.ActiveRow, strTOTALAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"OUT_AMT",.sprSht.ActiveRow, strOUT_AMT
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, strCOMMI_RATE
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",.sprSht.ActiveRow, strCOMMISSION
		mobjSCGLSpr.SetTextBinding .sprSht,"MED_GBN",.sprSht.ActiveRow, strMED_GBN
		mobjSCGLSpr.SetTextBinding .sprSht,"LOCATION",.sprSht.ActiveRow, strLOCATION
		mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",.sprSht.ActiveRow, strMEMO
		
		gXMLSetFlag xmlBind, meUPD_TRANS
		mstrPROCESS = False
   	end With
end Sub

'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
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
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "D")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub



'매체사 팝업 버튼
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE_POP()
End Sub


'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "D")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'팁 팝업 버튼
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub


'실제 데이터List 가져오기
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE1.value = vntRet(0,0) and .txtTIMNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTIMCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtTIMNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtCLIENTCODE1.value = trim(vntRet(4,0))       ' 코드명 표시
			.txtCLIENTNAME1.value = trim(vntRet(5,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
											trim(.txtTIMCODE1.value),trim(.txtTIMNAME1.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME1.value = trim(vntData(1,1))       ' 코드명 표시
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(5,1))
					SelectRtn
				Else
					Call TIMCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'----------------------------------------------------
'입력용 
'----------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "D")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'매체사 입력용
Sub imgREAL_MED_CODE_onclick
	Call REAL_MED_CODE1_POP()
End Sub

Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' 코드명 표시
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_BISNO",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "D")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,1))
					.txtREAL_MED_NAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call REAL_MED_CODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub



'팁 팝업 버튼
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
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP_ALL.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtTIMNAME.value = trim(vntRet(1,0))       ' 코드명 표시.
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' 코드명 표시
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' 코드명 표시
					
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
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
			vntData = mobjMDCOGET.GetTIMCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
									 		trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE_ALL") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME.value = trim(vntData(1,1))       ' 코드명 표시
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))	
						
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call TIMCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 달력
'****************************************************************************************
'청구일
Sub imgCalEndar_onclick
	'CalEndar를 화면에 표시
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar,"txtDEMANDDAY_onchange()"
	mcomecalender = false
	gXMLDataChanged xmlBind         
End Sub

'시작일
Sub imgCalEndar1_onclick
	mcomecalender1 = true
	gShowPopupCalEndar frmThis.txtTBRDSTDATE,frmThis.imgCalEndar1,"txtTBRDSTDATE_onchange()"
	mcomecalender1 = false
	gXMLDataChanged xmlBind
End Sub

'종료일
Sub imgCalEndar2_onclick
	mcomecalender2 = true
	gShowPopupCalEndar frmThis.txtTBRDEDDATE,frmThis.imgCalEndar2,"txtTBRDEDDATE_onchange()"
	mcomecalender2 = false
	gXMLDataChanged xmlBind
End Sub

'****************************************************************************************
' 입력필드 키다운 이벤트
'****************************************************************************************
Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
	'청구일세팅 게재월의 마지막일
		DateClean frmThis.txtYEARMON.value
		
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEMANDDAY_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTITLE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtTBRDSTDATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTBRDEDDATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtTBRDEDDATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTOTALAMT.focus()	
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTOTALAMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtLOCATION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtREAL_MED_CODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTITLE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMATTERNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTIMCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtREAL_MED_NAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtMATTERNAME_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMED_GBN.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtLOCATION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMEMO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTIMNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtMED_GBN_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTBRDSTDATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtMEMO_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'금액에서 엔터시 수수료 자동계산
Sub txtAMT_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMISSION_Cal
		COMMI_RATE_Cal
		frmThis.txtOUT_AMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'외주비에서 엔터시 수수료 자동계산
Sub txtOUT_AMT_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMISSION_Cal
		COMMI_RATE_Cal
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'수수료율에서 엔터시 수수료 자동계산
Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 입력필드 체인지 이벤트
'****************************************************************************************
Sub txtDEMANDDAY_onchange
	Dim strdate 
	Dim strDEMANDDAY
	strdate = ""
	strDEMANDDAY =""
	With frmThis
		strdate=.txtDEMANDDAY.value
	
		If mcomecalender Then
			strDEMANDDAY = strdate
		else
			If len(strdate) = 4 Then
				strDEMANDDAY = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strDEMANDDAY = strdate
			elseif len(strdate) = 3 Then
				strDEMANDDAY = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strDEMANDDAY = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtTBRDSTDATE_onchange
	Dim strdate 
	Dim strTBRDSTDATE, strTBRDSTDATE2
	Dim strOLDYEARMON
	strdate = ""
	strTBRDSTDATE =""
	strTBRDSTDATE2 = ""

	With frmThis
		strdate=.txtTBRDSTDATE.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender1 Then
			strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strTBRDSTDATE2 = strdate
		else
			If len(strdate) = 4 Then
				strTBRDSTDATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strTBRDSTDATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strTBRDSTDATE2 = strdate
			elseif len(strdate) = 3 Then
				strTBRDSTDATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strTBRDSTDATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strTBRDSTDATE2 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE2
			DateClean_TBRDSTDATE strTBRDSTDATE
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With

	gSetChange
End Sub

Sub txtTBRDEDDATE_onchange
	Dim strdate 
	Dim TBRDEDDATE, TBRDEDDATE2
	Dim strOLDYEARMON
	strdate = ""
	TBRDEDDATE =""
	TBRDEDDATE2 = ""

	With frmThis
		strdate=.txtTBRDEDDATE.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender2 Then
			TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			TBRDEDDATE2 = strdate
		else
			If len(strdate) = 4 Then
				TBRDEDDATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				TBRDEDDATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				TBRDEDDATE2 = strdate
			elseif len(strdate) = 3 Then
				TBRDEDDATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				TBRDEDDATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				TBRDEDDATE2 = strdate
			End If
		End If
		
		If frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, TBRDEDDATE2
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		End If
	END With
End Sub


Sub txtMATTERNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERNAME",frmThis.sprSht.ActiveRow, frmThis.txtMATTERNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtTIMNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, frmThis.txtTIMNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtTIMCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, frmThis.txtTIMCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtREAL_MED_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_CODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_CODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCOMMI_RATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, frmThis.txtCOMMI_RATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCOMMISSION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtYEARMON_onchange	
	DateClean frmThis.txtYEARMON.value
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub


Sub txtREAL_MED_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_CODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_CODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtTIMNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, frmThis.txtTIMNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtTIMCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, frmThis.txtTIMCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub


Sub txtMEMO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtMED_GBN_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_GBN",frmThis.sprSht.ActiveRow, frmThis.txtMED_GBN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtLOCATION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"LOCATION",frmThis.sprSht.ActiveRow, frmThis.txtLOCATION.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub


Sub txtOUT_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, frmThis.txtOUT_AMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtTOTALAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TOTALAMT",frmThis.sprSht.ActiveRow, frmThis.txtTOTALAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub


Sub txtTITLE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TITLE",frmThis.sprSht.ActiveRow, frmThis.txtTITLE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub chkAORFLAG_onchange
	Dim strVOCH_TYPE
	WITH frmThis
		If .sprSht.ActiveRow >0 Then
			if .chkAORFLAG.checked = true then
				strVOCH_TYPE = 3
			else
				strVOCH_TYPE = 2	
			end if 
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, strVOCH_TYPE
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end With
End Sub
'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'금액
Sub txtAMT_onblur
	With frmThis
		COMMISSION_Cal
		COMMI_RATE_Cal
		Call gFormatNumber(.txtAMT,0,True)
	end With
End Sub

'금액
Sub txtOUT_AMT_onblur
	With frmThis
		COMMISSION_Cal
		COMMI_RATE_Cal
		Call gFormatNumber(.txtOUT_AMT,0,True)
	end With
End Sub

'금액
Sub txtCOMMISSION_onblur
	With frmThis
		Call gFormatNumber(.txtCOMMISSION,0,True)
	end With
End Sub

'-----------------------------------------------------------------------------------------
' 천단위 나눔점 없애기 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'금액
Sub txtAMT_onfocus
	With frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end With
End Sub

'총계약금액
Sub txtTOTALAMT_onfocus
	With frmThis
		.txtTOTALAMT.value = Replace(.txtTOTALAMT.value,",","")
	end With
End Sub

'수수료
Sub txtCOMMISSION_onfocus
	With frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end With
End Sub

'외주비
Sub txtOUT_AMT_onfocus
	With frmThis
		.txtOUT_AMT.value = Replace(.txtOUT_AMT.value,",","")
	end With
End Sub


'****************************************************************************************
' 수수료 계산
'****************************************************************************************
Sub COMMISSION_Cal
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT
	Dim intOUT_AMT
	
	With frmThis
		
		intAMT = .txtAMT.value
		intOUT_AMT = .txtOUT_AMT.value
		IF intOUT_AMT = "" THEN
		intOUT_AMT = 0
		END IF
		
		If intAMT= "" Then  Exit Sub
			
		.txtCOMMISSION.value = intAMT - intOUT_AMT 
		
		
		txtCOMMISSION_onchange
		
		gSetChangeFlag .txtAMT
		gSetChangeFlag .txtOUT_AMT
		gSetChangeFlag .txtCOMMISSION
		
	End With
End Sub

'수수료율(내수율) 자동계산
Sub COMMI_RATE_Cal
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,intCOMMISSION,intOUT_AMT, dblCOMMI_RATE
	
	With frmThis
		If .txtAMT.value = "" then Exit Sub
		If .txtOUT_AMT.value = "" then Exit Sub
		
		intAMT		= int(.txtAMT.value)
		intOUT_AMT	= int(.txtOUT_AMT.value)
		
		If intAMT <> 0  Then
			dblCOMMI_RATE = gRound((intAMT -  intOUT_AMT) / intAMT * 100,2)  '월청구액 - 외주비 /월청구액 * 100 = 내수율
			.txtCOMMI_RATE.value = dblCOMMI_RATE
		End If
		
		txtCOMMI_RATE_onchange
		txtCOMMISSION_onchange
		
		gSetChangeFlag .txtAMT
		gSetChangeFlag .txtOUT_AMT
		gSetChangeFlag .txtCOMMI_RATE
		gSetChangeFlag .txtCOMMISSION
	End With
End Sub


'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If mstrPROCESS = True Then
			frmThis.sprSht.MaxRows = 0
		End If
		
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEMANDDAY",frmThis.sprSht.ActiveRow, frmThis.txtDEMANDDAY.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDSTDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDEDDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, ""
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		strRow = frmThis.sprSht.ActiveRow
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,FALSE,"YEARMON",1,strRow,FALSE
		
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus
		
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
	sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") OR _
		   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
		   
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strCOLUMN = "COMMISSION"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") Then
				strCOLUMN = "TOTALAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
				strCOLUMN = "OUT_AMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) OR _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

'마우스 금액 계산
Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	Dim strColFlag
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
		strCOLUMN = ""
		strColFlag = 0
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or  _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
						End If
					Next
					
					.txtSELECTAMT.value = strSUM
				End If
				
			else
				.txtSELECTAMT.value = 0
			End If
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	Dim strSTD_STEP, strSTD_CM, strSTD_FACE, strSTD_PAGE, strPRICE
   	Dim strAMT
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON")  Then 
			.txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
			call DateClean_SHEET(.txtYEARMON.value ,Row)
			
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDDAY") Then .txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE")  Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
															  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						.txtCLIENTCODE.value = vntData(0,1)
						.txtCLIENTNAME.value = vntData(1,1)
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMCODE")  Then .txtTIMCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetTIMCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, trim(vntData(5,1))
						
						.txtTIMCODE.value = trim(vntData(0,1))	    ' Code값 저장
						.txtTIMNAME.value = trim(vntData(1,1))       ' 코드명 표시
						.txtCLIENTCODE.value = trim(vntData(4,1))
						.txtCLIENTNAME.value = trim(vntData(5,1))
			
						.txtTIMNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtTIMNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_CODE")  Then .txtREAL_MED_CODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, ""
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(strCode),trim(strCodeName), "D")
			
				If not gDoErrorRtn ("mobjMDCOGET.GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_BISNO",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						
						.txtREAL_MED_CODE.value = trim(vntData(0,1))
						.txtREAL_MED_NAME.value = trim(vntData(1,1))	
						
						.txtREAL_MED_NAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME"), Row
						.txtREAL_MED_NAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End IF
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TITLE") Then 
			.txtTITLE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TITLE",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then 
			.txtMATTERNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MED_GBN") Then 
			.txtMED_GBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MED_GBN",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"LOCATION") Then 
			.txtLOCATION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"LOCATION",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO") Then 
			.txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT"), Row)
			.txtTOTALAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TOTALAMT",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"), Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then 
		    Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT"), Row)
			.txtOUT_AMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE"), Row)
			.txtCOMMI_RATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION"), Row)
			.txtCOMMISSION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TBRDSTDATE")  Then 
			Dim strdate
			Dim strPUB_DATE
			Dim strYEARMON
			
			strdate = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
			strYEARMON = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			
			
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",Row, strdate
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",Row, strYEARMON
			
			.txtTBRDSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
			.txtTBRDEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		End If
		
		'시트의 담당부서명 변경되면
   		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)
																								  

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.sprSht.focus 
					End If
   				End If
   			End If
   		end if 
		
	
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub SHEET_COMMI_RATE_Cal (Col, Row)
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,intOUT_AMT
	Dim dblCOMMI_RATE
	Dim intCOMMISSION
	With frmThis
	
		If Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
			If intAMT <> "" AND intOUT_AMT <> "" Then
				intCOMMISSION = intAMT - intOUT_AMT
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
				dblCOMMI_RATE = gRound((intCOMMISSION / (intAMT * 100)),2)
   				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
   				.txtCOMMI_RATE.value = dblCOMMI_RATE
   				.txtCOMMISSION.value = intCOMMISSION
   				
			ELSE
				IF intAMT = 0 THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intAMT
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 1
				END IF
				
			End If
		ELSEIF Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
			If intAMT <> 0 AND intOUT_AMT <> 0 Then
				intCOMMISSION = intAMT - intOUT_AMT
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
				dblCOMMI_RATE = gRound((intCOMMISSION / (intAMT * 100)),2)
   				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
   				.txtCOMMI_RATE.value = dblCOMMI_RATE
   				.txtCOMMISSION.value = intCOMMISSION
			ELSE
				IF intAMT = 0 THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intAMT
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 1
				END IF
				
			End If
		End If
	End With
end Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				
				.txtCLIENTCODE.value = vntRet(0,0)		
				.txtCLIENTNAME.value = vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)),"MED_CATV")
			
		    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_BISNO",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				
				.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code값 저장
				.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' 코드명 표시
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
			
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then			
			vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value) , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP_ALL.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(5,0)
				
		
			
				.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code값 저장
				.txtTIMNAME.value = trim(vntRet(1,0))       ' 코드명 표시
				.txtCLIENTCODE.value = trim(vntRet(4,0))    ' 코드명 표시
				.txtCLIENTNAME.value = trim(vntRet(5,0))    ' 코드명 표시
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then			
			vntInParams = array(trim(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If
		
		sprShtToFieldBinding Col, Row
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.txtCLIENTNAME.focus
		.sprSht.Focus
		
	End With
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
		
			sprShtToFieldBinding Col,Row
		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	end With
End Sub

Sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strATTR01
	Dim vntInParams
	Dim vntRet
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub


'시트에 데이터한로우의 정보를 헤더 필더에 바인딩
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '그리드 데이터가 없으면 나간다.
		
		.txtYEARMON.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtDEMANDDAY.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtTIMCODE.value       =   mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		.txtTIMNAME.value       =   mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row)
		.txtREAL_MED_CODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtREAL_MED_NAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtTITLE.value			=   mobjSCGLSpr.GetTextBinding(.sprSht,"TITLE",Row)
		.txtLOCATION.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"LOCATION",Row)
		.txtMATTERNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		.txtMEMO.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		.txtTOTALAMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TOTALAMT",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtOUT_AMT.value	    =	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		.txtCOMMISSION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		.txtCOMMI_RATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtTBRDSTDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
		.txtTBRDEDDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		.txtMED_GBN.value       =   mobjSCGLSpr.GetTextBinding(.sprSht,"MED_GBN",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row) = "3" THEN
			.chkAORFLAG.checked = TRUE
		ELSE
			.chkAORFLAG.checked = FALSE
		END IF 
   	end With
   
	Call gFormatNumber(frmThis.txtAMT,0,True)
	Call gFormatNumber(frmThis.txtCOMMISSION,0,True)
	Call gFormatNumber(frmThis.txtOUT_AMT,0,True)
	Call gFormatNumber(frmThis.txtTOTALAMT,0,True)
	Call Field_Lock ()
End Function

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(trim(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	End With
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성	
	set mobjOUTDOOR_MEDIUM	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR_MEDIUM")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 33, 0, 3, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | DEPT_CD | BTN | DEPT_NAME | MED_FLAG | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | VOCH_TYPE | COMMI_TRANS_NO | TRU_VOCH_NO | ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|년월|번호|구분|광고주코드|광고주|팀코드|팀|외주처코드|외주처|사업자번호|담당부서코드|담당부서명|매체구분코드|청구일자|계약시작일|계약종료일|매출구분|계약명|소재명|총계약금액|월청구금액|월외주비|내수율|내수액|제작종류|장소|비고|청구구분|거래명세서번호|매입전표번호|상세번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   5|   0|   5|         0|    13|     0|10|         0|    13|        13|           0|2|       8|           0|       8|         8|         8|       9|    15|    15|        10|        10|      10|     6|     9|      10|  10|  10|       0|            10|           0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"☆", "BTN"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | DEPT_CD | DEPT_NAME | MED_FLAG | GBN_FLAG | TITLE | MATTERNAME | MED_GBN | LOCATION | MEMO | ATTR01", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | TOTALAMT | AMT | OUT_AMT | COMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, " SEQ | COMMI_TRANS_NO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN | DEMANDDAY | TBRDSTDATE | TBRDEDDATE",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | CLIENTCODE | TIMCODE | REAL_MED_CODE | MED_FLAG | GBN_FLAG | TRU_VOCH_NO | VOCH_TYPE | ATTR01", true
		
		.sprSht.style.visibility = "visible"

    End With
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjOUTDOOR_MEDIUM = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	With frmThis
		.sprSht.MaxRows = 0
		
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		.txtYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	'청년월
		
		
		.txtTBRDSTDATE.value = gNowDate2
		DateClean_TBRDSTDATE Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	
		'청구일세팅 게재월의 마지막일
		DateClean .txtYEARMON.value
		
		.txtYEARMON1.focus
		
		Field_Lock
		
	End With
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'청구일 조회조건 생성
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

Sub DateClean_SHEET (strYEARMON, Row)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",Row, date2
	End With
End Sub

'시작일 종료일 조회조건 생성
Sub DateClean_TBRDSTDATE (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
		
	With frmThis
		if strYEARMON = "" THEN EXIT SUB
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		.txtTBRDEDDATE.value = date2
		
		If .sprSht.maxRows > 0 Then
			mobjSCGLSpr.SetTextBinding .sprSht, "TBRDEDDATE" , .sprSht.ActiveRow , .txtTBRDEDDATE.value
		End If
		
	End With
End Sub

'-----------------------------------------------------------------------------------------
' Field_Lock  거래명세서번호나 세금계산서 번호가 있으면 수정할수 없도록 필드를 ReadOnly처리
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",.sprSht.ActiveRow) <> "" Then
				.txtYEARMON.className       = "NOINPUT_L" : .txtYEARMON.readOnly		= True 
			End If
			
			If  mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> ""  Then
				'년도
				.txtYEARMON.className       = "NOINPUT" : .txtYEARMON.readOnly		= True 
				'방영기간
				.txtTBRDSTDATE.className	= "NOINPUT" : .txtTBRDSTDATE.readOnly		= True : .imgCalEndar.disabled	 = True
				.txtTBRDEDDATE.className	= "NOINPUT" : .txtTBRDEDDATE.readOnly		= True : .imgCalEndar1.disabled  = True
				'소재
				.txtMATTERNAME.className	= "NOINPUT_L" : .txtMATTERNAME.readOnly		= True : 
				'광고주
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				'팀
				.txtTIMNAME.className		= "NOINPUT_L" : .txtTIMNAME.readOnly		= True : .ImgTIMCODE.disabled = True
				.txtTIMCODE.className		= "NOINPUT_L" : .txtTIMCODE.readOnly		= True
				
				'계약명
				.txtTITLE.className			= "NOINPUT_L" : .txtTITLE.readOnly		= True
				
				'제작종류
				.txtMED_GBN.className		= "NOINPUT_L" : .txtMED_GBN.readOnly		= True
				'장소 
				.txtLOCATION.className		= "NOINPUT_L" : .txtLOCATION.readOnly		= True
				'청구일
				.txtDEMANDDAY.className		= "NOINPUT"   : .txtDEMANDDAY.readOnly		= True : .imgCalEndar2.disabled  = True 
				
				'매체사
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .imgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
			
				'비고/금액/수수료율/수수료
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly		= True
				.txtTOTALAMT.className		= "NOINPUT_R" : .txtTOTALAMT.readOnly		= True 
				.txtOUT_AMT.className		= "NOINPUT_R" : .txtOUT_AMT.readOnly		= True 
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True 
				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				.chkAORFLAG.disabled = True

			else 
				
				'년도
				.txtYEARMON.className       = "INPUT" : .txtYEARMON.readOnly		= false 
				
				'방영기간
				.txtTBRDSTDATE.className	= "INPUT" : .txtTBRDSTDATE.readOnly	= False : .imgCalEndar.disabled	  = False
				.txtTBRDEDDATE.className	= "INPUT" : .txtTBRDEDDATE.readOnly	= False : .imgCalEndar1.disabled  = False
				
				
				'청구일
				.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
				
				'소재
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : 
				
				'광고주
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
				'팀
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly	= False : .ImgTIMCODE.disabled = False
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly	= False
				
				'계약명
				.txtTITLE.className			= "INPUT_L" : .txtTITLE.readOnly		= False
				
				'제작종류
				.txtMED_GBN.className		= "INPUT_L" : .txtMED_GBN.readOnly		= False
				'장소 
				.txtLOCATION.className		= "INPUT_L" : .txtLOCATION.readOnly		= False
				
				'매체사
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .imgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
				
				'비고/단가/금액/수수료율/수수료
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly		= False
				.txtTOTALAMT.className		= "INPUT_R" : .txtTOTALAMT.readOnly		= False
				.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly		= False 
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
				.chkAORFLAG.disabled = False
			End If
		else
			'년도
			.txtYEARMON.className       = "INPUT" : .txtYEARMON.readOnly		= False 
			
			'방영기간
			.txtTBRDSTDATE.className	= "INPUT" : .txtTBRDSTDATE.readOnly	= False : .imgCalEndar.disabled	  = False
			.txtTBRDEDDATE.className	= "INPUT" : .txtTBRDEDDATE.readOnly	= False : .imgCalEndar1.disabled  = False
			'소재
			.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : 
			
			'제작종류
			.txtMED_GBN.className		= "INPUT_L" : .txtMED_GBN.readOnly		= False
			'장소 
			.txtLOCATION.className		= "INPUT_L" : .txtLOCATION.readOnly		= False
			
			'계약명
			.txtTITLE.className			= "INPUT_L" : .txtTITLE.readOnly		= False
			
			'광고주
			.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
			.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
			'팀
			.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly	= False : .ImgTIMCODE.disabled = False
			.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly	= False
			'청구일
			.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
			'매체사
			.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .imgREAL_MED_CODE.disabled = False
			.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
			
			'비고/단가/금액/수수료율/수수료
			.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
			.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly		= False
			.txtTOTALAMT.className		= "INPUT_R" : .txtTOTALAMT.readOnly		= False
			.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly		= False 
			.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
			.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
			.chkAORFLAG.disabled = False
		End If
	End With
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim strYEARMON, strCLIENTCODE,strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strTIMCODE, strTIMNAME, strTITLE
   	Dim strGUBUN
   	Dim intCnt2, strRows
	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		 = .txtYEARMON1.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCLIENTNAME	 = .txtCLIENTNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value
		strTIMCODE		 = .txtTIMCODE1.value
		strTIMNAME		 = .txtTIMNAME1.value
		strTITLE		 = .txtTITLE1.value
		

		vntData = mobjOUTDOOR_MEDIUM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, _
												strCLIENTCODE, _
												strREAL_MED_CODE, _
												strTIMCODE, strTITLE)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				
	   			For intCnt = 1 To .sprSht.MaxRows
	   				If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_VOCH_NO",intCnt) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> ""  Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,32,True
   				'검색시에 첫행을 MASTER와 바인딩 시키기 위함
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				InitPageData
   				'이전 검색어 담아놓기
   				PreSearchFiledValue strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strTITLE
   				.sprSht.MaxRows = 0
   			End If
   		End If
   		mstrPROCESS = True
   	end With
End Sub

'****************************************************************************************
'이전 검색어를 담아 놓는다.
'****************************************************************************************
Sub PreSearchFiledValue (strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strTITLE)
	With frmThis
		.txtYEARMON1.value		= strYEARMON
		.txtCLIENTCODE1.value	= strCLIENTCODE
		.txtCLIENTNAME1.value	= strCLIENTNAME
		.txtREAL_MED_CODE1.value= strREAL_MED_CODE
		.txtREAL_MED_NAME1.value= strREAL_MED_NAME
		.txtTIMCODE1.value		= strTIMCODE
		.txtTIMNAME1.value		= strTIMNAME
		.txtTITLE1.value		= strTITLE
	End With
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strDataCHK
	Dim lngCol, lngRow
	
	With frmThis
   		'데이터 Validation
		'On error resume Next
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "DEMANDDAY | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | TIMCODE | TIMNAME",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 청구일/광고주/매체사/팀는 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | DEPT_CD | BTN | DEPT_NAME | MED_FLAG | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | VOCH_TYPE | COMMI_TRANS_NO | TRU_VOCH_NO")
		
		intRtn = mobjOUTDOOR_MEDIUM.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
   		End If
   	end With
End Sub

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '실제데이터여부 플레
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_VOCH_NO",i) <> "" Then
					gErrorMsgBox "선택하신 " & i & "행의 자료는 거래명세표/매입전표가 존재 합니다." & vbcrlf & "먼저 거래명세표/매입전표를 삭제 하십시오!","삭제안내!"
					exit Sub
				else 
					lngchkCnt = lngchkCnt +1
				End If
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
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjOUTDOOR_MEDIUM.DeleteRtn(gstrConfigXml, strYEARMON,dblSEQ)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
	err.clear	
End Sub

'번호를 클리어한다.
Sub CleanField (objField1, objField2)
	If frmThis.sprSht.MaxRows > 0 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TRU_VOCH_NO",frmThis.sprSht.ActiveRow) = "" and _
		   mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"COMMI_TRANS_NO",frmThis.sprSht.ActiveRow) = "" Then
			
			if isobject(objField1) then 
				objField1.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField1.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			end if
			if isobject(objField2) then 
				objField2.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField2.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			End If
		End If
	ELSE
		if isobject(objField1) then 
			objField1.value = ""
		end if
		if isobject(objField2) then 
			objField2.value = ""
		End If
	End IF
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF"
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
											<td class="TITLE">옥외 청약관리</td>
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
						<TABLE class="SEARCHDATA" id="tblKey" height="48" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
									width="50">청구년월</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
										type="text" maxLength="6" size="10" name="txtYEARMON1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">광고주</TD>
								<TD class="SEARCHDATA" width="250"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 173px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
									<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
									width="50">팀</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="팀명" style="WIDTH: 173px; HEIGHT: 22px" type="text"
										maxLength="100" size="22" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" type="text"
										maxLength="6" size="6" name="txtTIMCODE1"></TD>
								<TD class="SEARCHDATA" width="50">
									<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="SEARCHLABEL" colspan="2"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)">매체사</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="매체사명" style="WIDTH: 173px; HEIGHT: 22px"
										type="text" maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTITLE1, '')">계약명</TD>
								<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtTITLE1" title="계약명" style="WIDTH: 246px; HEIGHT: 22px" type="text"
										maxLength="100" size="36" name="txtTITLE1"></TD>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 20px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="500" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" vAlign="absmiddle"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id='imgTableUp' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableUp.gif'
														align='absMiddle' border='0' name='imgTableUp'></span> &nbsp;&nbsp;&nbsp;&nbsp;합계 
												: <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
												<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="top" align="right" height="28">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="자료를 인쇄합니다." src="../../../images/imgCho.gIF"
													border="0" name="imgCho"></TD>
											<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="자료를 인쇄합니다." src="../../../images/imgNew.gIF"
													border="0" name="imgREG"></TD>
											<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													alt="자료를 인쇄합니다." src="../../../images/imglistcopy.gIF" border="0" name="Imgcopy"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" alt="자료를 인쇄합니다." src="../../../images/imgSave.gIF"
													border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													alt="자료를 인쇄합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 120px" vAlign="top" align="center">
									<TABLE class="DATA" id="tblHidden" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" width="50">년월</TD>
											<TD class="DATA" width="200"><INPUT dataFld="YEARMON" class="INPUT" id="txtYEARMON" title="년월" style="WIDTH: 118px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON)"
													size="13" name="txtYEARMON"></TD>
											<TD class="LABEL" width="50">청구일</TD>
											<TD class="DATA" width="200"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="청구일" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="14" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(TBRDSTDATE, txtTBRDEDDATE)"
												width="50">계약기간</TD>
											<TD class="DATA" width="200" style="WIDTH: 225px"><INPUT dataFld="TBRDSTDATE" class="INPUT" id="txtTBRDSTDATE" title="계약기간" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtTBRDSTDATE">&nbsp;<IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar1">&nbsp;~<INPUT dataFld="TBRDEDDATE" class="INPUT" id="txtTBRDEDDATE" title="계약기간" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="8" name="txtTBRDEDDATE">&nbsp;<IMG id="imgCalEndar2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar2"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtAMT, '')"
												width="50">금액</TD>
											<TD class="DATA">
												<INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="금액" style="WIDTH: 136px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="17" name="txtAMT">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME, txtCLIENTCODE)">광고주</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="33" name="txtCLIENTNAME">&nbsp;<IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">&nbsp;<INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="10" size="4" name="txtCLIENTCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTITLE, '')">계약명</TD>
											<TD class="DATA"><INPUT dataFld="TITLE" class="INPUT" id="txtTITLE" title="계약명" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="18" name="txtTITLE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTOTALAMT, '')">총금액</TD>
											<TD class="DATA"><INPUT dataFld="TOTALAMT" class="INPUT_R" id="txtTOTALAMT" title="총금액" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="9" size="17" name="txtTOTALAMT"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtOUT_AMT, '')">외주비</TD>
											<TD class="DATA">
												<INPUT dataFld="OUT_AMT" class="INPUT_R" id="txtOUT_AMT" title="외주비" style="WIDTH: 136px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="50" size="17" name="txtOUT_AMT">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME, txtTIMCODE)">팀</TD>
											<TD class="DATA"><INPUT dataFld="TIMNAME" class="INPUT_L" id="txtTIMNAME" title="팀명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="11" name="txtTIMNAME">&nbsp;<IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgTIMCODE">&nbsp;<INPUT dataFld="txtTIMCODE" class="INPUT_L" id="txtTIMCODE" title="담당부서" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="10" size="3" name="txtTIMCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMATTERNAME, '')">소재명</TD>
											<TD class="DATA"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="소재명" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="500" size="30" name="txtMATTERNAME"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtLOCATION, '')">장소</TD>
											<TD class="DATA"><INPUT dataFld="LOCATION" class="INPUT_L" id="txtLOCATION" title="장소" style="WIDTH: 222px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtLOCATION">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMI_RATE, '')">내수율</TD>
											<TD class="DATA">
												<INPUT dataFld="COMMI_RATE" class="INPUT_R" id="txtCOMMI_RATE" title="내수율" style="WIDTH: 64px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="5" name="txtCOMMI_RATE">&nbsp;% 
												AOR <INPUT id="chkAORFLAG" title="AOR플래그" type="checkbox" name="chkAORFLAG">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)">매체사</TD>
											<TD class="DATA"><INPUT dataFld="REAL_MED_NAME" class="INPUT_L" id="txtREAL_MED_NAME" title="매체사명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="32" name="txtREAL_MED_NAME">&nbsp;<IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE">&nbsp;<INPUT dataFld="REAL_MED_CODE" class="INPUT_L" id="txtREAL_MED_CODE" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="10" size="4" name="txtREAL_MED_CODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMED_GBN, '')">제작종류</TD>
											<TD class="DATA"><INPUT dataFld="MED_GBN" class="INPUT_L" id="txtMED_GBN" title="제작종류" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="30" name="txtMED_GBN">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEMO, '')">비고</TD>
											<TD class="DATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="비고" style="WIDTH: 222px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="120" size="15" name="txtMEMO"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMISSION, '')">내수액</TD>
											<TD class="DATA">
												<INPUT dataFld="COMMISSION" class="INPUT_R" id="txtCOMMISSION" title="내수액" style="WIDTH: 136px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="17" name="txtCOMMISSION">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--BodySplit End-->
						</TABLE>
						<TABLE id="tblSheet" height="65%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td class="DATA" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31856">
										<PARAM NAME="_ExtentY" VALUE="13309">
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
										<PARAM NAME="MaxCols" VALUE="44">
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
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
