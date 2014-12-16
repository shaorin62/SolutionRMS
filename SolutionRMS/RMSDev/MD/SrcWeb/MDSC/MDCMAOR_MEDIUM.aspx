<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMAOR_MEDIUM.aspx.vb" Inherits="MD.MDCMAOR_MEDIUM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AOR 매체분할 등록/조회</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'시스템구분 : MD/부킹 화면(MDCMBOOKING)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMAOR_MEDIUM.aspx
'기      능 : AOR 대행 매출 금액 조회 저장
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2012.05.15 OH SE HOON
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT id="clientEventHandlersVBS" language="vbscript">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOAORMEDIUM, mobjMDCOGET
Dim mstrCheck
Dim mstrHIDDEN
Dim mcomecalender

CONST meTAB = 9
mcomecalender = FALSE
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
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	On error resume next
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData

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
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'매체사 팝업 버튼
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData

		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "B")
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
				Else
					Call REAL_MED_CODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' 팝업 버튼[입력용]
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
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")

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
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'매체사 팝업 버튼
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' 코드명 표시
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
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
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "B")
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,1))
					.txtREAL_MED_NAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'제작사/대행사 팝업 
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code값 저장
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'코드명 표시
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtEXCLIENTCODE
		End If
	end With
End Sub

Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)

			vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code값 저장
					.txtEXCLIENTNAME.value = trim(vntData(2,1))	'코드명 표시
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'담당부서 팝업 
Sub imgDEPT_CD_onclick
	Call DEPT_CD_POP()
End Sub

Sub DEPT_CD_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtDEPT_NAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtDEPT_CD.value = trim(vntRet(0,0))	'Code값 저장
			.txtDEPT_NAME.value = trim(vntRet(1,0))	'코드명 표시
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtDEPT_CD
		End If
	end With
End Sub

'담당부서 팝업
Sub txtDEPT_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPT_NAME.value))
			
			If not gDoErrorRtn ("GetCC") Then
				If mlngRowCnt = 1 Then
					.txtDEPT_CD.value = trim(vntData(0,1))
					.txtDEPT_NAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call DEPT_CD_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'매체 팝업 버튼
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value), _
							trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "")
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMEDCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtMEDNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtREAL_MED_CODE.value = trim(vntRet(3,0))       ' 코드명 표시
			.txtREAL_MED_NAME.value = trim(vntRet(4,0))       ' 코드명 표시
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
											trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "")
			
			If not gDoErrorRtn ("GetMEDGUBNCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtMEDNAME.value = trim(vntData(1,1))       ' 코드명 표시
					.txtREAL_MED_CODE.value = trim(vntData(3,1))
					.txtREAL_MED_NAME.value = trim(vntData(4,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call MEDCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 게재일 달력
'****************************************************************************************
Sub imgCalEndar_onclick
	'CalEndar를 화면에 표시
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar,"txtDEMANDDAY_onchange()"
	Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DEMANDDAY"), frmThis.sprSht.ActiveRow)
	mcomecalender = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'****************************************************************************************
' 입력필드 키다운 이벤트
'****************************************************************************************
Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCARD_AMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbMED_FLAG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtREAL_MED_CODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMISSION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_CARD.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEX_CARD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEMANDDAY_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEXCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEXCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtOUT_AMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtOUT_AMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_AMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 입력필드 체인지 이벤트
'****************************************************************************************
Sub txtYEARMON_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		
		AMT_CAL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCARD_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, frmThis.txtCARD_AMT.value
		EXCARD_CAL frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub cmbMED_FLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		IF frmThis.cmbMED_FLAG.value =  "A" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "공중파"
		ELSEIF frmThis.cmbMED_FLAG.value =  "A2" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "케이블"
		ELSEIF frmThis.cmbMED_FLAG.value =  "T" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "종합편성방송"
		ELSEIF frmThis.cmbMED_FLAG.value =  "B" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "신문"
		ELSEIF frmThis.cmbMED_FLAG.value =  "C" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "잡지"
		ELSEIF frmThis.cmbMED_FLAG.value =  "O" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "인터넷"
		ELSEIF frmThis.cmbMED_FLAG.value =  "D" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "옥외"			
		END IF
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCOMMISSION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
		EXCARD_CAL frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCOMMI_RATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, frmThis.txtCOMMI_RATE.value
		COMMISSION_CAL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtEX_CARD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, frmThis.txtEX_CARD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEX_CARD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, frmThis.txtEX_CARD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtDEMANDDAY_onchange
	Dim strdate 
	Dim strDEMANDDAY
	strdate = "" : strDEMANDDAY = ""
	
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

Sub txtOUT_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, frmThis.txtOUT_AMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtEX_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, frmThis.txtEX_AMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtMEMO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'공급가액
Sub txtAMT_onblur
	With frmThis
		Call gFormatNumber(.txtAMT,0,True)
	end With
End Sub
'공급가액 부가세
Sub txtVAT_onblur
	With frmThis
		Call gFormatNumber(.txtVAT,0,True)
	end With
End Sub

'부가세 포함
Sub txtSUMAMTVAT_onblur
	With frmThis
		Call gFormatNumber(.txtSUMAMTVAT,0,True)
	end With
End Sub

'수수료
Sub txtCOMMISSION_onblur
	With frmThis
		Call gFormatNumber(.txtCOMMISSION,0,True)
	end With
End Sub

'카드 수수료
Sub txtCARD_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtCARD_AMT,0,True)
	end With
End Sub

'카드 제외 금액
Sub txtEX_CARD_onblur
	With frmThis
		Call gFormatNumber(.txtEX_CARD,0,True)
	end With
End Sub

'매체사확정금액
Sub txtOUT_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtOUT_AMT,0,True)
	end With
End Sub

'제작대행사 확정 금액
Sub EX_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtEX_AMT,0,True)
	end With
End Sub

'-----------------------------------------------------------------------------------------   
' 천단위 나눔점 없애기 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'공급가액
Sub txtAMT_onfocus
	With frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end With
End Sub

'공급가액 부가세
Sub txtVAT_onfocus
	With frmThis
		.txtVAT.value = Replace(.txtVAT.value,",","")
	end With
End Sub

'부가세 포함
Sub txtSUMAMTVAT_onfocus
	With frmThis
		.txtSUMAMTVAT.value = Replace(.txtSUMAMTVAT.value,",","")
	end With
End Sub

'수수료
Sub txtCOMMISSION_onfocus
	With frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end With
End Sub

'카드 수수료
Sub txtCARD_AMT_onfocus
	With frmThis
		.txtCARD_AMT.value = Replace(.txtCARD_AMT.value,",","")
	end With
End Sub

'카드 제외 금액
Sub txtEX_CARD_onfocus
	With frmThis
		.txtEX_CARD.value = Replace(.txtEX_CARD.value,",","")
	end With
End Sub

'매체사확정금액
Sub txtOUT_AMT_onfocus
	With frmThis
		.txtOUT_AMT.value = Replace(.txtOUT_AMT.value,",","")
	end With
End Sub

'제작대행사 확정 금액
Sub txtEX_AMT_onfocus
	With frmThis
		.txtEX_AMT.value = Replace(.txtEX_AMT.value,",","")
	end With
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		
		frmThis.txtSELECTAMT.value = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON1.value
		
		IF frmThis.cmbMED_FLAG.value = "A" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "공중파"
		ELSEIF frmThis.cmbMED_FLAG.value = "A2" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "케이블"
		ELSEIF frmThis.cmbMED_FLAG.value = "T" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "종합편성방송"
		ELSEIF frmThis.cmbMED_FLAG.value = "B" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "신문"
		ELSEIF frmThis.cmbMED_FLAG.value = "C" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "잡지"
		ELSEIF frmThis.cmbMED_FLAG.value = "O" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "인터넷"
		ELSEIF frmThis.cmbMED_FLAG.value = "D" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "옥외"
		END IF 
		'AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | OUT_AMT | EX_AMT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEMANDDAY",frmThis.sprSht.ActiveRow, gNowDate2
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VAT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, 15
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, 0
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, "G00076"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, "에스케이 플래닛(주)"
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub sprSht_Keyup(KeyCode, Shift) 
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
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
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or _ 
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") or _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT") Then
			
			strSUM = 0 : intSelCnt = 0 : intSelCnt1 = 0
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT")) Then
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

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i,j
	Dim vntData_col, vntData_row
	
	With frmThis
		strSUM = 0 : intColCnt = 0 : intRowCnt = 0
		
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or _ 
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") or _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT") Then
			
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intColCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intRowCnt)
					
					for i = 0 to intColCnt -1
						if vntData_col(i) <> "" then
							FOR j = 0 TO intRowCnt -1
								If vntData_row(j) <> "" Then
									if typename(mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))) = "String" then
										exit sub
									end if 
									strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
								End If
							Next
						end if 
					next
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
   	Dim strCode, strCodeName
   	Dim intCnt
	With frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strCode = "" : strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON") Then .txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDDAY") Then .txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MED_FLAG") Then 
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "공중파" THEN
				.cmbMED_FLAG.value = "A"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "케이블" THEN
				.cmbMED_FLAG.value = "A2"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "종합편성방송" THEN
				.cmbMED_FLAG.value = "T"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "신문" THEN
				.cmbMED_FLAG.value = "B"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "잡지" THEN
				.cmbMED_FLAG.value = "C"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "인터넷" THEN
				.cmbMED_FLAG.value = "O"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "옥외" THEN
				.cmbMED_FLAG.value = "D"
			END IF 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then 
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			AMT_CAL Col,Row
		end if 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then 
			.txtCOMMI_RATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
			COMMISSION_CAL Col,Row	'대행수수료 계산
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then 
			.txtCOMMISSION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
			EXCARD_CAL  Col, Row	'카드수수료제외금액 계산
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") Then 
			.txtCARD_AMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
			EXCARD_CAL  Col, Row	'카드수수료제외금액 계산
		end if

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") Then .txtEX_CARD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then .txtOUT_AMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT")  Then .txtEX_AMT.value  = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_AMT",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE") Then	.txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
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
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						.txtCLIENTCODE.value = vntData(0,1)
						.txtCLIENTNAME.value = vntData(1,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
	
		'매체 변경시
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDCODE") Then .txtMEDCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", _
													strCode, strCodeName, "MED_PRINT")

				If not gDoErrorRtn ("GetMEDGUBNCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntData(4,1)
						.txtMEDCODE.value = vntData(0,1)
						.txtMEDNAME.value = vntData(1,1)
						.txtREAL_MED_CODE.value = vntData(3,1)
						.txtREAL_MED_NAME.value = vntData(4,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_CODE") Then .txtREAL_MED_CODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row) = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntData(1,1))
						.txtREAL_MED_CODE.value = trim(vntData(0,1))	    ' Code값 저장
						.txtREAL_MED_NAME.value = trim(vntData(1,1))       ' 코드명 표시

						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_CD") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.txtDEPT_CD.value = trim(vntData(0,1))
						.txtDEPT_NAME.value = trim(vntData(1,1))
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTCODE") Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code값 저장
						.txtEXCLIENTNAME.value = trim(vntData(2,1))	'코드명 표시
						
						.txtEXCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME"), Row
						.txtEXCLIENTNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO") Then .txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))

			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				.txtCLIENTCODE.value = vntRet(0,0)		
				.txtCLIENTNAME.value = vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then		
			vntInParams = array("","" , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)), "MED_PRINT")

			vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntRet(4,0)
				.txtMEDCODE.value = vntRet(0,0)
				.txtMEDNAME.value = vntRet(1,0)
				.txtREAL_MED_CODE.value = vntRet(3,0)
				.txtREAL_MED_NAME.value = vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then		
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntRet(1,0))
				.txtREAL_MED_CODE.value = vntRet(0,0)
				.txtREAL_MED_NAME.value = vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)

			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				
				.txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code값 저장
				.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'코드명 표시

				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(1,0)
				
				.txtDEPT_CD.value = trim(vntRet(0,0))	'Code값 저장
				.txtDEPT_NAME.value = trim(vntRet(1,0))	'코드명 표시
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If

		sprShtToFieldBinding Col, Row
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다.
		.txtCLIENTNAME1.focus()
		.sprSht.Focus
	End With
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			sprShtToFieldBinding Col,Row
		elseif Row = 0 and Col = 1 Then
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
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtCARD_AMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "공중파" THEN
			.cmbMED_FLAG.value	= "A"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "케이블" THEN
			.cmbMED_FLAG.value	= "A2"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "종합편성방송" THEN
			.cmbMED_FLAG.value	= "T"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "신문" THEN
			.cmbMED_FLAG.value	= "B"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "잡지" THEN
			.cmbMED_FLAG.value	= "C"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "인터넷" THEN
			.cmbMED_FLAG.value	= "O"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "옥외" THEN
			.cmbMED_FLAG.value	= "D"
		END IF 
		
		.txtREAL_MED_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtREAL_MED_CODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtMEDNAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		.txtMEDCODE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		
		.txtCOMMISSION.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		.txtCOMMI_RATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtEX_CARD.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		.txtDEMANDDAY.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		.txtEXCLIENTNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		.txtEXCLIENTCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		
		.txtDEPT_NAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtDEPT_CD.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		
		.txtOUT_AMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		.txtEX_AMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_AMT",Row)
		
		.txtMEMO.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
   	end With
	
	Call gFormatNumber(frmThis.txtAMT,0,True)
	Call gFormatNumber(frmThis.txtCARD_AMT,0,True)
	Call gFormatNumber(frmThis.txtCOMMISSION,0,True)
	Call gFormatNumber(frmThis.txtEX_CARD,0,True)
	Call gFormatNumber(frmThis.txtOUT_AMT,0,True)
	Call gFormatNumber(frmThis.txtEX_AMT,0,True)
	
	Call Field_Lock ()
End Function

'------------------------------------------------------------------------------------------
' Field_Lock  거래명세서번호나 세금계산서 번호가 있으면 수정할수 없도록 필드를 ReadOnly처리
'------------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			'거래명세서가 생성되면 필드를 잠근다.
			If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> "" Then
				.txtYEARMON.className		= "NOINPUT" : .txtYEARMON.readOnly		= True
				'광고주
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly			= True
				.txtCARD_AMT.className		= "NOINPUT_R" : .txtCARD_AMT.readOnly		= True
				.cmbMED_FLAG.disabled = True
				
				'매체사
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				
				'매체
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True

				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True
				.txtEX_CARD.className		= "NOINPUT_R" : .txtEX_CARD.readOnly		= True
				.txtDEMANDDAY.className		= "NOINPUT" : .txtDEMANDDAY.readOnly		= True : .imgCalEndar.disabled = True
				
				'제작대행사
				.txtEXCLIENTNAME.className = "NOINPUT_L" : .txtEXCLIENTNAME.readOnly	= True : .ImgEXCLIENTCODE.disabled = True
				.txtEXCLIENTCODE.className = "NOINPUT_L" : .txtEXCLIENTCODE.readOnly	= True
				
				'담당부서
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .ImgDEPT_CD.disabled = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				
				.txtOUT_AMT.className		= "NOINPUT_R" : .txtOUT_AMT.readOnly		= True
				.txtEX_AMT.className		= "NOINPUT_R" : .txtEX_AMT.readOnly			= True
				
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True

			else 
				.txtYEARMON.className		= "INPUT" : .txtYEARMON.readOnly			= False
				'광고주
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly		= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly		= False
				
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly				= False
				.txtCARD_AMT.className		= "INPUT_R" : .txtCARD_AMT.readOnly			= False
				.cmbMED_FLAG.disabled = False
				
				'매체사
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly	= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly	= False
				
				'매체
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly			= False : .ImgMEDCODE.disabled = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly			= False
				
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly		= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly		= False
				.txtEX_CARD.className		= "INPUT_R" : .txtEX_CARD.readOnly			= False
				.txtDEMANDDAY.className		= "INPUT" : .txtDEMANDDAY.readOnly		= False	: .imgCalEndar.disabled = false

				'제작대행사
				.txtEXCLIENTNAME.className = "INPUT_L" : .txtEXCLIENTNAME.readOnly		= False : .ImgEXCLIENTCODE.disabled = False
				.txtEXCLIENTCODE.className = "INPUT_L" : .txtEXCLIENTCODE.readOnly		= False
				
				'담당부서
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly		= False : .ImgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly			= False
				
				.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly			= False
				.txtEX_AMT.className		= "INPUT_R" : .txtEX_AMT.readOnly			= False
				
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly				= False
			End If
		else
			.txtYEARMON.className		= "INPUT" : .txtYEARMON.readOnly				= False
			'광고주
			.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly			= False : .ImgCLIENTCODE.disabled = False
			.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly			= False
			
			.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly					= False
			.txtCARD_AMT.className		= "INPUT_R" : .txtCARD_AMT.readOnly				= False
			.cmbMED_FLAG.disabled = False
			
			'매체사
			.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly		= False : .ImgREAL_MED_CODE.disabled = False
			.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly		= False
			
			'매체
			.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly				= False : .ImgMEDCODE.disabled = False
			.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly				= False
			
			.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly			= False
			.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly			= False
			.txtEX_CARD.className		= "INPUT_R" : .txtEX_CARD.readOnly				= False
			.txtDEMANDDAY.className		= "INPUT" : .txtDEMANDDAY.readOnly			= False : .imgCalEndar.disabled = false
							
			'제작대행사
			.txtEXCLIENTNAME.className = "INPUT_L" : .txtEXCLIENTNAME.readOnly			= False : .ImgEXCLIENTCODE.disabled = False
			.txtEXCLIENTCODE.className = "INPUT_L" : .txtEXCLIENTCODE.readOnly			= False
		
			'담당부서
			.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly			= False : .ImgDEPT_CD.disabled = False
			.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly				= False
							
			.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly				= False
			.txtEX_AMT.className		= "INPUT_R" : .txtEX_AMT.readOnly				= False
			
			.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly					= False

		End If
	End With
End Sub

'공급가액 계산
sub	AMT_CAL (ByVal Col, ByVal Row)
	Dim intAMT
	Dim intVAT
	Dim intSUMAMTVAT
	with frmThis

		intAMT = 0 : intVAT = 0 : intSUMAMTVAT = 0

		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intVAT = intAMT * 0.1
		intSUMAMTVAT = intAMT + intVAT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VAT",frmThis.sprSht.ActiveRow, intVAT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",frmThis.sprSht.ActiveRow, intSUMAMTVAT
		
		COMMISSION_CAL Col, Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'대행 수수료 계산
SUB COMMISSION_CAL (ByVal Col, ByVal Row)
	Dim intAMT
	Dim intCOMMI_RATE
	Dim intCOMMISSION
	with frmThis

		intAMT = 0 : intCOMMI_RATE = 0 : intCOMMISSION = 0
		
		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		intCOMMISSION = round(intAMT * (intCOMMI_RATE / 100),0)
		
		.txtCOMMISSION.value = intCOMMISSION
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, intCOMMISSION

		CARD_CAL Col, Row

	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
END SUB

'카드 수수료 계산(1.32%)
sub CARD_CAL (ByVal Col, ByVal Row)
	Dim intAMT			'공급가액
	Dim intCARD_AMT		'카드 수수료
	
	with frmThis
		intAMT = 0 : intCARD_AMT = 0 
		
		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intCARD_AMT = round(intAMT * 0.0132,0)
		
		.txtCARD_AMT.value = intCARD_AMT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, intCARD_AMT
		
		EXCARD_CAL Col,Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'카드수수료 제외 계산
sub EXCARD_CAL (ByVal Col, ByVal Row)
	Dim intCARD_AMT		'카드 수수료
	Dim intCOMMISSION	'대행수수료
	Dim intEX_CARD		'카드 수수료제외금액
	
	with frmthis
		intCARD_AMT = 0 : intCOMMISSION = 0 : intEX_CARD = 0

		intCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		intCARD_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
		
		intEX_CARD = intCOMMISSION - intCARD_AMT
		.txtEX_CARD.value = intEX_CARD
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, intEX_CARD
		
		EX_CAL Col,Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'매체 대행사 분할 계산
sub EX_CAL (ByVal Col, ByVal Row)
	Dim intEX_CARD		'카드 수수료제외금액
	Dim intOUT_AMT		'매체사 확정금액
	Dim intEX_AMT		'매체 대행사 확정금액
	
	with frmthis
		intEX_CARD = 0 : intOUT_AMT = 0 : intEX_AMT = 0

		
		intEX_CARD = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		
		intOUT_AMT = clng(intEX_CARD) * 0.3
		intEX_AMT = clng(intEX_CARD) * 0.7
		
		.txtOUT_AMT.value = intOUT_AMT
		.txtEX_AMT.value = intEX_AMT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, intOUT_AMT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, intEX_AMT
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub


'========================================================================================
' UI업무 프로시져 
'========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDCOAORMEDIUM	= gCreateRemoteObject("cMDSC.ccMDSCAORMEDIUM")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	mobjSCGLCtl.DoEventQueue
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 28, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | MED_FLAG | CLIENTCODE | CLIENTNAME| DEMANDDAY | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | EX_AMT | MEMO | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|년월|순번|매체구분|광고주코드|광고주명|청구일|공급가액|VAT|VAT포함금액|수수료율|수수료|카드수수료(1.32%)|카드수수료제외|매체코드|매체명|매체사코드|매체사명|매체사확정금액|제작대행사코드|제작대행사명|담당부서코드|담당부서명|제작대행사확정금액|비고|거래명세서번호|세금계산서번호|전표번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   8|   4|      10|         0|      14|    10|      14| 12|         14|       6|    14|				 14|            14|		  0|    10|     	0|      12|			   14|			   0|          14|           0|        12|                14|  15|             0|             0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "MED_FLAG", -1, -1, "공중파" & vbTab & "케이블" & vbTab & "종합편성방송" & vbTab & "신문" & vbTab & "잡지" & vbTab & "인터넷" & vbTab & "옥외" , 10, 90, False, False
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME| REAL_MED_CODE | REAL_MED_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT | VAT | SUMAMTVAT | COMMISSION | CARD_AMT | EX_CARD | OUT_AMT | EX_AMT ", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "SEQ | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE | MEDCODE | REAL_MED_CODE | EXCLIENTCODE | DEPT_CD | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | YEARMON | CLIENTCODE | CLIENTNAME| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | DEMANDDAY | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD ",-1,-1,2,2,False
		.sprSht.style.visibility = "visible"
    End With
	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCOAORMEDIUM = Nothing
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
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		
		.txtDEMANDDAY.value  = gNowDate2
		.cmbMED_FLAG.value = "A"
		
	End With
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim vntData2
	Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
   	Dim strRows
	Dim intCnt, intCnt2

	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		intCnt2 = 1

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strYEARMON = "" : strCLIENTCODE = "" : strCLIENTNAME = "" :	strREAL_MED_CODE = "" :	strREAL_MED_NAME = "" 

		strYEARMON		 = .txtYEARMON1.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCLIENTNAME	 = .txtCLIENTNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value

		vntData = mobjMDCOAORMEDIUM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, _
											  strCLIENTCODE, strCLIENTNAME, _
											  strREAL_MED_CODE, strREAL_MED_NAME)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)

   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	   			For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> "" Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next

				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,1,27,True
   				'검색시에 첫행을 MASTER와 바인딩 시키기 위함
   				sprShtToFieldBinding 2, 1
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				InitPageData
   			End If
   			Field_Lock
   		End If
   	end With
End Sub

Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strDataCHK
	Dim lngCol, lngRow

	With frmThis
   		'데이터 Validation
		'If DataValidation =False Then exit Sub
		'On error resume Next
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "YEARMON | MED_FLAG | MEDCODE | REAL_MED_CODE | EXCLIENTCODE | DEPT_CD",lngCol, lngRow, False) 
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 년월/매체구분/매체/매체사/제작대행사/담당부서 (은)는 필수 입력 사항 입니다..","저장안내"
			Exit Sub
		End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | MED_FLAG | CLIENTCODE | CLIENTNAME| DEMANDDAY | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | EX_AMT | MEMO | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO")
		intRtn = mobjMDCOAORMEDIUM.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox intRtn &" 건의 자료가 저장되었습니다.","저장안내!"
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
	Dim lngchkCnt

	lngchkCnt = 0
	With frmThis
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
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
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)

				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDCOAORMEDIUM.DeleteRtn(gstrConfigXml,strYEARMON,dblSEQ)
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
				End If
   				intCnt = intCnt + 1
   			End If
		Next

		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox intCnt & "건의 자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	End With
	err.clear
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">AOR 대행매출</td>
										</tr>
									</table>
								</TD>
								<TD height="20" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="처리중입니다."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table Start-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblKey" class="SEARCHDATA" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
									width="50">년월</TD>
								<TD class="SEARCHDATA" width="100"><INPUT accessKey="NUM" style="WIDTH: 96px; HEIGHT: 22px" id="txtYEARMON1" class="INPUT"
										title="년월조회" maxLength="6" size="10" name="txtYEARMON1"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">광고주</TD>
								<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 173px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="코드명"
										maxLength="100" align="left" size="22" name="txtCLIENTNAME1"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
									<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT_L" title="코드조회"
										maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD style="WIDTH: 45px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)">매체사</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 173px; HEIGHT: 22px" id="txtREAL_MED_NAME1" class="INPUT_L" title="매체사명"
										maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG style="CURSOR: hand" id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgREAL_MED_CODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
									<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtREAL_MED_CODE1" class="INPUT_L" title="매체사코드"
										maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" align="right"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF"
										height="20"></TD>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 20px" class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="500" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
										<tr>
											<td class="TITLE" vAlign="middle"><span style="CURSOR: hand" id="spnHIDDEN" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG style="CURSOR: hand" id="imgTableUp" border="0" name="imgTableUp" alt="자료를 검색합니다."
														align="absMiddle" src="../../../images/imgTableUp.gif"></span> &nbsp;&nbsp;&nbsp;&nbsp;선택 
												합계 : <INPUT style="WIDTH: 120px; HEIGHT: 22px" id="txtSELECTAMT" class="NOINPUTB_R" title="선택금액"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="top" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
										<TR>
											<TD><IMG style="CURSOR: hand" id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgCho"
													alt="자료를 초기화." src="../../../images/imgCho.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" border="0" name="imgREG"
													alt="신규자료를 생성합니다.." src="../../../images/imgNew.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" border="0" name="imgSave"
													alt="자료를 저장합니다." src="../../../images/imgSave.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
													alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 120px" vAlign="top" align="center">
									<TABLE id="tblHidden" class="DATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD class="LABEL" width="70">년월</TD>
											<TD style="WIDTH: 150px" class="DATA"><INPUT accessKey="NUM" style="WIDTH: 118px; HEIGHT: 22px" id="txtYEARMON" dataSrc="#xmlBind"
													class="INPUT" title="년월" dataFld="YEARMON" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" maxLength="6" size="13"
													name="txtYEARMON"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="60">광고주</TD>
											<TD style="WIDTH: 200px" class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="광고주명" dataFld="CLIENTNAME" maxLength="100" size="33" name="txtCLIENTNAME">&nbsp;<IMG style="CURSOR: hand" id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">&nbsp;<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="광고주코드" dataFld="CLIENTCODE" maxLength="10" size="4" name="txtCLIENTCODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtAMT, '')"
												width="70">공급가액</TD>
											<TD style="WIDTH: 200px" class="DATA"><INPUT accessKey="NUM" style="WIDTH: 196px; HEIGHT: 22px" id="txtAMT" dataSrc="#xmlBind"
													class="INPUT_R" title="공급가액" dataFld="AMT" maxLength="13" size="17" name="txtAMT">
											</TD>
											<TD class="LABEL" width="70">카드수수료</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtCARD_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="카드수수료금액" dataFld="CARD_AMT" maxLength="13" size="17" name="txtCARD_AMT">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL">매체구분</TD>
											<TD style="WIDTH: 148px" class="DATA"><SELECT style="WIDTH: 112px" id="cmbMED_FLAG" dataSrc="#xmlBind" title="매체구분" dataFld="MED_FLAG"
													name="cmbMED_FLAG">
													<OPTION selected value="A">공중파</OPTION>
													<OPTION value="A2">케이블</OPTION>
													<OPTION value="T">종합편성방송</OPTION>
													<OPTION value="B">신문</OPTION>
													<OPTION value="C">잡지</OPTION>
													<OPTION value="O">인터넷</OPTION>
													<OPTION value="D">옥외</OPTION>
												</SELECT>
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)">매체사</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtREAL_MED_NAME" dataSrc="#xmlBind" class="INPUT_L"
													title="매체사명" dataFld="REAL_MED_NAME" maxLength="100" size="32" name="txtREAL_MED_NAME">&nbsp;<IMG style="CURSOR: hand" id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgREAL_MED_CODE" align="absMiddle" src="../../../images/imgPopup.gIF">&nbsp;<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtREAL_MED_CODE" dataSrc="#xmlBind" class="INPUT_L"
													title="매체사코드" dataFld="REAL_MED_CODE" maxLength="10" size="4" name="txtREAL_MED_CODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtCOMMISSION, '')">수수료</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 123px; HEIGHT: 22px" id="txtCOMMISSION" dataSrc="#xmlBind"
													class="INPUT_R" title="수수료" dataFld="COMMISSION" maxLength="13" size="17" name="txtCOMMISSION">
												<INPUT style="WIDTH: 60px; HEIGHT: 22px" id="txtCOMMI_RATE" dataSrc="#xmlBind" class="INPUT_R"
													title="수수료율" dataFld="COMMI_RATE" maxLength="6" size="5" name="txtCOMMI_RATE">%
											</TD>
											<TD class="LABEL">카드제외</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtEX_CARD" dataSrc="#xmlBind"
													class="INPUT_R" title="카드수수료제외금액" dataFld="EX_CARD" maxLength="13" size="17" name="txtEX_CARD">
											</TD>
										</TR>
										<tr>
											<TD class="LABEL">청구일</TD>
											<TD class="DATA"><INPUT accessKey="DATE,M" style="WIDTH: 120px; HEIGHT: 22px" id="txtDEMANDDAY" dataSrc="#xmlBind"
													class="INPUT" title="청구일" dataFld="DEMANDDAY" maxLength="10" size="14" name="txtDEMANDDAY">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16">
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtMEDNAME, txtMEDCODE)">매체명</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtMEDNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="매체명" dataFld="MEDNAME" maxLength="100" size="13" name="txtMEDNAME"> <IMG style="CURSOR: hand" id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgMEDCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtMEDCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="매체명코드" dataFld="MEDCODE" maxLength="6" size="2" name="txtMEDCODE"></TD>
											<TD class="LABEL">매체확정금</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 196px; HEIGHT: 22px" id="txtOUT_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="매체사확정금액" dataFld="OUT_AMT" maxLength="13" size="17" name="txtOUT_AMT">
											</TD>
											<TD class="LABEL">대행확정금</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtEX_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="제작대행확정금액" dataFld="EX_AMT" maxLength="13" size="17" name="txtEX_AMT">
											</TD>
										</tr>
										<tr>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtDEPT_NAME, txtDEPT_CD)">담당부서</TD>
											<TD class="DATA"><INPUT style="WIDTH: 75px; HEIGHT: 22px" id="txtDEPT_NAME" dataSrc="#xmlBind" class="INPUT_L"
													title="담당부서명" dataFld="DEPT_NAME" maxLength="100" size="6" name="txtDEPT_NAME">
												<IMG style="CURSOR: hand" id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="imgDEPT_CD"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtDEPT_CD" dataSrc="#xmlBind"
													class="INPUT_L" title="담당부서코드" dataFld="DEPT_CD" maxLength="6" size="3" name="txtDEPT_CD"></TD>
											<TD style="CURSOR: hand; HEIGHT: 22px" class="LABEL" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)">제작대행</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtEXCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="제작사명" dataFld="EXCLIENTNAME" maxLength="100" size="30" name="txtEXCLIENTNAME">
												<IMG style="CURSOR: hand" id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgEXCLIENTCODE"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtEXCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="제작사코드" dataFld="EXCLIENTCODE" maxLength="10" size="4" name="txtEXCLIENTCODE"></TD>
											<TD class="LABEL">메모</TD>
											<TD class="DATA" colSpan="4"><INPUT style="WIDTH: 397px; HEIGHT: 22px" id="txtMEMO" dataSrc="#xmlBind" class="INPUT_R"
													title="비고" dataFld="MEMO" maxLength="255" size="17" name="txtMEMO"></TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="BODYSPLIT"></TD>
							</TR>
							<!--BodySplit End--></TABLE>
						<TABLE id="tblSheet" border="0" cellSpacing="0" cellPadding="0" width="100%" height="65%">
							<TR>
								<td style="WIDTH: 100%; HEIGHT: 100%" class="DATA" vAlign="top" align="center">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31882">
										<PARAM NAME="_ExtentY" VALUE="13520">
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
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
