<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMTOTALLIST.aspx.vb" Inherits="MD.MDCMTOTALLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>개별청약 승인/조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : MD/CATVLIST 청약승인화면(MDCMCATV)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : CATVLIST.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
'			 3) 2009/09/11 By Hwang Duck su
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
Dim mobjMDCMTOTAL
Dim mobjMDCMGET, mobjMDCMBOOKLIST 
Dim mstrCheck

CONST meTAB = 9

mstrCheck = True


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
	if frmThis.txtYEARMON1.value = "" then
		gErrorMsgBox "청구년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSetting_onclick
	Call ProcessRtn_ConfirmOK()
End Sub

Sub ImgConfirmCancel_onclick
	ProcessRtn_ConfirmCancel
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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
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
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
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
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
					SelectRtn
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
			vntData = mobjMDCMGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
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

'매체 팝업 버튼
Sub ImgMEDCODE1_onclick
	Call MEDCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub MEDCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value), _
							trim(.txtMEDCODE1.value), trim(.txtMEDNAME1.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE1.value = vntRet(0,0) and .txtMEDNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMEDCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtMEDNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtREAL_MED_CODE1.value = trim(vntRet(3,0))       ' 코드명 표시
			.txtREAL_MED_NAME1.value = trim(vntRet(4,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetMEDCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), _
											trim(.txtMEDCODE1.value),trim(.txtMEDNAME1.value))
			
			If not gDoErrorRtn ("GetMEDCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE1.value = trim(vntData(0,1))	    ' Code값 저장
					.txtMEDNAME1.value = trim(vntData(1,1))       ' 코드명 표시
					.txtREAL_MED_CODE1.value = trim(vntData(3,1))
					.txtREAL_MED_NAME1.value = trim(vntData(4,1))
					SelectRtn
				Else
					Call MEDCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'브랜드
Sub ImgSUBSEQ1_onclick
	Call SUBSEQCODE1_POP()
End Sub

Sub SUBSEQCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ1.value), trim(.txtSUBSEQNAME1.value), trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ1.value = vntRet(0,0) and .txtSUBSEQNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtSUBSEQ1.value = trim(vntRet(0,0))		' 브랜드 표시
			.txtSUBSEQNAME1.value = trim(vntRet(1,0))	' 브랜드명 표시
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' 광고주 표시
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' 광고주명 표시
			.txtTIMCODE1.value = trim(vntRet(4,0))	' 광고주명 표시
			.txtTIMNAME1.value = trim(vntRet(5,0))	' 광고주명 표시
			SelectRtn
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ1.value),trim(.txtSUBSEQNAME1.value),  _
												trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ1.value = trim(vntData(0,1))
					.txtSUBSEQNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(2,1))		' 광고주 표시
					.txtCLIENTNAME1.value = trim(vntData(3,1))	' 광고주
					.txtTIMCODE1.value = trim(vntData(4,1))	' 광고주
					.txtTIMNAME1.value = trim(vntData(5,1))	' 광고주
					SelectRtn
				Else
					Call SUBSEQCODE1_POP()
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
Sub imgCalFrom_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtFPUB_DATE,frmThis.imgCalFrom,"txtFPUB_DATE_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalTo_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtTPUB_DATE,frmThis.imgCalTo,"txtTPUB_DATE_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtFPUB_DATE_onchange
	gSetChange
End Sub

Sub txtTPUB_DATE_onchange
	gSetChange
End Sub

Sub txtYEARMON1_onkeydown
	'or window.event.keyCode = meTAB 탭일때는 아님 엔터일때만 조회
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = 1  then 
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
			
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"TRU_TRANS_NO",intCnt) <> "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If			
			Next
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


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
		'sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strCOLUMN = "COMMISSION"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
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
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDCMTOTAL		= gCreateRemoteObject("cMDCT.ccMDCTTOTAL_MEDIUM")
	set mobjMDCMBOOKLIST	= gCreateRemoteObject("cMDCO.ccMDCOBOOKINGLIST")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK |GFLAG| GFLAGNAME | CONFIRMFLAG | YEARMON | SEQ | DEMANDDAY | MEDNAME | CLIENTNAME | TIMNAME | MATTERNAME | AMT | COMMISSION | COMMI_RATE| TBRDSTDATE | TBRDEDDATE | CNT | VOCH_TYPE | TRU_TAX_FLAG | BIGO | TRU_TRANS_NO | COMMI_TRANS_NO"

		mobjSCGLSpr.SetHeader .sprSht,		 "선택|승인|승인여부|명세서여부|년월|순번|청구일|매체명|광고주|CIC/팀|소재|금액|수수료|수수료율|시작일|종료일|초수|전표구분|VAT|비고"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|       7|         0|   0|   0|     9|    13|    13|    13|  14|  10|     9|       9|     9|     9|   6|       8|  4|  10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK|TRU_TAX_FLAG"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "DEMANDDAY | MEDNAME | CLIENTNAME | MATTERNAME |BIGO", -1, -1, 50
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|COMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE ", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "GFLAG| GFLAGNAME | CONFIRMFLAG | YEARMON | SEQ | DEMANDDAY | MEDNAME | CLIENTNAME | TIMNAME | MATTERNAME | AMT | COMMISSION | COMMI_RATE| TBRDSTDATE | TBRDEDDATE | CNT | VOCH_TYPE | TRU_TAX_FLAG | BIGO | TRU_TRANS_NO | COMMI_TRANS_NO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GFLAGNAME|DEMANDDAY|CNT",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME | CLIENTNAME | MATTERNAME | TIMNAME | BIGO ",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ |  GFLAG | TRU_TRANS_NO | COMMI_TRANS_NO", true
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMTOTAL = Nothing
	set mobjMDCMBOOKLIST = Nothing
	set mobjMDCMGET = Nothing
	
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
		.txtYEARMON1.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet초기화
		
		.sprSht.MaxRows = 0
		.txtCLIENTNAME1.focus()
		Get_COMBO_VALUE
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


'-----------------------------------------------------------------------------------------
' 그리드 콤보박스 설정
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData, vntData_VOCH, vntData_DUTY
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_VOCH = mobjMDCMTOTAL.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE",,,vntData_VOCH,,60 
			mobjSCGLSpr.TypeComboBox = True 
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
	Dim strTIMCODE, strTIMNAME,strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME
   	Dim strGFLAG
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

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
		strMEDCODE		 = .txtMEDCODE1.value
		strMEDNAME		 = .txtMEDNAME1.value
		strSUBSEQ		 = .txtSUBSEQ1.value
		strSUBSEQNAME	 = .txtSUBSEQNAME1.value
		
		
		vntData = mobjMDCMBOOKLIST.SelectRtn_TOTAL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												   strYEARMON, _
												   strCLIENTCODE, strCLIENTNAME, _
												   strREAL_MED_CODE, strREAL_MED_NAME, _
												   strTIMCODE, strTIMNAME, _
												   strMEDCODE,strMEDNAME, _
												   strSUBSEQ,strSUBSEQNAME)

		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"GFLAG",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding( .sprSht,"TRU_TRANS_NO",intCnt) <> "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					
				Else
					'체크
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
				End If			
			Next
			
			AMT_SUM
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE	
   		end if
   	end with
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

'------------------------------------------
' 승인 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG 
	
	strFLAG = "CONFIRM"
	
	with frmThis
   		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				IF mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",intCnt) = "1" THEN
   					gErrorMsgBox "승인된 내역입니다.","저장안내!"
   					SELECTRTN
   					EXIT SUB
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|SEQ|CONFIRMFLAG")
		strFLAG = "CONFIRM"
		intRtn = mobjMDCMBOOKLIST.ProcessRtn_ConfirmTotal_OK(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmTotal_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 승인 취소로직
'------------------------------------------
Sub ProcessRtn_ConfirmCancel
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG
	strFLAG = "CANCEL"
	with frmThis
   		'데이터 Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
			 IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",intCnt) = "0" THEN
   					gErrorMsgBox "미승인된 내역입니다.","저장안내!"
   					SELECTRTN
   					EXIT SUB
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		If lngCHKSUM = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|SEQ|CONFIRMFLAG")
		strFLAG = "CANCEL"
		intRtn = mobjMDCMBOOKLIST.ProcessRtn_ConfirmTotal_OK(gstrConfigXml,vntData,strFLAG)
	
		if not gDoErrorRtn ("ProcessRtn_ConfirmTotal_OK") then 
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인취소" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
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
													<TABLE cellSpacing="0" cellPadding="0" width="198" background="../../../images/back_p.gIF"
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
												<td class="TITLE">청약관리 - 개별청약 조회 및 승인</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End--></TD>
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
							<TABLE height="95%" id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
													width="50">청구년월</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="10" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="50">광고주</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 123px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
													<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
													width="50">팀</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtTIMNAME1" title="팀명" style="WIDTH: 123px; HEIGHT: 22px" type="text"
														maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTIMCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1, txtSUBSEQ1)"
													width="50">브랜드</TD>
												<td class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="브랜드명" style="WIDTH: 136px; HEIGHT: 22px"
														type="text" maxLength="100" size="17" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgSUBSEQ1"> <INPUT class="INPUT_L" id="txtSUBSEQ1" title="시퀀스코드" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="8" name="txtSUBSEQ1">
												</td>
											</TR>
											<TR>
												<TD class="SEARCHDATA" width="250" colspan="2"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
													width="50">매체사</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="매체사명" style="WIDTH: 123px; HEIGHT: 22px"
														type="text" maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" name="txtREAL_MED_CODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME1, txtMEDCODE1)"
													width="50">매체명</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtMEDNAME1" title="매체명" style="WIDTH: 123px; HEIGHT: 22px"
														type="text" maxLength="100" size="15" name="txtMEDNAME1"> <IMG id="ImgMEDCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgMEDCODE1"> <INPUT class="INPUT_L" id="txtMEDCODE1" title="매체명코드" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" size="2" name="txtMEDCODE1"></TD>
												<td class="SEARCHDATA" colspan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
														align="right" border="0" name="imgQuery">&nbsp;
												</td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD align="left" width="400" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
																<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
																	height="20" alt="자료를승인처리합니다." src="../../../images/imgAgree.gIF" width="54" border="0"
																	name="imgSetting"></TD>
															<td><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCancelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCancel.gIF'"
																	height="20" alt="승인처리를 취소합니다." src="../../../images/imgAgreeCancel.gif" width="71"
																	border="0" name="ImgConfirmCancel"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																	name="imgExcel"></td>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="12488">
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
							</TABLE>
						</TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
