<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMAORCOMMIAL.aspx.vb" Inherits="MD.MDCMAORCOMMIAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AOR 거래명세서</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 위수탁거래명세서 등록 화면(MDCMPRINTTRANS1.aspx)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 위수탁거래명세서 입력/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/28 By Kim Tae Yub
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" >
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDSCAORCOMMI, mobjMDCOGET, mobjSCCOGET
Dim mstrCheck, mstrCheck1
Dim mstrGrid
CONST meTAB = 9

mstrCheck=True
mstrCheck1=True
mstrGrid = False
'=========================================================================================
' 이벤트 프로시져
'=========================================================================================
'입력 필드 숨기기
Sub Set_TBL_HIDDEN(byVal strmode)
	With frmThis
		If  strmode = "EXTENTION"  Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "60%"
			document.getElementById("tblSheet2").style.height = "30%"
		ElseIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ElseIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "60%"
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
Sub imgQuery_onclick
	If frmThis.txtYEARMON1.value = "" and frmThis.txtCLIENTCODE1.value = "" Then
		gErrorMsgBox "조회조건을 입력하시오.","조회안내"
		Exit Sub
	End If
	mstrGrid = False
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
Sub ImgCRE_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "상세항목 이 없습니다.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_CUST
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = True
		mobjSCGLSpr.ExcelExportOption = True
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = True
		mobjSCGLSpr.ExcelExportOption = True
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
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
	Dim strCNT
	Dim vntData
	Dim intRtn
	Dim strUSERID
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	If frmThis.sprSht_HDR.MaxRows = 0 Then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	With frmThis
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDSCAORCOMMI.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMAORCOMMI.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjMDSCAORCOMMI.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "call printSetTimeout('" & strTRANSYEARMON & "', '" & strTRANSNO & "')", 10000
	End With
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout(strTRANSYEARMON, strTRANSNO)
	Dim intRtn, intRtn2
	With frmThis
		intRtn = mobjMDSCAORCOMMI.DeleteRtn_temp(gstrConfigXml)
	End With
End Sub


'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON1.value,1,4) & "-" & MID(frmThis.txtYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

'승인 요청 버튼 클릭
Sub ImgConfirmRequest_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Confirm_User
	gFlowWait meWAIT_OFF
End Sub

Sub ProcessRtn_Confirm_User ()
   	Dim vntData
   	Dim intRtn
	Dim strTRANSYEARMON
	Dim strCLIENTCODE, strCLIENTNAME
	
	Dim intCnt,intCnt2,intCnt3, chkcnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg
	
	'SMS 정보
	Dim vntData_info
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		IF .txtEMPNO.value = "" THEN
			gErrorMsgBox "승인요청자를 입력 하십시오",""
			exit sub
		END IF 
		
		chkcnt=0
		For intCnt = 1 To .sprSht_HDR.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "승인요청할 데이터를 체크 하십시오",""
			exit sub
		end if
		
		For intCnt2 = 1 To .sprSht_HDR.MaxRows
			'그리드의 제작건명 을 가져온다
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt2) = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REAL_MED_NAME",intCnt2)
				 Exit For
			End If
		Next
	
		If chkcnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] 승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 승인요청이있습니다"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] 외" & chkcnt-1 & "건의승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 외" & chkcnt-1 & "건의승인요청이있습니다"
			End If
		End If
		
		intSaveRtn = gYesNoMsgbox("해당데이터를 승인요청 하시겠습니까?","승인요청 확인")
		IF intSaveRtn = vbYes then 
		
			'승인을 수락하였으므로 SMS 발송
			'보내는 사람의 정보 가져오기
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			'vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
			
			'보내는사람정보
			'strFromUserName	= vntData_info(0,2)
			'strFromUserEmail	= vntData_info(1,2)
			'strFromUserPhone	= vntData_info(2,2)

			'받는사람 정보
			'strToUserName		=  vntData_info(0,1)
			'strToUserEmail		=  vntData_info(1,1)
			'strToUserPhone		=  vntData_info(2,1)

			'call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)

			'저장플레그 설정
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meINS_TRANS
			gXMLSetFlag xmlBind, meINS_TRANS

			'쉬트의 변경된 데이터만 가져온다.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO")
			
			strTRANSYEARMON = .txtYEARMON1.value
			strCLIENTCODE	= .txtCLIENTCODE1.value
			strCLIENTNAME	= .txtCLIENTNAME1.value
			
			intRtn = mobjMDSCAORCOMMI.ProcessRtn_Confirm_User(gstrConfigXml,vntData,.txtEMPNO.value)
			
   			if not gDoErrorRtn ("ProcessRtn_Confirm_USER") then
				'모든 플래그 클리어
				mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
				initpagedata
				gOkMsgBox "승인요청이 되었습니다.","확인"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					selectRtn
				Else
					initpagedata
				End If
				DateClean
   			end if
   		End If
   	End With
End Sub

'-----------------------------------------------------------------------------------------
' 사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'실제 데이터List 가져오기
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< 받아오는경우

		vntRet = gShowModalWindow("../MDCO/MDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()				' 포커스 이동
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			
			if not gDoErrorRtn ("GetMDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP ()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	With frmThis
		vntInParams = array(.txtYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "commi", "AOR") 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,445)

		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then Exit Sub ' 변경된 데이터가 없다면 Exit
			If vntRet(3,0) = "완료" Then
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(4,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			Else
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			End If
			selectRtn
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag 변경 알림
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols

		On error resume next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"","commi", "AOR")
			If not gDoErrorRtn ("GetTRANSCUSTNO") Then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					selectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub	

'-----------------------------------------------------------------------------------------
' 청구일자 세팅
'-----------------------------------------------------------------------------------------
Sub txtYEARMON1_onblur
	With frmThis
		If .txtYEARMON1.value <> "" AND Len(.txtYEARMON1.value) = 6 Then DateClean
	End With
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

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	With frmThis
		If Row = 0 and Col = 1 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			Elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 To .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		Elseif Row > 0 AND Col > 1 Then
			mstrGrid = True
			CALL Grid_Setting ()
			SelectRtn_DTL Col, Row
			'mstrGrid = False
		End If
	End With
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	With frmThis
		If mstrGrid = False Then
			If Row = 0 and Col = 1 Then
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_DTL, 1, 1, , , "", , , , , mstrCheck1
				If mstrCheck1 = True Then 
					mstrCheck1 = False
				Elseif mstrCheck1 = False Then 
					mstrCheck1 = True
				End If
				
				for intcnt = 1 To .sprSht_DTL.MaxRows
					sprSht_DTL_Change 1, intcnt
				next
			End If
		End If
	End With
End Sub  

Sub sprSht_HDR_Keyup(KeyCode, Shift)
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
		mstrGrid = True
		CALL Grid_Setting ()
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
			.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)

			FOR i = 0 To intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT")) Then
					FOR j = 0 To intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
			
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		Else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	
	With frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		If .sprSht_HDR.MaxRows >0 Then
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
				.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
					
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intColCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intRowCnt)
				
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			End If
		Else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		End If
	End With
End Sub

Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub

Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub

Sub sprSht_DTL_Keyup(KeyCode, Shift)
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

	With frmThis
		If mstrGrid Then
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
				.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 To intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT")) Then
						FOR j = 0 To intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			Else
				.txtSELECTAMT.value = 0
			End If
		Else
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
			   .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0

				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 To intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT")) Then
						FOR j = 0 To intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			Else
				.txtSELECTAMT.value = 0
			End If
		End If
	End With
End Sub

Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	
	With frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		If mstrGrid Then
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") OR  .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT") Then
						
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intColCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intRowCnt)
					for i = 0 to intColCnt -1
						if vntData_col(i) <> "" then
							FOR j = 0 TO intRowCnt -1
								If vntData_row(j) <> "" Then
									if typename(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))) = "String" then
										exit sub
									end if 
									strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
								End If
							Next
						end if 
					next
					.txtSELECTAMT.value = strSUM
					Call gFormatNumber(.txtSELECTAMT,0,True)
				End If
			Else
				.txtSELECTAMT.value = 0
			End If
		ELSE
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"OUT_AMT") OR _ 
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT")  Then
						
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intColCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intRowCnt)
					for i = 0 to intColCnt -1
						if vntData_col(i) <> "" then
							FOR j = 0 TO intRowCnt -1
								If vntData_row(j) <> "" Then
									if typename(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))) = "String" then
										exit sub
									end if 
									strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
									
								End If
							Next
						end if 
					next
					.txtSELECTAMT.value = strSUM
					Call gFormatNumber(.txtSELECTAMT,0,True)
				End If
			Else
				.txtSELECTAMT.value = 0
			End If
		END IF
	End With
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	With frmThis
	
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row  
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
	set mobjMDSCAORCOMMI= gCreateRemoteObject("cMDSC.ccMDSCAORCOMMI")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue

	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    With frmThis
		'거래명세서 헤더 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 17, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | REQUESTGBN | CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | TRANSNO | REAL_MED_NAME | DEMANDDAY | PRINTDAY | AMT | VAT | SUMAMTVAT | CONFIRM_USER | CONFIRM_DATE | MED_FLAGNAME | MEMO | CNT"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "선택|요청|승인|계산서|거래년월|번호|청구지|청구일|발행일|공급가액|부가세|계|승인자|승인일|매체구분|비고|상세행수"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "  4|   4|   4|     6|       8|   5|    15|     8|     8|      10|    10|10|    10|     9|      12|  15|      7"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY | PRINTDAY | CONFIRM_DATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TRANSNO | AMT | VAT | SUMAMTVAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "REQUESTGBN | CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | REAL_MED_NAME | CONFIRM_USER | CONFIRM_DATE | MED_FLAGNAME | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, True, "REQUESTGBN | CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | TRANSNO | REAL_MED_NAME | DEMANDDAY | PRINTDAY | AMT | VAT | SUMAMTVAT | CONFIRM_USER | MED_FLAGNAME | MEMO | CNT"
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "REQUESTGBN | CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | MED_FLAGNAME | CONFIRM_USER" ,-1,-1,2,2,False

		.sprSht_HDR.style.visibility = "visible"
    End With	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSCAORCOMMI = Nothing
	set mobjMDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

Sub Grid_Setting ()
	With frmThis
		mobjSCGLCtl.DoEventQueue
		If mstrGrid Then
			'Sheet 기본Color 지정
			gSetSheetDefaultColor() 
			'******************************************************************
			''거래명세서 디테일
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 27, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | TITLE | MED_FLAGNAME | AMT | COMMI_RATE | COMMISSION | VAT | OUT_AMT | MEMO | DEPT_CD | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON "
			mobjSCGLSpr.SetHeader .sprSht_DTL,		"거래명세서년월|거래명세서번호|순번|발행순번|신탁순번|광고주코드|광고주명|매체사코드|매체사명|매체코드|매체명|계약명|매체구분명|취급액|수수료율|수수료|부가세|매체사확정금액|비고|담당부서코드|담당부서명|청구일|발행일|세금계산서년월|세금계산서번호|승인유무|신탁년월" 
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "          0|	         0|	  0|       0|       4|	       0|      12|         0|      12|       0|    10|    10|         8|    10|       6|    10|     8|             0|  10|           0|        10|     8|     8|            10|            10|       0|       0"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CONFIRMFLAG"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY | PRINTDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | AMT | COMMI_RATE | COMMISSION | VAT | OUT_AMT ", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "TRANSYEARMON | CLIENTNAME | REAL_MED_NAME | MEDNAME | DEPT_NAME | TITLE | MED_FLAGNAME | MEMO | TAXYEARMON | TAXNO | TRUST_YEARMON", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, True, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | TITLE | MED_FLAGNAME | AMT | COMMI_RATE | COMMISSION | VAT | OUT_AMT | MEMO | DEPT_CD | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON " 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | SEQ | TRUST_SEQ | CLIENTNAME | REAL_MED_NAME | MEDNAME | TITLE | MED_FLAGNAME | AMT | COMMI_RATE | VAT | OUT_AMT | DEPT_NAME | DEMANDDAY | PRINTDAY | MEMO | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON", False
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSNO | PRINT_SEQ", True
		Else
			'Sheet 기본Color 지정
			gSetSheetDefaultColor() 
			'******************************************************************
			'청약내역 그리드
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 23, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_DTL,   "CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | DEPT_CD | DEPT_NAME | TITLE | DEMANDDAY | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | OUT_AMT | MEMO | TRANSCUSTRANK"
			mobjSCGLSpr.SetHeader .sprSht_DTL,		   "선택|년월|순번|매체구분|매체구분명|광고주코드|광고주명|매체사코드|매체사명|매체코드|매체명|담당부서코드|담당부서명|계약명|청구일|취급액|수수료율|수수료|부가세|계|매체확정금액|비고|TRANSCUSTRANK"
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "   4|	0|   4|       0|        12|         0|      12|         0|      12|       0|    10|           0|        12|    12|     8|    12|       8|    12|    10|12|           0|  10|            0"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "SEQ | AMT | COMMISSION | VAT | SUMAMTVAT | OUT_AMT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE ", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, " YEARMON | MED_FLAG | MED_FLAGNAME | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | DEPT_CD | DEPT_NAME | TITLE | MEMO | TRANSCUSTRANK", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, True, " YEARMON | MED_FLAG | MED_FLAGNAME | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | DEPT_CD | DEPT_NAME | TITLE | DEMANDDAY | TRANSCUSTRANK" 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, FALSE, " CHK " 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "YEARMON | MED_FLAG | CLIENTCODE | REAL_MED_CODE | MEDCODE | DEPT_CD | TRANSCUSTRANK", True
			
		End If
		.sprSht_DTL.style.visibility = "visible"
	End With
End Sub
'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	'초기 데이터 설정
	With frmThis
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		DateClean

		.txtPRINTDAY.value  = gNowDate
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		mstrGrid = FALSE
		CALL Grid_Setting ()
	End With
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn_CUST ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME
	chkcnt = 0
	
	With frmThis
		If mstrGrid Then Exit Sub

		intColFlag = 0
		For intCnt = 1 To .sprSht_DTL.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 Then
				chkcnt = chkcnt + 1
			End If

			'그룹최대값 설정
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSCUSTRANK",intCnt))
			If intColFlag < bsdiv Then
				intColFlag = bsdiv
			End If
		next
		If chkcnt = 0 Then
			gErrorMsgBox "거래명세서를 생성할 데이터를 체크 하십시오",""
			Exit Sub
		End If

		'저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'데이터 Validation
		If DataValidation =False Then Exit Sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | DEPT_CD | DEPT_NAME | TITLE | DEMANDDAY | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | OUT_AMT | MEMO | TRANSCUSTRANK")		

		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value

		intRtn = mobjMDSCAORCOMMI.ProcessRtn_CUST(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON,intColFlag)

   		If not gDoErrorRtn ("ProcessRtn_CUST") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "거래명세서가 생성되었습니다.","확인"

			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				SelectRtn
			Else
				initPageData
			End If
			DateClean
   		End If
   	End With
End Sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
	'On error resume next
	With frmThis
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "발행일은 필수 입력 사항 입니다.",""
			Exit Function
		End If
  	End With
	DataValidation = True
End Function

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON, strDEMANDYEARMON
	Dim strREAL_MED_CODE
   	Dim i, strCols
    
	'On error resume next
	With frmThis
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "조회시 년월은 반드시 넣어야 합니다.","조회입력오류"
			Exit Sub
		End If 
		
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0 : .sprSht_DTL.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value 
		strREAL_MED_CODE= .txtCLIENTCODE1.value
		
		CALL Grid_Setting()
		vntData = mobjMDSCAORCOMMI.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strREAL_MED_CODE)

		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			Else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If

   		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
   		vntData2 = mobjMDSCAORCOMMI.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
														strYEARMON, strREAL_MED_CODE)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData2,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			Else
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   		End If
   	End With
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols

	'On error resume next
	With frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)

		vntData = mobjMDSCAORCOMMI.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strTRANSYEARMON, strTRANSNO)

		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			Else
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = True
   		End If
   	End With
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0

		For lngCnt = 1 To .sprSht_DTL.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMISSION", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next

		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		Else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
   	Dim lngchkCnt
   	
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "삭제할 내역이 없습니다.","삭제안내!"
			Exit Sub
		End If
		
		For i = 1 To .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSNO",i)

				vntData = mobjMDSCAORCOMMI.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO)
				If mlngRowCnt > 0 Then
					gErrorMsgBox i & "행의 거래명세서는 세금계산서가 발생한 상세내역이 존재합니다.","삭제안내!"
					Exit Sub
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			Exit Sub
		End If

		If gDoErrorRtn ("DeleteRtn") Then Exit Sub
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then Exit Sub
		
		intCnt = 0
		mobjSCGLSpr.SetFlag  .sprSht_HDR, meINS_TRANS

		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO ")
		intRtn = mobjMDSCAORCOMMI.DeleteRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("DeleteRtn") Then
			'선택된 자료를 끝에서 부터 삭제
			for i = .sprSht_HDR.MaxRows To 1 Step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next

			gErrorMsgBox "거래명세서가 삭제되었습니다.","삭제안내!"
			If .sprSht_HDR.MaxRows > 0 Then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, 1,1
				
				mstrGrid = True
				CALL Grid_Setting ()
				SelectRtn_DTL 1,1
			Else
				mstrGrid = False
				SelectRtn
			End If
   		End If
	End With
	err.clear
End Sub

-->
		</script>
		<script language="javascript">
		//SMS 발송
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
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
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="220" background="../../../images/back_p.gIF"
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
											<td class="TITLE">AOR - 거래명세서 생성/조회/삭제</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
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
						<TABLE id="tblBody" height="93%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1,'')"
												width="60">청구년월</TD>
											<TD class="SEARCHDATA" style="WIDTH: 100px"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" size="7" name="txtYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="60">매체사</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 203px; HEIGHT: 22px"
													maxLength="100" align="left" size="27" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHDATA"></TD>
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
									</TABLE>
								</TD>
							<tr>
								<td style="HEIGHT: 93px">
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD class="TITLE" style="WIDTH: 100%; HEIGHT: 8px" vAlign="absmiddle"></TD>
										</TR>
										<TR>
											<TD class="TITLE" width="210" vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_normal.gif'
														align='absMiddle' border='0' name='btn_normal'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')">
													<IMG id='btn_multi' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_multi.gif'
														align='absMiddle' border='0' name='btn_multi'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')">
													<IMG id='btn_hide' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_hide.gif'
														align='absMiddle' border='0' name='btn_hide'></span>
											</TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																readOnly maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD style="FONT-SIZE: 12px; FONT-WEIGHT: bold"><span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">승인자:</span>
															&nbsp;<INPUT class="INPUT_L" id="txtEMPNAME" title="사원조회" style="WIDTH: 62px; HEIGHT: 22px" maxLength="255"
																align="left" size="5" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
																border="0" name="ImgEMPNO" title="승인권자선택"> <INPUT class="INPUT" id="txtEMPNO" title="사번조회" style="WIDTH: 46px; HEIGHT: 22px" readOnly
																maxLength="8" align="left" size="2" name="txtEMPNO">&nbsp; <IMG id="ImgConfirmRequest" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'" height="20" alt="선택된 거래명세서를 승인요청합니다." src="../../../images/ImgConfirmRequest.gIF"
																align="absMiddle" border="0" name="ImgConfirmRequest"></TD>
														<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="화면을 초기화 합니다."
																src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="선택된 거래명세서를 삭제합니다." src="../../../images/imgDelete.gIF" border="0"
																name="imgDelete"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR id="tblBody1">
								<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_HDR" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="4101">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD class="TITLE" align="left" width="400" height="22" vAlign="absmiddle"></TD>
											<TD class="TITLE" vAlign="absmiddle" align="left" width="500" height="22">청구일자 : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="브랜드명" style="WIDTH: 100px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" maxLength="100" size="32" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">&nbsp;&nbsp;&nbsp;&nbsp; 
												발행일자 : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="발행일자" style="WIDTH: 94px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="100" size="10" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalPrintday"></TD>
											<TD vAlign="middle" align="right" height="22">
												<!--Common Button Start-->
												<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgCRE" onmouseover="JavaScript:this.src='../../../images/ImgCREOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/ImgCRE.gif'" alt="거래명세서를 생성합니다."
																src="../../../images/ImgCRE.gif" border="0" name="ImgCRE"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="개별 거래명세서를 출력합니다.." src="../../../images/imgPrint.gIF" border="0"
																name="imgPrint"></TD>
														<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
													</TR>
												</TABLE>
												<!--Common Button End-->
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR id="tblBody2">
								<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
									<DIV id="pnlTab2" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="8440">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<!--Bottom Split End-->
						</TABLE>
						<!--Input Define Table End-->
					</TD>
				</TR>
				<!--Top TR End-->
			</TABLE>
		</FORM>
		<iframe id="frmSMS" style="WIDTH: 500px;DISPLAY: none;HEIGHT: 500px" name="frmSMS"></iframe>
	</body>
</HTML>
