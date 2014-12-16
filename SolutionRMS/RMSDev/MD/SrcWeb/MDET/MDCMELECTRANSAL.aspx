<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANSAL.aspx.vb" Inherits="MD.MDCMELECTRANSAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>거래명세서 생성</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 위수탁거래명세서 등록 화면(MDCMCATVTRANSAL.aspx)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMCATVTRANSAL.aspx
'기      능 : 위수탁거래명세서 입력/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/21 By HWANG DUCK SU
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDETELECTRANS, mobjMDETELEC
Dim mobjMDCMGET
Dim mstrCheck, mstrCheck1, mstrGrid

CONST meTAB = 9

mstrCheck=True
mstrCheck1=True
mstrGrid = FALSE

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
'입력 필드 숨기기
Sub Set_TBL_HIDDEN(byVal strmode)
	With frmThis
		If  strmode = "EXTENTION"  Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "70%"
			document.getElementById("tblSheet2").style.height = "30%"
		ELSEIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ELSEIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "70%"
		END IF
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
	IF frmThis.txtYEARMON1.value = "" and frmThis.txtCLIENTCODE1.value = "" then
		gErrorMsgBox "조회조건을 입력하시오.","조회안내"
		Exit Sub
	end if
	
	mstrGrid = FALSE
	
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

	IF frmThis.chkSPONSOR.checked = True then
		ProcessRtn_ALL
	Else
		ProcessRtn
	End if
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgALLPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData, vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
     				 'MDCMCATVTRANS_NEW.rpt
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = "" 
		vntDataTemp = mobjMDETELECTRANS.ProcessRtn_TEMP_ALL(gstrConfigXml, vntData, strUSERID)

		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "call printSetTimeout_All()", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout_All()
	Dim intRtn
	with frmThis
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

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
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjMDETELECTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "call printSetTimeout('" & strTRANSYEARMON & "', '" & strTRANSNO & "')", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout(strTRANSYEARMON, strTRANSNO)
	Dim intRtn, intRtn2
	with frmThis
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'intRtn2 = mobjMDETELECTRANS.DeleteRtnUpdate_PRINTSEQ(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
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
	
	with frmThis
		vntInParams = array(.txtYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "trans", "ELEC") 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			IF vntRet(3,0) = "완료" THEN
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(4,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			ELSE
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code값 저장
				.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			END IF
			selectRtn
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
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
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"","trans", "ELEC")
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					selectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub	

'-----------------------------------------------------------------------------------------
' 팁팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtYEARMON1.value), trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "trans", "ELEC") 
							
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSTIMPOP.aspx",vntInParams , 413,465)
		'TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE
		if isArray(vntRet) then
			.txtYEARMON1.value = trim(vntRet(0,0))
			.txtTIMCODE1.value = trim(vntRet(5,0))
			.txtTIMNAME1.value = trim(vntRet(1,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			.txtCLIENTNAME1.value = trim(vntRet(2,0))
			selectrtn
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetTRANSTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												  trim(.txtYEARMON1.value),trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
												  trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "", "trans", "ELEC")
			
			if not gDoErrorRtn ("GetTRANSTIMCODE") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = trim(vntData(0,1))
					.txtTIMCODE1.value = trim(vntData(5,1))
					.txtTIMNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(2,1))
					selectrtn
				Else
					Call TIMCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if	
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
				Else
					Call SUBSEQCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' 소재코드팝업 버튼
'-----------------------------------------------------------------------------------------
'광고주 시퀀스가져오기
Sub ImgMATTER1_onclick
	Call MATTERCODE1_POP()
	
End Sub

Sub MATTERCODE1_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		
		vntInParams = array(trim(.txtCLIENTNAME1.value), trim(.txtTIMNAME1.value), trim(.txtSUBSEQNAME1.value),"", _
							trim(.txtMATTERNAME1.value), "", "A") '<< 받아오는경우
							
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP_ALL.aspx",vntInParams , 780,630)
		if isArray(vntRet) then
			if .txtMATTERCODE1.value = vntRet(0,0) and .txtMATTERNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			.txtMATTERCODE1.value = trim(vntRet(0,0))	' 소재코드 표시
			.txtMATTERNAME1.value = trim(vntRet(1,0))	' 소재명 표시
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' 광고주코드 표시
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' 광고주명 표시
			.txtTIMCODE1.value = trim(vntRet(4,0))		' 팀코드 표시
			.txtTIMNAME1.value = trim(vntRet(5,0))		' 팀명 표시
			.txtSUBSEQ1.value = trim(vntRet(6,0))		' 브랜드 표시
			.txtSUBSEQNAME1.value = trim(vntRet(7,0))	' 브랜드명 표시
			'.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' 제작사코드 표시
			'.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' 제작사코드 표시
			'.txtDEPT_CD.value = trim(vntRet(10,0))		' 부서코드 표시
			'.txtDEPT_NAME.value = trim(vntRet(11,0))	' 부서명 표시
			'.txtCLIENTSUBCODE.value = trim(vntRet(12,0))	' 사업부코드 표시
			'.txtCLIENTSUBNAME.value = trim(vntRet(13,0))	' 사업부명 표시
			'.txtGREATCODE.value = trim(vntRet(14,0))	' 광고처코드 표시
			'.txtGREATNAME.value = trim(vntRet(15,0))	' 광고처명 표시
			
     	end if
	End with
	'gSetChange
	
End Sub

Sub txtMATTERNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME1.value),trim(.txtTIMNAME1.value), trim(.txtSUBSEQNAME1.value),"", _
											trim(.txtMATTERNAME1.value), "", "A")
			if not gDoErrorRtn ("GetMATTER") then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE1.value = trim(vntData(0,1))		' 브랜드 표시
					.txtMATTERNAME1.value = trim(vntData(1,1))	' 브랜드명 표시 2,3,6,7
					.txtCLIENTCODE1.value = trim(vntData(2,1))
					.txtCLIENTNAME1.value = trim(vntData(3,1))
					.txtTIMCODE1.value	  = trim(vntData(4,1))	' 팀코드 표시
					.txtTIMNAME1.value	  = trim(vntData(5,1))	' 팀명 표시
					.txtSUBSEQ1.value      = trim(vntData(6,1))
					.txtSUBSEQNAME1.value  = trim(vntData(7,1))
				Else
					Call MATTERCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub




'****************************************************************************************
'조회필드 onclick onkeydown 
'****************************************************************************************

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
' 게재일 달력
'****************************************************************************************
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

'-----------------------------------------------------------------------------------------
' 청구일자 세팅
'-----------------------------------------------------------------------------------------
Sub txtYEARMON1_onblur
	With frmThis
		If .txtYEARMON1.value <> "" AND Len(.txtYEARMON1.value) = 6 Then DateClean
	End With
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

Sub sprSht_CLIENT_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 AND Col > 1 then
			mstrGrid = TRUE
			Call SelectRtn_HDR (Col, Row)
			Call selectRtn_ELECTRIC_MEDIUM ()
			mstrGrid = false
		end if
	end with
End Sub


Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		elseif Row > 0 AND Col > 1 then
			mstrGrid = TRUE
			SelectRtn_DTL Col, Row
			mstrGrid = false
		end if
	end with
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		IF mstrGrid = false THEN
			if Row = 0 and Col = 1 then
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_DTL, 1, 1, , , "", , , , , mstrCheck1
				if mstrCheck1 = True then 
					mstrCheck1 = False
				elseif mstrCheck1 = False then 
					mstrCheck1 = True
				end if
				for intcnt = 1 to .sprSht_DTL.MaxRows
					sprSht_DTL_Change 1, intcnt
				next
			end if
		END IF
	end with
End Sub  


sub sprSht_CLIENT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CLIENT, ""
		end if
	end with
end sub

sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub

Sub sprSht_CLIENT_Keyup(KeyCode, Shift)
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
		Call SelectRtn_HDR (frmThis.sprSht_CLIENT.ActiveCol,frmThis.sprSht_CLIENT.ActiveRow)
		Call SelectRtn_ELECTRIC_MEDIUM ()
	End If
End Sub


Sub sprSht_HDR_Keyup(KeyCode, Shift)
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
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
			.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") Then
				strCOLUMN = "VAT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
				strCOLUMN = "SUMAMTVAT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))
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

Sub sprSht_DTL_Keyup(KeyCode, Shift)
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

	With frmThis
		If mstrGrid Then
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") _
			   or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
					strCOLUMN = "PRICE"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE")) Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
					
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			End If
		else
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")  THEN

				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT"))   Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
					
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			End If
		end if
	End With
End Sub


Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_HDR.MaxRows >0 Then
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") Then
				If .sprSht_HDR.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)
					
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
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,strCol,vntData_row(j))
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


Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
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
		If mstrGrid Then
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")   Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		ELSE
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or _
			       .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") THEN
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		END IF
		
	End With
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim i
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strSEQ
	Dim strPRINT_SEQ
	Dim intRtn

	with frmThis
		If mstrGrid Then
			'하위 그리드가 세금계산서 상세내역일때 출력순번을 정할때 발생
			'if Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRINT_SEQ") then
			'	for i=1 to .sprSht_DTL.MaxRows
			'		if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) <> "" then
			'			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) = 0 then
			'				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
			'			else
			'				
			'				if Row <> i then
			'					'입력한숫자와 같은 출력순번이 들어가있다면 오류
			'					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row) = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) then
			'						gErrorMsgBox "출력순서가 중복입력되었습니다.",""
			'						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
			'						.txtCLIENTNAME1.focus() 
			'						.sprSht_DTL.focus()
			'						EXIT SUB
			'					end if
			'				end if
			'			end if
			'		end if
			'	next
				
			'	strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSYEARMON",Row)
			'	strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSNO",Row)
			'	strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",Row)
			'	strPRINT_SEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row)
				
			'	intRtn = mobjMDETELECTRANS.UPDATE_PRINTSEQ(gstrConfigXml,strTRANSYEARMON,strTRANSNO, strSEQ, strPRINT_SEQ)
			'end if
		END IF
	end with	
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
	set mobjMDETELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")
	set mobjMDETELEC		= gCreateRemoteObject("cMDET.ccMDETELECT_MEDIUM")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    With frmThis
		
		'******************************************************************
		'거래명세서 광고주 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_CLIENT	
		mobjSCGLSpr.SpreadLayout .sprSht_CLIENT, 4, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_CLIENT, "TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.SetHeader .sprSht_CLIENT,		  "거래명세서년월|거래명세서번호|광고주코드|광고주명"
		mobjSCGLSpr.SetColWidth .sprSht_CLIENT, "-1", "             0|             0|         0|      20"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "0", "15"
		mobjSCGLSpr.SetCellsLock2 .sprSht_CLIENT, true, "CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.ColHidden .sprSht_CLIENT, "TRANSYEARMON | TRANSNO | CLIENTCODE", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_CLIENT, "CLIENTNAME" ,-1,-1,0,2,false
		.sprSht_CLIENT.style.visibility = "visible"
		
		'******************************************************************
		'거래명세서 헤더 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 12, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | TRANSYEARMON | TRANSNO | SEQ | VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | VAT | SUMAMTVAT | CNT"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "선택|거래명세서년월|거래명세서번호|순번|구분|승인|매체사코드|매체사|대행금액|부가세|계|상세행수"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "   4|            0|             0|   0|   6|   6|         0|    30|      20|    20|15|      10"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT | VAT | SUMAMTVAT |  CNT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | VAT | SUMAMTVAT "
		mobjSCGLSpr.ColHidden .sprSht_HDR, "REAL_MED_CODE", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "REAL_MED_NAME" ,-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "TRANSYEARMON | TRANSNO | SEQ | VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE" ,-1,-1,2,2,false
		.sprSht_HDR.style.visibility = "visible"
		
		
		'******************************************************************
		'' 청약내역 조회  및   거래명세서 디테일 
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 37, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		  "선택|거래명세서년월|거래명세서번호|순번|광고주코드|광고주명|사업부코드|사업부명|매체코드|매체명|매체사코드|매체사|브랜드코드|브랜드명|팀코드|팀명|소재코드|소재명|프로그램명|지역|방송요일|부서코드|청구일자|발행일자|공급가액|회수|광고대금|부가세|매체종류코드|매체종류|세금계산서년월|세금계산서번호|신탁순번|TRANSCUSTRANK | TRANSSPONRANK | 전표구분 | 제작사코드" 
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "  4|             0|	            0|  0|         0|	    0|          0|       0|       0|    15|         0|    15|         0|      15|     0|  15|       0|    15|        10|   6|      10|       0|       0|       8|       9|   5|       9|     9|            0|       9|             0|             0|       0|             0 |            5|         0|          0"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "PRICE | AMT| VAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht_DTL, "MEDNAME|REAL_MED_NAME|PROGRAM|ADLOCALFLAG|WEEKDAY|MED_FLAG|SUBSEQNAME|TIMNAME|MATTERNAME", -1, -1, 50
		mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | REAL_MED_CODE | DEPT_CD | DEMANDDAY | TAXYEARMON | TAXNO | SUBSEQ | TIMCODE | MATTERCODE | INPUT_MEDFLAG | TRANSCUSTRANK | VOCH_TYPE | EXCLIENTCODE", true	
		.sprSht_DTL.style.visibility = "visible"
		
    End With

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDETELEC = Nothing
	set mobjMDETELECTRANS = Nothing
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
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		DateClean

		.txtPRINTDAY.value  = gNowDate
		.sprSht_CLIENT.MaxRows = 0
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		'.chkSPONSOR.checked = TRUE
		.cmbREAL_MED_CODE1.selectedIndex = 0
		.cmbMEDGUBUN1.selectedIndex = 0

	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON1.value,1,4) & "-" & MID(frmThis.txtYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
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
		
		vntData_VOCH = mobjMDETELEC.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE",,,vntData_VOCH,,60 
			mobjSCGLSpr.TypeComboBox = True 
			
   		End If    
   	End With
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData_CLIENT, vntData_HDR , vntData_DTL
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE, strCLIENTNAME 
   	Dim strTIMCODE, strTIMNAME
	Dim strSUBSEQ, strSUBSEQNAME
	Dim strMATTERCODE , strMATTERNAME
	Dim strREAL_MED_CODE
	Dim strMEDGUBN
    
	'On error resume next
	with frmThis
	
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "조회시 년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		
		'Sheet초기화
		.sprSht_CLIENT.MaxRows = 0
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strSUBSEQ		= .txtSUBSEQ1.value
		strSUBSEQNAME	= .txtSUBSEQNAME1.value
		strMATTERCODE	= .txtMATTERCODE1.value
		strMATTERNAME	= .txtMATTERNAME1.value
		strREAL_MED_CODE= .cmbREAL_MED_CODE1.value
		strMEDGUBN		= .cmbMEDGUBUN1.value
		
		vntData_CLIENT = mobjMDETELECTRANS.SelectRtn_CLIENT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strCLIENTNAME, _
							   			strTIMCODE, strTIMNAME, strSUBSEQ, strSUBSEQNAME, strMATTERCODE, strMATTERNAME, strREAL_MED_CODE, strMEDGUBN)
		
		If not gDoErrorRtn ("SelectRtn_CLIENT") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_CLIENT,vntData_CLIENT,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				Call SelectRtn_HDR (1, 1)
   				Call SelectRtn_ELECTRIC_MEDIUM ()
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   		
   		

   	end with
End Sub

Sub SelectRtn_HDR (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strCLIENTCODE
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"TRANSYEARMON",Row)
		strCLIENTCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",Row)
				
		vntData = mobjMDETELECTRANS.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strCLIENTCODE)
																							
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   			mstrGrid = TRUE
   		End If
   	end with
End Sub


Sub SelectRtn_ELECTRIC_MEDIUM ()
	Dim vntData , vntDataCHK
	Dim strTRANSYEARMON
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE, strCLIENTNAME 
   	Dim strTIMCODE, strTIMNAME
	Dim strSUBSEQ, strSUBSEQNAME
	Dim strMATTERCODE , strMATTERNAME
	Dim strREAL_MED_CODE
	Dim strMEDGUBN
    
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strSUBSEQ		= .txtSUBSEQ1.value
		strSUBSEQNAME	= .txtSUBSEQNAME1.value
		strMATTERCODE	= .txtMATTERCODE1.value
		strMATTERNAME	= .txtMATTERNAME1.value
		strREAL_MED_CODE= .cmbREAL_MED_CODE1.value
		strMEDGUBN		= .cmbMEDGUBUN1.value
		
		If strCLIENTCODE = "" Then
   			strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",.sprSht_CLIENT.ActiveRow)
   		End If
   		
   		'청약 확정 하지 않은 데이터 검증한다. [브랜드][담당부서]
   		vntDataCHK = mobjMDETELECTRANS.SelectRtn_DTLCHK(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE)
		If not gDoErrorRtn ("SelectRtn_DTLCHK") Then
			IF mlngRowCnt > 0 THEN
			
			gErrorMsgBox "청약 내용검증이 되지 않은 데이터가 있습니다 ." & vbcrlf & "   청약 검증이 되지 않은 데이터는 거래명세서에서 보이지 않습니다.","거래명서세 안내"
			END IF
		END IF    		
		
		vntData = mobjMDETELECTRANS.SelectRtn_ELECTRIC_MEDIUM(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strCLIENTNAME, _
							   								  strTIMCODE, strTIMNAME, strSUBSEQ, strSUBSEQNAME, strMATTERCODE, strMATTERNAME, _
							   								  strREAL_MED_CODE, strMEDGUBN)
																							
		If not gDoErrorRtn ("SelectRtn_ELECTRIC_MEDIUM") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTL, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			else
   				gWriteText lblStatusDTL, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = false
   		End If
   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols, intcnt
    
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0
		intcnt = 1

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)
		
		vntData = mobjMDETELECTRANS.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
				for intcnt = 1 to .sprSht_DTL.MaxRows  
					mobjSCGLSpr.SetCellTypeStatic .sprSht_DTL, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CHK",intcnt," "
				next
				
   				gWriteText lblStatusDTL, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			else
   				gWriteText lblStatusDTL, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = TRUE
   		End If
   	end with
End Sub



'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM, IntVAT, IntVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		
		For lngCnt = 1 To .sprSht_DTL.MaxRows
			IntAMT = 0
			IntVAT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT", lngCnt)
			IntVAT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"VAT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
			IntVATSUM = IntVATSUM + IntVAT
		Next
		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
			.txtSUMVAT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			.txtSUMVAT.value = IntVATSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
			Call gFormatNumber(frmThis.txtSUMVAT,0,True)
		End If
	End With
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
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	Dim strVOCH_TYPE
	
	chkcnt = 0
	intColFlag = 0
	strVOCH_TYPE = 0
	
	with frmThis
		
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'그룹최대값 설정
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSSPONRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "거래명세서를 생성할 데이터를 체크 하십시오",""
			exit sub
		end if
			
		'저장플레그 설정

		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE")
	
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		
		intRtn = mobjMDETELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag,strVOCH_TYPE)
   		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "거래명세서가 생성되었습니다.","확인"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				.txtTIMCODE1.value = strTIMCODE
				.txtTIMNAME1.value = strTIMNAME
				selectRtn
			Else
				initpagedata
			End If
			DateClean
   		end if
   	end with
End Sub

Sub ProcessRtn_ALL ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	Dim strVOCH_TYPE
	
	chkcnt = 0
	strVOCH_TYPE=1
	with frmThis
		intColFlag = 0
		
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'그룹최대값 설정
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSCUSTRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "거래명세서를 생성할 데이터를 체크 하십시오",""
			exit sub
		end if
			
		'저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE")
	
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		
		msgbox "ProcessRtn_ALL"
		intRtn = mobjMDETELECTRANS.ProcessRtn_ALL(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag,strVOCH_TYPE)
   		if not gDoErrorRtn ("ProcessRtn_ALL") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "거래명세서가 생성되었습니다.","확인"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				.txtTIMCODE1.value = strTIMCODE
				.txtTIMNAME1.value = strTIMNAME
				selectRtn
			Else
				initpagedata
			End If
			DateClean
   		end if
   	end with
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
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "발행일은 필수 입력 사항 입니다.",""
			Exit Function
		End If
  	End with
	DataValidation = true
End Function


'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	Dim strPRINTDAY
   	Dim strMED_FLAG
   	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
   	Dim strCLIENTCODE, strCLIENTNAME
   	Dim lngchkCnt
   	
	with frmThis
		strDESCRIPTION = ""

		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "삭제할 내역이 없습니다.","삭제안내!"
			Exit Sub
		END IF
		
		For i = 1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSNO",i)
				
				'DB로 삭제전 확인을해야 여러개의 페이지끼리 중복삭제등의 예기치못한 오류를 막는다.(통일)
				'헤더의 TRANSYEARMON과 TRANSNO로 DTL의 상세내역(거래명세서 발생내역)이 존재하는지 조회
				vntData = mobjMDETELECTRANS.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO) 
				If mlngRowCnt > 0 Then
					gErrorMsgBox i & "행의 거래명세서는 세금계산서가 발생한 상세내역이 존재합니다.","삭제안내!"
					Exit Sub
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT SUB
		END IF
				
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0

		'그리드의 플래그를 INSERT로 바꾼다.  초기값은 UPDATE...
		mobjSCGLSpr.SetFlag  .sprSht_HDR, meINS_TRANS
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO ")
		
		intRtn = mobjMDETELECTRANS.DeleteRtn(gstrConfigXml,vntData)

		IF not gDoErrorRtn ("DeleteRtn") then
			'선택된 자료를 끝에서 부터 삭제
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "거래명세서가 삭제되었습니다.","삭제안내!"
			if .sprSht_HDR.MaxRows > 0 then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, 1,1
				mstrGrid = true
				SelectRtn_DTL 1,1
			else
				mstrGrid = FALSE
				SelectRtn
			end if
   		End IF
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
												<td class="TITLE">거래명세서 생성</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
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
												<TD class="SEARCHDATA" width="215"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">광고주</TD>
												<TD class="SEARCHDATA" width="215"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="17" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">매체사</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbREAL_MED_CODE1" title="매체사" name="cmbREAL_MED_CODE1">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="B00107">한국방송광고공사본사</OPTION>
														<OPTION value="B00111">한국방송광고공사부산지사</OPTION>
														<OPTION value="B00109">한국방송광고공사대구지사</OPTION>
														<OPTION value="B00110">한국방송광고공사대전지사</OPTION>
														<OPTION value="B00108">한국방송광고공사광주지사</OPTION>
														<OPTION value="B00112">한국방송광고공사전북지사</OPTION>
														<OPTION value="B01092">미디어크리에이트</OPTION>
													</SELECT><SELECT id="cmbMEDGUBUN1" title="매체구분" style="WIDTH: 56px" name="cmbMEDGUBUN1">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="01">TV</OPTION>
														<OPTION value="02">RADIO</OPTION>
														<OPTION value="10">DMB</OPTION>
													</SELECT>&nbsp;&nbsp;&nbsp;&nbsp; 협찬포함 <INPUT id="chkSPONSOR" title="협찬" type="checkbox" name="chkSPONSOR"></TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
														align="right" border="0" name="imgQuery">
												</td>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1,txtSUBSEQ1)">브랜드</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="브랜드코드명" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="16" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgSUBSEQ1"> <INPUT class="INPUT" id="txtSUBSEQ1" title="브랜드명조회" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" align="left" size="5" name="txtSUBSEQ1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)">팀</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="사업부코드명" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="11" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT" id="txtTIMCODE1" title="사업부명" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" align="left" size="5" name="txtTIMCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME1, txtMATTERCODE1)">소재명</TD>
												<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtMATTERNAME1" title="소재코드" style="WIDTH: 162px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="21" name="txtMATTERNAME1"> <IMG id="ImgMATTER1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgMATTER1"> <INPUT class="INPUT" id="txtMATTERCODE1" title="소재코드명" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtMATTERCODE1"></TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD class="TITLE" style="HEIGHT: 8px" vAlign="absmiddle"></TD>
											</TR>
											<TR>
												<TD class="TITLE" width="210"  vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_normal.gif' align='absMiddle' border='0' name='btn_normal'></span>&nbsp;
																<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')"><IMG id='btn_multi' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_multi.gif' align='absMiddle' border='0' name='btn_multi'></span>&nbsp;
																<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')"><IMG id='btn_hide' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_hide.gif' align='absMiddle' border='0' name='btn_hide'></span>
															</TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="297" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle" width="292">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
																<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<td class="TITLE" style="WIDTH: 340px" vAlign="absmiddle">청구일 : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="브랜드명" style="WIDTH: 85px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="8" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">&nbsp;&nbsp; 
													발행일 : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="발행일자" style="WIDTH: 85px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="8" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0"
														name="imgCalPrintday"></td>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="화면을 초기화 합니다."
																	src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
															<TD><IMG id="ImgAllCustSave" onmouseover="JavaScript:this.src='../../../images/ImgAllCustSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAllCustSave.gIF'"
																	height="20" alt="광고주별로 거래명세서를 일괄생성합니다.." src="../../../images/ImgAllCustSave.gIF"
																	border="0" name="ImgAllCustSave"></TD>
															<TD><IMG id="imgDeleteALL" onmouseover="JavaScript:this.src='../../../images/imgDeleteALLOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDeleteALL.gif'"
																	height="20" alt="선택된 거래명세서를 전체삭제합니다." src="../../../images/imgDeleteALL.gIF" border="0"
																	name="imgDeleteALL"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="선택된 거래명세서를 삭제합니다." src="../../../images/imgDelete.gIF" border="0"
																	name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" id="TD1" style="WIDTH: 100%; HEIGHT: 3px" runat="server"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR id="tblBody1">
									<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
										<table height="100%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
											<tr>
												<td width="20%">
													<DIV id="pnlTab_1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
														ms_positioning="GridLayout">
														<OBJECT id="sprSht_CLIENT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
															VIEWASTEXT DESIGNTIMEDRAGDROP="213">
															<PARAM NAME="_Version" VALUE="393216">
															<PARAM NAME="_ExtentX" VALUE="6324">
															<PARAM NAME="_ExtentY" VALUE="3519">
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
												</td>
												<td align="left" width="80%">
													<DIV id="pnlTab1_2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
														align="left" ms_positioning="GridLayout">
														<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
															<PARAM NAME="_Version" VALUE="393216">
															<PARAM NAME="_ExtentX" VALUE="25400">
															<PARAM NAME="_ExtentY" VALUE="3519">
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
												</td>
											</tr>
										</table>
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
												<TD align="left" width="297" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle" width="292">부가세합계 : <INPUT class="NOINPUTB_R" id="txtSUMVAT" title="부가세합계" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMVAT">
															</td>
														</tr>
													</table>
												</TD>
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
													<!--Common Button End--></TD>
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
									<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 70%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="8573">
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
									<TD class="BOTTOMSPLIT" id="lblStatusDTL" style="WIDTH: 100%"></TD>
								</TR>
								<TR>
									<TD></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
