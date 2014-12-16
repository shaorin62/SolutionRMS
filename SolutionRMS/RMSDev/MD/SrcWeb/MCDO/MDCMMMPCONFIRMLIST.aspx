<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMMPCONFIRMLIST.aspx.vb" Inherits="MD.MDCMMMPCONFIRMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>라이나MMP 청구 승인요청</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOCUSTLIST.aspx
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/08 By KTY
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
Dim mobjMDCOMMPCONFIRMLIST '공통코드, 클래스
Dim mobjMDCOGET
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	'EndPage
End Sub

'-----------------------------
' 닫기
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'---------------------------------------------------
' 명령 버튼 클릭 이벤트
'---------------------------------------------------

'-----------------------------------
'조회
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'청구 요청및 상세 내역 저장 
'-----------------------------------
Sub ImgConfirmRequest_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Confirm_User
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'인쇄 버튼 클릭시
'-----------------------------
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
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
		if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONFIRMGBN",.sprSht_HDR.ActiveRow) = "" then
			gErrorMsgBox "승인되지 않은데이터는 출력 하실 수 없습니다.","출력 안내!"
			exit sub
		end if 
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDCOMMPCONFIRMLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMMMPCOMMI.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",.sprSht_HDR.ActiveRow)
		'한번에 한개의 데이터만 출력 한다.(임시 추가분 월별로 한개의 데이터 생성 함)
		strCNT			= 1
		
		strUSERID = ""
		vntData = mobjMDCOMMPCONFIRMLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
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
		intRtn = mobjMDCOMMPCONFIRMLIST.DeleteRtn_temp(gstrConfigXml)
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
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
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
end sub

'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	With frmThis
		mobjSCGLSpr.CellChanged .sprSht_HDR, Col, Row
	End With
End Sub


Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	With frmThis
		
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub

'-----------------------------------
'쉬트 클릭
'-----------------------------------
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_DTL Col, Row
		End If
	End With
End Sub

'-----------------------------------
'쉬트 더블클릭
'-----------------------------------
sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		End If
	End With
End sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End sub

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
End Sub

Sub txtCLIENTNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
	
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjMDCOMMPCONFIRMLIST	= gCreateRemoteObject("cMDCO.ccMDCOMMPCONFIRMLIST")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjSCCOGET				= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		'MMP 기초데이터 시트
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 19, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "선택|승인유무|년월|순번|광고주코드|광고주명|매체사코드|매체사명|담당부서코드|담당부서명|청구년월|신규가입금액|수수료율|MMP청구금액|승인구분|요청자|요청일자|승인자|승인일자"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", " 4|       8|   8|   4|         0|      10|         0|      10|           0|         8|      10|          10|       8|         12|       0|     8|       8|     8|       8"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT | MMP_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
		mobjSCGLSpr.colhidden .sprSht_HDR, "CLIENTCODE | REAL_MED_CODE | DEPT_CD | CONFIRMFLAG",true
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, " SEQ | CONFIRMGBN " ,-1,-1,2,2,false
		
		
		'MMP 데이터 각매체별 현황 및 나눔 처리 시트
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 10, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "년월|순번|광고주코드|광고주명|매체구분|매체별합계|매체별금액|분담율|MMP금액|매체별분담금"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 8|   4|         8|      12|      10|        12|        12|     8|     12|          12"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "AMTSUM | AMT  | HDR_AMT | DIVAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG  ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT "
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ | CLIENTCODE ",-1,-1,2,2,False
		mobjSCGLSpr.CellGroupingEach .sprSht_DTL, "AMTSUM | HDR_AMT "
		
		.sprSht_HDR.style.visibility = "visible"
		.sprSht_DTL.style.visibility = "visible"
	
    End With

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCOMMPCONFIRMLIST = Nothing
	set mobjMDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis

	'초기 데이터 설정
	With frmThis
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	End With
End Sub

'------------------------------------------
' HDR 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON
   	Dim strCLIENTNAME
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'변수 초기화
		strYEARMON = "" : strCLIENTNAME = "" 
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTNAME	= .txtCLIENTNAME.value 
		
		vntData = mobjMDCOMMPCONFIRMLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTNAME)

		If not gDoErrorRtn ("SelectRtn_HDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_HDR, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht_HDR.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",intCnt) <> "" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, " CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
					ELSE
						mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, false, "CHK "
					END IF 
				Next
				
   				gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
	   			
   				Call SelectRtn_DTL(1,1)
   			else
   				gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				.sprSht_HDR.MaxRows = 0
				.sprSht_DTL.MaxRows = 0
			end if
   		End If
   	End With
End Sub

'------------------------------------------
' DTL 데이터 조회
'------------------------------------------
Sub SelectRtn_DTL(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, intCnt, intCnt2
	Dim strYEARMON
	Dim lngSEQ 
	Dim strCLIENTCODE
	Dim strREQUEST_USER
	Dim intAMT, intSUMAMT, intHDR_AMT
	
	With frmThis
		'sprSht2초기화
		.sprSht_DTL.MaxRows = 0
		strYEARMON = "" : lngSEQ = 0 : strREQUEST_USER = "" : strCLIENTCODE = ""
		intAMT = 0		: intSUMAMT = 0 : intHDR_AMT = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",Row)
		lngSEQ			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",Row)
		strCLIENTCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTCODE",Row)
		strREQUEST_USER	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",Row)
		intHDR_AMT		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MMP_AMT",Row)
		
		
		vntData = mobjMDCOMMPCONFIRMLIST.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, _ 
														lngSEQ,strCLIENTCODE,strREQUEST_USER )
														

		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",Row) <> "" Then
					
					mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT"
				
				ELSE
					For intCnt = 1 To .sprSht_DTL.MaxRows
						'승인되지 않은데이터의 SEQ 에 HDR의 키를 넣는다.
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"SEQ",intCnt , lngSEQ
						
						'DTL 내역의 합계금액을 구한다.
						intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",intCnt)
						intSUMAMT = intSUMAMT + intAMT
					Next
					
					'반복하며 하단의 데이터 세팅
					for intCnt2 = 1 To .sprSht_DTL.MaxRows
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"AMTSUM",intCnt2 , intSUMAMT
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"HDR_AMT",intCnt2 , intHDR_AMT
					next
					
					'각매체의 금액의 분할율을 구하여 HDR 총금액을 분할율로 매체별 재분할 실시
					call DIV_RATE (intSUMAMT,intHDR_AMT)
				END IF
			ELSE
				.sprSht_DTL.MaxRows = 0
			End If	
		End If
   		gWriteText lblStatusDTR, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
	End With
End Sub

SUB DIV_RATE(ByVal intSUMAMT, ByVal intHDR_AMT)
	Dim i
	Dim intAMT
	Dim intRATE
	Dim intDIVAMT
	With frmThis
		
		for i = 1 To .sprSht_DTL.MaxRows
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",i)
			
			intRATE = (clng(intAMT) / clng(intSUMAMT)) * 100
			
			intDIVAMT = intHDR_AMT * intRATE / 100
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"RATE",i , intRATE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"DIVAMT",i , intDIVAMT
		next
		
	End With
END SUB
	
Sub ProcessRtn_Confirm_User ()
   	Dim vntData
   	Dim VNTDATA_DTL
   	Dim intRtn
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
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTNAME",intCnt2)
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
			
			vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
			
			'보내는사람정보
			strFromUserName		= vntData_info(0,2)
			strFromUserEmail	= vntData_info(1,2)
			strFromUserPhone	= vntData_info(2,2)

			'받는사람 정보
			strToUserName		=  vntData_info(0,1)
			strToUserEmail		=  vntData_info(1,1)
			strToUserPhone		=  vntData_info(2,1)
		
			call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
			
			'저장플레그 설정
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meINS_TRANS
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS

			'쉬트의 변경된 데이터만 가져온다.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | YEARMON | SEQ")
			'하단의 분담된 금액도 함께 저장 한다.
			vntData_DTL = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT")
			
			intRtn = mobjMDCOMMPCONFIRMLIST.ProcessRtn_Confirm_User(gstrConfigXml,vntData,vntData_DTL, .txtEMPNO.value)

   			if not gDoErrorRtn ("ProcessRtn_Confirm_User") then
				'모든 플래그 클리어
				mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
				initpagedata	
				gOkMsgBox "승인요청이 되었습니다.","확인"
				
				If intRtn <> 0  Then
					.txtCLIENTNAME1.value = strCLIENTNAME
					selectRtn
				Else
					initpagedata
				End If
				DateClean
   			end if
   		End If
   	end with
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
													<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
												<td class="TITLE">라이나 MMP 기초대이터 승인요청</td>
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
									<TD align="left" width="100%" height="1">
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								
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
												<TD class="SEARCHLABEL" style="WIDTH: 42px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">년월</TD>
												<TD class="SEARCHDATA" width="113" style="WIDTH: 113px"><INPUT class="INPUT" id="txtYEARMON" title="년월" style="WIDTH : 100px; HEIGHT : 22px" maxLength="100"
														align="left" size="6" name="txtYEARMON" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 53px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, '')">광고주명</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 200px; HEIGHT: 22px"
														maxLength="100" align="left" size="28" name="txtCLIENTNAME"></TD>
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
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
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
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'" height="20" alt="선택된 데이터를 승인요청합니다." src="../../../images/ImgConfirmRequest.gIF"
																	align="absMiddle" border="0" name="ImgConfirmRequest"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End-->
												</TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
											ms_positioning="GridLayout">
											<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_HDR" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="7117">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
								</TR>
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="승인된 데이터를 출력합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
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
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
											ms_positioning="GridLayout">
											<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="7117">
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
							</TABLE>
						</TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
		<iframe id="frmSMS" style="WIDTH: 500px;DISPLAY: none;HEIGHT: 500px" name="frmSMS"></iframe> <!-- -->
	</body>
</HTML>
