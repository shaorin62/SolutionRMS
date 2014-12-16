<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOB.aspx.vb" Inherits="PD.PDCMJOB" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOB.aspx
'기      능 : JOBLIST 조회
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
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
Dim mobjPDCOJOB
Dim mobjSCCOGET
Dim mobjPDCOGET
Dim mlngRowCnt,mlngColCnt
Dim mstrCheck
Dim mstrSortCol
Dim mstrSortOrder
Dim mstrSortOrderCnt

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
'조회
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'저장
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'옵션엑셀
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'닫기
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------
' 광고주 및 JOBNO 팝업 버튼[조회용]
'------------------------------------
Sub ImgCLIENTCODE1_onclick
	with frmThis
		IF .cmbSEARCH.value = "1" then
			Call CLIENTCODE1_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	End With
End Sub

'광고주 - 실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' 코드명 표시		
     	end if
	End with
	imgQuery_onclick
	gSetChange
End Sub

'JOBNO - 실제 데이터List 가져오기
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array( trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	imgQuery_onclick
	gSetChange
End Sub

'광고주 또는 JOBNO 한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			if .cmbSEARCH.value = "1" Then '프로젝트 코드 라면
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
				if not gDoErrorRtn ("GetHIGHCUSTCODE") then
					If mlngRowCnt = 1 Then
						.txtCLIENTCODE1.value = trim(vntData(0,1))
						.txtCLIENTNAME1.value = trim(vntData(1,1))
   						imgQuery_onclick
					Else
						Call CLIENTCODE1_POP()
					End If
   				end if
   			Else
   				vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value))
				
				if not gDoErrorRtn ("GetJOBNO") then
					If mlngRowCnt = 1 Then
						.txtCLIENTCODE1.value = trim(vntData(0,0))
						.txtCLIENTNAME1.value = trim(vntData(1,0))
   						imgQuery_onclick
					Else
						Call SEARCHJOB_POP()
					End If
   				end if
   			End If
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
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
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value))
		
		vntRet = gShowModalWindow("../../../PD/SrcWeb/PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			gSetChangeFlag .txtEMPNO
			gSetChangeFlag .txtEMPNAME
     	end if
	End with
	imgQuery_onclick
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
			vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetPDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
   					imgQuery_onclick
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' 달력
'****************************************************************************************
'조회용
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

'****************************************************************************************
' 조회필드 체인지
'****************************************************************************************
Sub cmbENDFLAG1_onchange()
	imgQuery_onclick
end Sub

'****************************************************************************************
' 입력필드 체인지
'****************************************************************************************
Sub txtCHECK_POINT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CHECK_POINT",frmThis.sprSht.ActiveRow, frmThis.txtCHECK_POINT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub cmbSEARCH_onchange
	with frmThis
		.txtCLIENTNAME1.value = ""
		.txtCLIENTCODE1.value = ""
	End with
	gSetChange
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	With frmThis
		if Row = 0 and Col = 1 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			
			for intcnt = 1 to .sprSht.MaxRows
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intcnt
			Next
		End If
		
		If Row <> 0 Then sprShtToFieldBinding Col,Row
		
	end With
End Sub

Sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strWith, strHeight
	Dim strJOBNO, strSUBNO
	Dim strPREESTNO, strPRIJOBNAME, strPROJECTNM
	Dim strCLIENTCODE, strCLIENTNAME
	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
	Dim strTIMCODE, strTIMNAME
	Dim strSUBSEQ, strSUBSEQNAME
	Dim strJOBGUBN, strJOBGUBNNAME, JOBPARTNAME

	With frmThis
		If Row = 0 and Col >1 Then
			mstrSortCol = Col
			if mstrSortOrderCnt = 1 then
				mobjSCGLSpr.SetSheetSortUser  .sprSht, mstrSortCol, 1
				mstrSortOrderCnt = 2
			else
				mobjSCGLSpr.SetSheetSortUser  .sprSht, mstrSortCol, 2
				mstrSortOrderCnt = 1
			end if 
		Else
			strWith =  Screen.width
			strHeight =  Screen.height - 100
			
			strJOBNO		= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strSUBNO		= mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",.sprSht.ActiveRow)
			strJOBNAME		= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",.sprSht.ActiveRow)	
			strPREESTNO		= mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",.sprSht.ActiveRow)		
			strPRIJOBNAME	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",.sprSht.ActiveRow)	
			strPROJECTNM	= mobjSCGLSpr.GetTextBinding( .sprSht,"PROJECTNM",.sprSht.ActiveRow) 
			strCLIENTNAME	= mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow) 
			strJOBGUBNNAME  = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBNNAME",.sprSht.ActiveRow) 
			strCLIENTCODE	= mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow)	 
			strTIMCODE		= mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",.sprSht.ActiveRow)	
			strSUBSEQ		= mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",.sprSht.ActiveRow)
			strJOBGUBN		= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBN",.sprSht.ActiveRow)	
			strJOBPARTNAME	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBPARTNAME",.sprSht.ActiveRow)	
			
			vntInParams = array(strJOBNO, strSUBNO, strJOBNAME, strPREESTNO, strPRIJOBNAME, strPROJECTNM, _
								strCLIENTNAME, strJOBGUBNNAME, strCLIENTCODE, strTIMCODE, strSUBSEQ, _
								strJOBGUBN, strJOBPARTNAME)
								
			vntRet = gShowModalWindow("PDCMJOBMST.aspx",vntInParams , strWith, strHeight)

			imgQuery_onclick
		End If
	End With
End Sub

'스프레드 내 버튼을 클릭 하였을때 발생 하는 이벤트
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	
	with frmThis
	    '담당자내역 버튼 부분
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",Row))
			vntRet = gShowModalWindow("PDCMACTUALRATELISTPOP.aspx",vntInParams , 815,700)
		end if
	End with
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strSUM
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	
	With frmThis
		If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
			sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
					strCOLUMN = "DIVAMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
					strCOLUMN = "ADJAMT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) Then
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
		End If
	End With
END SUB

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	Dim strColFlag
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
		strColFlag = 0
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
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
	With frmThis
		mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	end with
End Sub
	
'-------------------------------------------------
''시트에 데이터한로우의 정보를 헤더 필더에 바인딩
'-------------------------------------------------
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '그리드 데이터가 없으면 나간다.
	
'		.txtCHECK_POINT.value = replace(mobjSCGLSpr.GetTextBinding(.sprSht,"CHECK_POINT",Row),"★", vbCrlf)
		.txtCHECK_POINT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CHECK_POINT",Row)
		
   	end With
End Function

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성									  
	set mobjPDCOJOB = gCreateRemoteObject("cPDCO.ccPDCOJOB")
	set mobjPDCOGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
	
	gSetSheetDefaultColor
	with frmThis
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 36, 0, 6, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | REQDAY | ENDFLAG | ENDFLAGNAME | PROJECTNM | JOBNAME | JOBNO | PREESTNO | SEQ | CLIENTNAME | TIMNAME | JOBGUBNNAME | JOBPARTNAME | DIVAMT | ADJAMT | DEPTNAME | EMPNAME | EMPCNT | BTN | CPEMPNAME | DEMANDYEARMON | ADJDAY | CONTRACTNO | CHECK_POINT | JOBNO_DIVAMTFLAG | RANKJOB | PRIJOBNAME | JOBGUBN | SUBSEQ | SUBSEQNAME | CLIENTCODE | TIMCODE | COMMITIONVALUE | RATE | INCOM | SETYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		    "선택|등록일|상태코드|상태|프로젝트명|JOB명|JOBNO|PREESTNO|SUBNO|광고주|팀|매체부문|매체분류|견적금액|청구금액|담당부서|담당자|담당자수|담당자내역|등록자|청구일|외주정산일|미계약|CHECKPOINT|job/div위치|그룹핑|대표JOB명|JOB구분코드|브랜드코드|브랜드|광고주코드|팀코드|수수료율|내수율|내수액|결산일"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|    8|        0|   5|        25|   25|    7|       0|    0|    18|18|       7|       7|      11|      11|       9|     9|       9|		 9|     9|     9|         0|     0|        20|         0|      0|        0|          0|         0|     0|         0|     0|       0|    10|    11|     0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"담당자내역", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT | ADJAMT | INCOM | EMPCNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY | DEMANDYEARMON", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHECK_POINT", -1, -1, 1000
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PROJECTNM | JOBNAME | CLIENTNAME | TIMNAME | CPEMPNAME |  CHECK_POINT | DEPTNAME",-1,-1,0,2,false ' 왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht, "REQDAY | ENDFLAGNAME | JOBNO | PREESTNO | SEQ | JOBGUBNNAME | DEMANDYEARMON | ADJDAY | CONTRACTNO | SETYEARMON | JOBNO_DIVAMTFLAG|JOBPARTNAME | EMPNAME",-1,-1,2,2,false  '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"REQDAY | ENDFLAG | ENDFLAGNAME | PROJECTNM | JOBNAME | JOBNO | PREESTNO | SEQ | CLIENTNAME | TIMNAME | JOBGUBNNAME | DIVAMT | ADJAMT | CPEMPNAME | DEMANDYEARMON | ADJDAY | CONTRACTNO | SETYEARMON | CHECK_POINT | JOBNO_DIVAMTFLAG | INCOM|RATE"
		mobjSCGLSpr.ColHidden .sprSht, "ENDFLAG | PREESTNO | JOBNO_DIVAMTFLAG | RANKJOB |  PRIJOBNAME | CHECK_POINT | JOBGUBN | SUBSEQ | SUBSEQNAME | CLIENTCODE | TIMCODE | COMMITIONVALUE | EMPCNT", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"JOBNO"
		pnlTab1.style.visibility = "visible" 
	End with
    
	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjPDCOJOB = Nothing
	set mobjSCCOGET = Nothing
	set mobjPDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
		
		.txtFROM.focus
		DateClean
		'.txtFROM.value = ""
		.cmbSEARCH.value = "1"
		Call SEARCHCOMBO_TYPE()
		.cmbENDFLAG1.selectedIndex = -1
	End with
	
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

'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()
	Dim vntENDFLAG
	Dim vntJOBTYPE
  
    With frmThis   

		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntENDFLAG = mobjPDCOJOB.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  'JOB종류 호출
		vntJOBTYPE = mobjPDCOJOB.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB종류 호출
		
		if not gDoErrorRtn ("GetDataType") then 
			mobjSCGLSpr.TypeComboBox = True 
			gLoadComboBox .cmbENDFLAG1, vntENDFLAG, False
			gLoadComboBox .cmbJOBTYPE,  vntJOBTYPE, False
   		end if    				   		
   	end with     
End Sub

'=========================================================================================
' 데이터 조회
'=========================================================================================
Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim intCnt
    Dim strFROM ,strTO
    
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO	= MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		vntData = mobjPDCOJOB.SelectRtn(gstrConfigXml,mlngRowCnt, mlngColCnt, TRIM(strFROM), TRIM(strTO), _
										TRIM(.txtCLIENTCODE1.value), TRIM(.txtCLIENTNAME1.value), _
										TRIM(.cmbENDFLAG1.value), TRIM(.cmbSEARCH.value), TRIM(.cmbJOBTYPE.value), _
										TRIM(.txtEMPNO.value),TRIM(.txtEMPNAME.value) )
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then

				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows '조회된 내역을 처음부터 끝까지 돌면서
					'JOB별 컬러 통일
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKJOB",intCnt) Mod 2 = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"EMPCNT",intCnt) > 1 Then
						mobjSCGLSpr.SetCellShadow .sprSht, 16, 17, intCnt, intCnt,&HCCFFFF, &H000000,False
					End If
					
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHECK_POINT",intCnt) <> "" Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CHECK_POINT",intCnt, replace(mobjSCGLSpr.GetTextBinding(.sprSht,"CHECK_POINT",intCnt),"★", vbCrlf)
					END IF
				Next
				
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			Else	
   				.sprSht.MaxRows = 0
   				gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   			end If
   		end if
   		
   		mobjSCGLSpr.SetSheetSortUser  .sprSht, mstrSortCol
   	end with
   	'검색시에 첫행을 MASTER와 바인딩 시키기 위함
    sprShtToFieldBinding 2, 1
   	AMT_SUM
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
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT", lngCnt)
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
	Dim strSEQ 
	Dim strDataCHK
	Dim lngCol, lngRow
	Dim strCHECK_POINT
	Dim intCnt
	Dim lngSum
	Dim intRtnSave
	
	With frmThis
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | REQDAY | ENDFLAG | PROJECTNM | JOBNAME | JOBNO | PREESTNO | SEQ | CLIENTNAME | TIMNAME | JOBGUBNNAME | DIVAMT | ADJAMT | CPEMPNAME | DEMANDYEARMON | ADJDAY | CONTRACTNO | SETYEARMON | CHECK_POINT|JOBNO_DIVAMTFLAG")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		lngSum = 0
		strCHECK_POINT = .txtCHECK_POINT.value
		
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK", intCnt) = "1" Then
				lngSum = lngSum + 1
			End If
		Next
		
		If lngSum <> "0" And strCHECK_POINT = "" Then
			intRtnSave = gYesNoMsgbox("하나라도 선택을 하였을 경우 반드시 Check Point 를 입력하셔야 합니다." & vbcrlf & "선택된 데이터를 값없음 으로 변경하시겠습니까?","처리안내")
			IF intRtnSave <> vbYes then 
				exit Sub
			Else
				.txtCHECK_POINT.focus()
			End If
		End If
		
		intRtn = mobjPDCOJOB.ProcessRtn(gstrConfigXml, vntData, strCHECK_POINT)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			imgQuery_onclick
			.sprSht.focus()
   		End If
   	end With
End Sub


		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR width="1024">
					<td></td>
				</TR>
				<TR height="85%">
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="27">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF"
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
											<td class="TITLE">JOB관리</td>
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
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="50" border="0">
										<TR>
											<TD></TD>
										</TR>
									</TABLE>
								</TD> <!--Common Button End--></TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 95%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="left" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()" width="70">등록일</TD>
											<TD class="SEARCHDATA" width="224"><INPUT class="INPUT" id="txtFROM" title="기간검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"> <IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
													border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="기간검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
													type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
													align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1,txtCLIENTCODE1)"
												width="84"><SELECT id="cmbSEARCH" title="광고주,JOBNO" style="WIDTH: 88px" name="cmbSEARCH">
													<OPTION value="1" selected>광고주</OPTION>
													<OPTION value="2">JOBNO</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA" style="WIDTH: 222px" width="222"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="조회용광고주명" style="WIDTH: 140px; HEIGHT: 22px"
													type="text" maxLength="100" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="조회용광고주코드" style="WIDTH: 57px; HEIGHT: 22px"
													type="text" maxLength="7" size="4" name="txtCLIENTCODE1"></TD>
											<td class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()">상태/유형</TD>
											<TD class="SEARCHDATA"><SELECT dataFld="ENDFLAG" id="cmbENDFLAG1" title="완료구분" style="WIDTH: 80px" dataSrc="#xmlBind"
													name="cmbENDFLAG1"></SELECT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<select id="cmbJOBTYPE" style="WIDTH: 88px">
													<OPTION value="" selected></OPTION>
													<OPTION value="B">B:</OPTION>
													<OPTION value="C">C:</OPTION>
													<OPTION value="D">D:</OPTION>
													<OPTION value="G">G:</OPTION>
													<OPTION value="I">I:</OPTION>
													<OPTION value="O">O:</OPTION>
													<OPTION value="P">P:</OPTION>
													<OPTION value="R">R:</OPTION>
													<OPTION value="S">S:</OPTION>
												</select></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNAME,txtEMPNO)"
												width="84"><SELECT id="cmbUSERNAME" title="사원명선택" style="WIDTH: 88px" name="cmbUSERNAME">
													<OPTION value="1" selected>등록자</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtEMPNAME" title="승인권자" style="WIDTH: 140px; HEIGHT: 22px"
													type="text" maxLength="100" size="16" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													title="승인권자선택" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgEMPNO"> <INPUT class="INPUT" id="txtEMPNO" title="승인권자사번" style="WIDTH: 57px; HEIGHT: 22px" type="text"
													maxLength="7" size="4" name="txtEMPNO">
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 15px"></TD>
							</TR>
							<!--추가-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 60%" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton2" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50"
													border="0">
													<TR>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<!--테스트 끝--></TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit Start-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="12118">
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
								<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<!--Brench End-->
				<!--Bottom Split Start-->
				<TR height="15%">
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="150" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="68" background="../../../images/back_p.gIF"
													border="0">
													<TR>
														<TD align="left" width="100%" height="2"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td class="TITLE" title="개별저장은 체크를 푸시고, 공통저장은 체크를 하여주십시오.">Check Point</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<table height="60%" cellSpacing="0" cellPadding="0" width="100%" border="1">
							<tr>
								<td align="center">
									<!--BACKGROUND-COLOR: #ebf2fa--><textarea dataFld="CHECK_POINT" id="txtCHECK_POINT" style="WIDTH: 100%; HEIGHT: 100%" dataSrc="#xmlBind"
										name="txtCHECK_POINT" wrap="hard"></textarea></td>
							</tr>
						</table>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
