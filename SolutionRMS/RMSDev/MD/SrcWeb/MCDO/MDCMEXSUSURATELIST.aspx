<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMEXSUSURATELIST.aspx.vb" Inherits="MD.MDCMEXSUSURATELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>제작대행사 수수료율 등록</title>
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
'HISTORY    :1) 2009/11/25 By kty
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
Dim mobjMDCOEXSUSURATE
Dim mobjMDCOGET
Dim mlngRowCnt,mlngColCnt
Dim mstrRow

mstrRow= 1
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
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

sub imgAddRow_onclick ()
	With frmThis
		call sprSht_DTL_Keydown(meINS_ROW, 0)
		.txtEXCLIENTCODE.focus
		.sprSht_DTL.focus
	End With 
end sub


Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
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
			selectrtn
     	end if
	End with
	
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
			vntData = mobjMDCOGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(0,0))
					.txtEXCLIENTNAME.value = trim(vntData(1,0))
					selectrtn
				Else
					Call EXCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'****************************************************************************************
' 쉬트 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
'		if Row = 0 and Col = 1 then
'			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
'			if mstrCheck = True then 
'				mstrCheck = False
'			elseif mstrCheck = False then 
'				mstrCheck = True
'			end if
'			for intcnt = 1 to .sprSht_HDR.MaxRows
'				sprSht_HDR_Change 1, intcnt
'			next
'		elseif Row > 0 AND Col > 1 then

		if Row > 0 AND Col > 1 then
			SelectRtn_DTL Col, Row
		end if
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		end if
	end with
end sub


Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim strCode
	Dim strCodeName
	Dim vntData
	
	with frmThis
		If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"RATE", Row) > 100 Then
			gErrorMsgBox "비율은 100을넘을수 없습니다.","처리안내"
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"RATE",Row, 0
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"FDATE") Then
			strFDATE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"FDATE",Row) 
			strEXCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EXCLIENTCODE",Row) 
			
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strEDATE = mobjMDCOEXSUSURATE.SelectRtn_FDATE(gstrConfigXml,mlngRowCnt,mlngColCnt,strFDATE, strEXCLIENTCODE)
			
			IF strEDATE <> "" THEN
				gErrorMsgBox "중복된 날짜입니다.","처리안내"	
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"FDATE",Row, strEDATE
			END IF 
			
		END IF 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EDATE") Then
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EDATE",Row) <> "" THEN
				IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"FDATE",Row) > mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EDATE",Row)  THEN
					gErrorMsgBox "종료일은 시작일보다 작을수없습니다.","처리안내"	
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EDATE",Row, mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"FDATE",Row)
				END IF 
			END IF 
		END IF 
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht_DTL, Col-1,Row
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME"), Row
						.sprSht_DTL.focus 
						mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+1, Row
					End If
   				End If
   			End If
		End If
		mobjSCGLSpr.SetTextBinding .sprSht_DTL,"INPUTFLAG",Row, "INPUT"
	End With
	
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub


Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		
		.sprSht_DTL.Focus
	End With
End Sub

'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If mstrPROCESS = True Then
			frmThis.sprSht.MaxRows = 0
		End If
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	
	with frmThis
	
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		
		IF .sprSht.MaxRows = 0 THEN
			gErrorMsgBox "조회된 대행사가 없습니다.","저장안내"
			EXIT SUB
		END IF 
		
		
		'IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EDATE",.sprSht_DTL.MaxRows) = "" THEN
		'	gErrorMsgBox "종료일을 넣어야 추가할수 있습니다.","저장안내"
		'	EXIT SUB
		'END IF 
			
		'IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",.sprSht_DTL.MaxRows) = "" THEN
		'	EXIT SUB
		'END IF 
		
		
		If KeyCode = meINS_ROW Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1)
					
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,false,"FDATE | EDATE | RATE | BIGO | CLIENTCODE | BTN | CLIENTNAME ",.sprSht_DTL.MaxRows,.sprSht_DTL.MaxRows,false
			
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"YEARMON",.sprSht_DTL.ACTIVEROW, Mid(gNowDate,1,4) & Mid(gNowDate,6, 2)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",.sprSht_DTL.ACTIVEROW, mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",.sprSht.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"RATE",.sprSht_DTL.ACTIVEROW, 0
			mobjSCGLSpr.ActiveCell .sprSht_DTL, 1,.sprSht_DTL.MaxRows
			.txtEXCLIENTNAME.focus
			.sprSht_DTL.focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, 3,2
		End If
		
	End with
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
		SelectRtn_DTL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		
	End If
End Sub

'시트 버튼 클릭 이벤트
Sub sprSht_DTL_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
			.sprSht_DTL.Focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, Col, Row
	
		End If
	End with
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjMDCOEXSUSURATE = gCreateRemoteObject("cMDCO.ccMDCOEXSUSURATE")
	set mobjMDCOGET	   = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'*** 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "YEARMON | EXCLIENTCODE | EXCLIENTNAME | GBN"
		mobjSCGLSpr.SetHeader .sprSht,		    "등록년월|대행사코드|제작대행사명|AOR구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "       6|         9|          17|      6"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "YEARMON | EXCLIENTCODE | EXCLIENTNAME | GBN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | EXCLIENTCODE | GBN",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EXCLIENTNAME ",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON", true
		
		
		'**************************************************
		'***디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 14, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 9, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_DTL,    "YEARMON | EXCLIENTCODE | EXCLIENTNAME | SEQ | FDATE | EDATE | RATE | BIGO | CLIENTCODE | BTN | CLIENTNAME | INPUTFLAG | CUSER | UUSER"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		    "등록년월|대행사코드|제작대행사명|순번|시작일|종료일|수수료율|비고|광고주코드|광고주명|INPUTFLAG|생성자|수정자"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1",  "       9|         9|          30|   4|     9|     9|      12|  15|         7|2|    12|        0|     7|     7"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "SEQ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "FDATE | EDATE "
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_DTL,"..", "BTN"
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true, "YEARMON | EXCLIENTCODE | EXCLIENTNAME | SEQ | FDATE | EDATE | RATE | BIGO | CLIENTCODE | BTN | CLIENTNAME | CUSER | UUSER"
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ |YEARMON | EXCLIENTCODE | CLIENTCODE | CLIENTNAME | CUSER | UUSER ",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "EXCLIENTNAME | BIGO",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht_DTL, "YEARMON | EXCLIENTCODE | EXCLIENTNAME | INPUTFLAG", true
		
		
	End with

	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCOEXSUSURATE = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0
	
	'.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	End with
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCOEXSUSURATE.SelectRtn_EXCLIENT(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value  )

		if not gDoErrorRtn ("SelectRtn_EXCLIENT") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				SelectRtn_DTL 1,mstrRow
   			Else
   				.sprSht.MaxRows = 0
   				.sprSht_DTL.MaxRows = 0
   			end If 
			
   			mobjSCGLSpr.ActiveCell .sprSht, 2, mstrRow
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		
   	end with
End Sub



Sub SelectRtn_DTL (Col, Row)
	Dim vntData
   	Dim i, strCols
   	Dim strEXCLIENTCODE
    
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strEXCLIENTCODE	= mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
				
		vntData = mobjMDCOEXSUSURATE.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strEXCLIENTCODE)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   			else
   				.sprSht_DTL.MaxRows = 0
   			End If
   			
   			
   			For intCnt = 1 To .sprSht_DTL.MaxRows 
				If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EDATE",intCnt) <> "" Then
					mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HAAF290, &H000000,False 
				ELSE
					mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
				End If 
				
				If  mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"FDATE",intCnt) <> "" AND  mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EDATE",intCnt) <> "" Then
					mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,TRUE,"FDATE | EDATE | RATE | BIGO",intCnt,intCnt,false
					mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,FALSE,"CLIENTCODE | BTN | CLIENTNAME ",intCnt,intCnt,false
				ELSE
					mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,FALSE,"EDATE | RATE | BIGO | CLIENTCODE | BTN | CLIENTNAME ",intCnt,intCnt,false
				END IF
				
			Next
   			
   			gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		End If
   	end with
End Sub



Sub ProcessRtn()

	Dim intRtn
   	Dim vntData
   	Dim intCnt
   	Dim strFDATE,strEDATE, strEXCLIENTCODE
   	Dim intRtnSave,intRtnSave2
   	
	with frmThis
		mstrRow = .sprSht.ActiveRow
		
		IF .sprSht_DTL.MAXROWS = 0 THEN
			gErrorMsgBox "저장할 데이타가 없습니다.","저장안내"
			exit Sub
		END IF 
		
		IF mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"FDATE",.sprSht_DTL.ActiveRow) = ""  THEN
			gErrorMsgBox "시작일은 필수입니다.","저장안내"
			exit Sub
		END IF
		
		IF mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EDATE",.sprSht_DTL.ActiveRow) <> ""  THEN
			intRtnSave = gYesNoMsgbox("종료 일을 저장 하시면 해당 데이터는 수정 하실수 없습니다." & vbcrlf & "이대로 진행 하시겠습니까?","처리안내")
			IF intRtnSave <> vbYes then 
				exit Sub
			END IF
		END IF
		
		for intcnt =1 to .sprSht_DTL.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"INPUTFLAG",intcnt) = "INPUT" AND _
				(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"RATE",intcnt)="" or mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"RATE",intcnt)=0 ) then 
			
				intRtnSave2 = gYesNoMsgbox("변경하신 데이터에 수수료가 0인값이 있습니다." & vbcrlf & "이대로 진행 하시겠습니까?","처리안내")
				IF intRtnSave2 <> vbYes then 
					exit Sub
				END IF 
			end if 
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"YEARMON | EXCLIENTCODE | EXCLIENTNAME | SEQ | FDATE | EDATE | RATE | BIGO | CLIENTCODE | BTN | CLIENTNAME | INPUTFLAG")
				
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		'처리 업무객체 호출
		intRtn = mobjMDCOEXSUSURATE.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gErrorMsgBox "저장되었습니다.","저장안내"
			SelectRtn
   		end if
		
   	end with
End Sub


		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="98%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="152" background="../../../images/back_p.gIF"
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
											<td class="TITLE">제작대행사 수수료율 등록</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
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
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<!--<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
												width="90">&nbsp;년월
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 399px"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 96px; HEIGHT: 22px" type="text" maxLength="8"
													size="10" name="txtYEARMON" accessKey="NUM"></TD>
											-->
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="90">제작대행사
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="코드명" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="34" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgEXCLIENTCODE"> <INPUT class="INPUT" id="txtEXCLIENTCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtEXCLIENTCODE"></TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 1040px; HEIGHT: 10px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											
											<TD vAlign="middle" align="right" height="28">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																alt="자료를 한행 추가합니다." src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
								<!--내용 및 그리드-->
							</TR>
							<tr>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
										<TR>
											<td style="WIDTH: 30%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="15928">
														<PARAM NAME="_ExtentY" VALUE="10081">
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
											<td style="WIDTH: 70%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_DTL" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="15928">
														<PARAM NAME="_ExtentY" VALUE="10081">
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
										</TR>
										<TR>
											<TD class="BOTTOMSPLIT" id="lblStatus_CLIENT" style="WIDTH: 1040px"></TD>
											<TD class="BOTTOMSPLIT" id="lblStatus_OUT" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
								</TD>
							</tr>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
