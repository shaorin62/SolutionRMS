<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMACTUALRATE.aspx.vb" Inherits="PD.PDCMACTUALRATE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>실적분배율 입력</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/제작관리번호 등록 화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBNO.aspx
'기      능 : 제작관리번호 C/D/U/R
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/19 By Kim Tae Ho
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjTRLNREG, mobjPDCMACTUALRATE '공통코드, 클래스
Dim mobjPDCMGET
Dim mstrCheck
Dim mstrFlag
Dim mstrBindCHK
Const meTab = 9
mstrFlag = "New"
Dim mstrNoClick 
mstrNoClick = False
mstrBindCHK = False
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub


'=========================================================================================
' 버튼 이벤트 시작
'=========================================================================================
'-----------------------------------
' 조회 버튼
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
' 저장 버튼 
'-----------------------------------
Sub imgSave_sprSht_JOBNODEPT_onclick ()
	IF frmThis.sprSht_JOBNODEPT.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn_DTL_JOBNODEPT
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_sprSht_ACTUALRATE_onclick ()
	IF frmThis.sprSht_ACTUALRATE.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn_DTL_ACTUALRATE(0)
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------
' 삭제버튼
'-----------------------------	
Sub imgDelete_sprSht_JOBNODEPT_onclick()
	with frmThis 
	
	End with 
	gFlowWait meWAIT_ON
	DeleteRtn_DTL_JOBNODEPT
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_sprSht_ACTUALRATE_onclick()
	with frmThis 
	
	End with 
	gFlowWait meWAIT_ON
	DeleteRtn_DTL_ACTUALRATE
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'행추가
'-----------------------------
sub imgAddRow_sprSht_JOBNODEPT_onclick ()
	Dim strRow,strJOBNO
	With frmThis
		strRow = .sprSht_JOBNO.ActiveRow

		If strRow > 0  Then
			call sprSht_JOBNODEPT_Keydown(meINS_ROW, 0)
		Else
			msgbox "조회후 추가 가능합니다."
			exit sub
		End If
	End With 
end sub

sub imgAddRow_sprSht_ACTUALRATE_onclick ()
	Dim strRow,strJOBNO
	With frmThis
		strRow = .sprSht_JOBNO.ActiveRow
		If strRow > 0  Then
			call sprSht_ACTUALRATE_Keydown(meINS_ROW, 0)
		Else
			msgbox "조회후 추가 가능합니다."
			exit sub
		End If
		
	End With 
end sub


'-----------------------------
' 달력
'-----------------------------	
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

'-----------------------------
' 엑셀버튼
'-----------------------------	
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		'mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'=========================================================================================
' 버튼관련 끝
'=========================================================================================




'=========================================================================================
' UI업무 프로시져 시작  - INIT INITPAGEDATA ...
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjPDCMACTUALRATE = gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
   
		gSetSheetColor mobjSCGLSpr, .sprSht_JOBNO
		mobjSCGLSpr.SpreadLayout .sprSht_JOBNO, 11, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_JOBNO, "PROJECTNO|JOBNO|SEQ|JOBNAME|CLIENTNAME|CLIENTSUBNAME|SUBSEQNAME|JOBGUBNNAME|CREPARTNAME|REQDAY|ENDFLAGNAME"
		mobjSCGLSpr.SetHeader .sprSht_JOBNO,        "프로젝트|JOBNO|SEQ|JOB명|광고주|사업부|브랜드|매체부문|매체분류|의뢰일|완료구분"
		mobjSCGLSpr.SetColWidth .sprSht_JOBNO, "-1","        10|  10|4|   25|    15|    15|    15|      15|      10|     10|     10"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNO, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNO, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_JOBNO, "REQDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNO, true, "PROJECTNO|JOBNAME|SEQ|JOBNO|CLIENTNAME|CLIENTSUBNAME|SUBSEQNAME|JOBGUBNNAME|CREPARTNAME|REQDAY|ENDFLAGNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNO, "SEQ|JOBNAME|CLIENTNAME|CLIENTSUBNAME|SUBSEQNAME|JOBGUBNNAME|CREPARTNAME|REQDAY",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNO, "PROJECTNO|JOBNO|ENDFLAGNAME",-1,-1,2,2,false '가운데
		.sprSht_JOBNO.style.visibility = "visible"
		
		

		
		gSetSheetColor mobjSCGLSpr, .sprSht_JOBNODEPT
		mobjSCGLSpr.SpreadLayout .sprSht_JOBNODEPT, 8, 0, 3, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 2, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_JOBNODEPT, "SEQ|EMPNAME|BTN2|EMPNO|DEPTNAME|BTN|DEPTCODE|JOBNOSEQ"
		mobjSCGLSpr.SetHeader .sprSht_JOBNODEPT,        "순번|담당자|담당자사번|담당부서|담당부서코드|JOBSEQ"
		mobjSCGLSpr.SetColWidth .sprSht_JOBNODEPT, "-1","  5|       10|2|     10|    18|2|    10|6" 
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_JOBNODEPT,"..", "BTN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_JOBNODEPT,"..", "BTN2"
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "SEQ|EMPNAME|BTN2|EMPNO|DEPTNAME|BTN|DEPTCODE|JOBNOSEQ",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT, true, "SEQ|JOBNOSEQ"
		.sprSht_JOBNODEPT.style.visibility = "visible"
		
		
		
		gSetSheetColor mobjSCGLSpr, .sprSht_ACTUALRATE
		mobjSCGLSpr.SpreadLayout .sprSht_ACTUALRATE, 6, 0, 3, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_ACTUALRATE, 2, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_ACTUALRATE, "SEQ|ACTDEPTNAME|BTN|ACTDEPTCD|ACTRATE|JOBNOSEQ"
		mobjSCGLSpr.SetHeader .sprSht_ACTUALRATE,        "순번|실적부서|실적부서코드|실적분배비율|JOBSEQ"
		mobjSCGLSpr.SetColWidth .sprSht_ACTUALRATE, "-1","   5|       18|2|       10|         10|6"
		mobjSCGLSpr.SetRowHeight .sprSht_ACTUALRATE, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_ACTUALRATE, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_ACTUALRATE,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_ACTUALRATE, "ACTRATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht_ACTUALRATE, "SEQ|ACTDEPTNAME|BTN|ACTDEPTCD|JOBNOSEQ",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_ACTUALRATE, "",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_ACTUALRATE, true, "SEQ|JOBNOSEQ"
		.sprSht_ACTUALRATE.style.visibility = "visible"
	
    End With

	'화면 초기값 설정
	InitPageData
End Sub
'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.sprSht_JOBNO.MaxRows = 0
		.sprSht_JOBNODEPT.MaxRows = 0
		.sprSht_ACTUALRATE.MaxRows = 0
		
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""
	End with
End Sub

Sub EndPage()
	set mobjPDCMACTUALRATE = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
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

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		
   	
   	End with
	DataValidation = true
End Function
'=========================================================================================
' UI업무 프로시져 끝
'=========================================================================================





'****************************************************************************************
' UI 시작 - 조회 저장 수정 삭제  
'****************************************************************************************
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	mstrNoClick = False
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols , intCnt
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_JOBNO.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		vntData = mobjPDCMACTUALRATE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNO.value),Trim(.txtJOBNAME.value))
		
		
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_JOBNO,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_JOBNO,meCLS_FLAG
			
			For intCnt = 1 To .sprSht_JOBNO.MaxRows '조회된 내역을 처음부터 끝까지 돌면서
				If mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"SEQ",intCnt) = 1 Then  '특정값에 해당 되면 파란색
					mobjSCGLSpr.SetCellShadow .sprSht_JOBNO, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False '이게 파란색
				Else
					mobjSCGLSpr.SetCellShadow .sprSht_JOBNO, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False '이게 흰색
				End If
			Next
			
			
			If mlngRowCnt < 1 Then
			.sprSht_JOBNO.MaxRows = 0	
			End If
			
			gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			
			if mlngRowCnt <> 0 then
				Call sprSht_JOBNO_Click(1,1)
			end if
		End If		

	END WITH
	
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
	
End Sub


Sub SelectRtn_DTL_JOBNODEPT ()
	Dim vntData
   	Dim strRow,strJOBNO , strJOBNOSEQ
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_JOBNODEPT.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'시트1의 JOB번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNOSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"SEQ",strRow)
		
		vntData = mobjPDCMACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,strJOBNOSEQ)
		
		
		If not gDoErrorRtn ("SelectRtn_DTL_JOBNODEPT") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_JOBNODEPT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_JOBNODEPT,meCLS_FLAG
			
			If mlngRowCnt < 1 Then
			.sprSht_JOBNODEPT.MaxRows = 0	
			End If
			
			'gWriteText lblstatus_JOBNODEPT, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
		
		End If		

	END WITH
	
	'조회완료메세지
	'gWriteText "", "자료가 검색" & mePROC_DONE
End Sub



Sub SelectRtn_DTL_ACTUALRATE()
	Dim vntData
   	Dim strRow,strJOBNO , strJOBNOSEQ
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_ACTUALRATE.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'시트1의 JOB번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNOSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"SEQ",strRow)
		
		vntData = mobjPDCMACTUALRATE.SelectRtn_DTL_ACTUALRATE(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,strJOBNOSEQ)
		
		
		If not gDoErrorRtn ("SelectRtn_DTL_ACTUALRATE") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_ACTUALRATE,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_ACTUALRATE,meCLS_FLAG
			
			If mlngRowCnt < 1 Then
			.sprSht_ACTUALRATE.MaxRows = 0	
			End If
			
			'gWriteText lblstatus_JOBNODEPT, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
		
		End If		

	END WITH
	
	'조회완료메세지
	'gWriteText "", "자료가 검색" & mePROC_DONE
End Sub






'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn_DTL_JOBNODEPT ()
    Dim intRtn
  	Dim vntData
	Dim strJOBNO ,strSEQFlag , strJOBNOSEQ
	Dim strRow ,strMaxRow , strOLDSEQ
	Dim intCnt, intSubCnt , intCode , intEDITCODE
	Dim strEMPNAME  , strEMPNAME_CHECK
	with frmThis

  		vntData = mobjSCGLSpr.GetDataRows(.sprSht_JOBNODEPT,"SEQ|EMPNAME|BTN2|EMPNO|DEPTNAME|BTN|DEPTCODE|JOBNOSEQ")
  		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		If .sprSht_JOBNODEPT.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
		
		
		'하나라도 무조건 입력
		strRow = .sprSht_JOBNODEPT.ActiveRow
		strEMPNAME = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"EMPNAME",strRow)
		If strEMPNAME = "" THEN 
			gErrorMsgBox "공백 내역은 저장할수 없습니다.","저장안내"
			mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,strRow
			Exit Sub
		End IF
		
	
		'시트1의 JOB번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNOSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"SEQ",strRow)
		

		'시트2의 SEQ유무로 NEW인지 UPDATE인지에 사용
		strRow = .sprSht_JOBNODEPT.ActiveRow
		strOLDSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"SEQ",strRow)
		
		
		
		'조회된 내역을 처음부터 끝까지 돌면서 중복 사원 체크
		strMaxRow = .sprSht_JOBNODEPT.MaxRows	
		For intCnt = 1 To strMaxRow 
			strEMPNAME = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"EMPNAME",intCnt)
			
			
			For intSubCnt = 1 To strMaxRow
				If intCnt <> intSubCnt Then 
					strEMPNAME_CHECK = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"EMPNAME",intSubCnt)
					If strEMPNAME = strEMPNAME_CHECK Then
						gErrorMsgBox "담당자" & intCnt & "라인에 중복된 사원명이 있습니다.","중복안내"
						mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,intSubCnt
						Exit Sub
					End IF
				End IF
			Next
		Next
		
		
		
		
		'시트2에서 SEQ에따른 NEW ,  UPDATE
		if strOLDSEQ = "" then
			strSEQFlag = "new"
			intRtn = mobjPDCMACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData, strSEQFlag,strJOBNO,strOLDSEQ,strJOBNOSEQ)
		else
			strSEQFlag = "update"
			intRtn = mobjPDCMACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData, strSEQFlag,strJOBNO,strOLDSEQ,"")
		end if
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
			if strSEQFlag = "new" then
				gErrorMsgBox " 자료가 신규저장" & mePROC_DONE,"저장안내"
				SelectRtn
				For intCnt = 1 To .sprSht_JOBNODEPT.MaxRows 
					intCode = intCnt 
				Next
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, 1,intCode
			else
				gErrorMsgBox " 자료가" & intRtn & " 건 수정저장" & mePROC_DONE,"저장안내" 
				SelectRtn
				For intCnt = 1 To .sprSht_JOBNODEPT.MaxRows 
					intEDITCODE = intCnt 
				Next
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, 1,intEDITCODE
			end if
			
  		end if
 	end with
End Sub


Sub ProcessRtn_DTL_ACTUALRATE(strDeleteYN)
    Dim intRtn
  	Dim vntData
	Dim strJOBNO ,strSEQFlag , strRATE ,strTotalRATE, strJOBNOSEQ
	Dim strRow , strOLDSEQ
	Dim intCnt , intCode ,intEDITCODE 
	Dim strACTDEPTNAME
	Dim dblACTLRATE
	with frmThis

  		vntData = mobjSCGLSpr.GetDataRows(.sprSht_ACTUALRATE,"SEQ|ACTDEPTNAME|BTN|ACTDEPTCD|ACTRATE|JOBNOSEQ")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		If .sprSht_ACTUALRATE.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
		
		'하나라도 무조건 입력
		strRow = .sprSht_ACTUALRATE.ActiveRow
		strACTDEPTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_ACTUALRATE,"ACTDEPTNAME",strRow)
		If strACTDEPTNAME = "" THEN 
			gErrorMsgBox "공백 내역은 저장할수 없습니다.","저장안내"
			mobjSCGLSpr.DeleteRow .sprSht_ACTUALRATE,strRow
			
			Exit Sub
		End IF
		
		
		strRow = .sprSht_ACTUALRATE.ActiveRow
		strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_ACTUALRATE,"ACTRATE",strRow)
		
		
		
		
		If strDeleteYN = 0 Then
			strTotalRATE = 0
			'저장한 내역을 처음부터 끝까지 돌면서 입력하지 않은 실적분배비율이 있는지 확인
			'입력한 실적분배비율들의 합구함
			For intCnt = 1 To .sprSht_ACTUALRATE.MaxRows 
				strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_ACTUALRATE,"ACTRATE",intCnt)
				
				If strRATE = 0 Then
					gErrorMsgBox "실적분배비율을 반드시 입력하여야 합니다.","저장안내"
					Exit Sub
				End IF
				
				strTotalRATE= strTotalRATE+strRATE
			Next
			
			'저장한 내역을 처음부터 끝까지 돌면서 합이 100인지 확인
			If strTotalRATE <> 100 Then 
				gErrorMsgBox "실적분배비율의 총합은 100 이어야 합니다.","저장안내"
				Exit Sub
			End If
		End If
		
	
		'시트1의 JOB번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNOSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"SEQ",strRow)
		strRow = .sprSht_ACTUALRATE.MaxRows

		
		'시트3의 SEQ유무로 NEW인지 UPDATE인지에 사용
		strRow = .sprSht_ACTUALRATE.ActiveRow
		strOLDSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_ACTUALRATE,"SEQ",intCnt)
		

		
		
		'추가 저장및 수정저장
		intRtn = mobjPDCMACTUALRATE.ProcessRtn_DTL_ACTUALRATE(gstrConfigXml,vntData,strJOBNO,strJOBNOSEQ)


		'건수저장 
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_ACTUALRATE,meCLS_FLAG
			
			gErrorMsgBox " 자료가" & intRtn & " 건 저장 " & mePROC_DONE,"저장안내" 
			SelectRtn
			For intCnt = 1 To .sprSht_ACTUALRATE.MaxRows 
				intEDITCODE = intCnt 
			Next
			mobjSCGLSpr.ActiveCell .sprSht_ACTUALRATE, 1,intEDITCODE
			
  		end if
  		
 	end with
End Sub




'------------------------------------------
' JOBNODEPT 삭제
'------------------------------------------
Sub DeleteRtn_DTL_JOBNODEPT ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ

	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht_JOBNODEPT,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			
			IF mobjSCGLSpr.GetFlagMode(.sprSht_JOBNODEPT,vntData(i)) <> meINS_TRANS then
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"SEQ",vntData(i)))
				intRtn = mobjPDCMACTUALRATE.DeleteRtn_DTL_JOBNODEPT(gstrConfigXml, strSEQ)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				gWriteText "", "[" & strSEQ & "] 자료 " & intRtn & "건이 삭제" & mePROC_DONE
   			End IF
		next
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_JOBNODEPT
		SelectRtn
	End with
	err.clear
End Sub


'------------------------------------------
' ACTUALRATE 삭제
'------------------------------------------
Sub DeleteRtn_DTL_ACTUALRATE ()
	Dim vntData
	Dim intSelCnt, intRtn, i , intCnt
	dim strYEARMON
	Dim strSEQ
	Dim strRATE , strTotalRATE

	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht_ACTUALRATE,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		'삭제할 내역을뺀 나머지를 돌면서 실적분배비율의 합이 100인지 확인
		For intCnt = 1 To .sprSht_ACTUALRATE.MaxRows
			If intCnt <> .sprSht_ACTUALRATE.ActiveRow Then
				strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_ACTUALRATE,"ACTRATE",intCnt)
				strTotalRATE= strTotalRATE+strRATE
			End if
		Next
		
		If strTotalRATE <> 100 Then
			gErrorMsgBox "삭제할데이터를 제외한 나머지 실적분배비율을 반드시 100으로 만들어야 합니다.","저장안내"
			Exit Sub
		End IF
			
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			IF mobjSCGLSpr.GetFlagMode(.sprSht_ACTUALRATE,vntData(i)) <> meINS_TRANS then
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht_ACTUALRATE,"SEQ",vntData(i)))
				intRtn = mobjPDCMACTUALRATE.DeleteRtn_DTL_ACTUALRATE(gstrConfigXml, strSEQ)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				gWriteText "", "[" & strSEQ & "] 자료 " & intRtn & "건이 삭제" & mePROC_DONE
   			End IF
		next
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_ACTUALRATE
		ProcessRtn_DTL_ACTUALRATE(1)
		SelectRtn
	End with
	err.clear
End Sub

'****************************************************************************************
' ui관련 프로세스 끝
'****************************************************************************************



'****************************************************************************************
' 버튼팝업 시작
'****************************************************************************************
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
		SELECTRTN
	end if
	
End Sub



'-----------------------------------------------------------------------------------------
' 담당부서 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgDEPTCD_onclick
	Call DEPT_POP()
End Sub

Sub DEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC 코드/명,optional(현재사용여부,사용검사일,추가조회 필드,Key Like여부)
		vntInParams = array(trim(.txtDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = trim(vntRet(0,0))	'Code값 저장
			.txtDEPTNAME.value = trim(vntRet(1,0))	'코드명 표시
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtEMPNAME.focus()
			gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub


'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTNAME.value)
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = trim(vntData(0,0))
					.txtDEPTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtEMPNAME.focus()
				Else
					Call DEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
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
		vntInParams = array(trim(.txtDEPTCD.value), trim(.txtDEPTNAME.value), trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtDEPTCD.value = trim(vntRet(2,0))  ' Code값 저장
			.txtDEPTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			
			if .sprSht.ActiveRow >0 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",.sprSht.ActiveRow, .txtEMPNO.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",.sprSht.ActiveRow, .txtEMPNAME.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtMEMO.focus()					' 포커스 이동
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
			gSetChangeFlag .txtDEPTCD
			gSetChangeFlag .txtDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub sprSht_enter_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A",.txtDEPTCD.value,.txtDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					.txtDEPTCD.value = trim(vntData(2,1))
					.txtDEPTNAME.value = trim(vntData(3,1))
					
					if .sprSht.ActiveRow >0 Then
			
						mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",.sprSht.ActiveRow, .txtEMPNO.value
						mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",.sprSht.ActiveRow, .txtEMPNAME.value
						
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtMEMO.focus()
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
'****************************************************************************************
' 버튼팝업 끝
'****************************************************************************************







'****************************************************************************************
' SHEET관련 시작
'****************************************************************************************
Sub sprSht_JOBNO_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_JOBNO, Col, Row
End Sub

Sub sprSht_JOBNODEPT_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	msgbox 11
	mobjSCGLSpr.CellChanged frmThis.sprSht_JOBNODEPT, Col, Row
End Sub

Sub sprSht_ACTUALRATE_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_ACTUALRATE, Col, Row
End Sub



Sub sprSht_JOBNO_Click(ByVal Col, ByVal Row)
	Dim strRow,strJOBNO
	with frmThis
		strRow = .sprSht_JOBNO.ActiveRow
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNAME",strRow)
		.txtJOBNO.value = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNO,"JOBNO",strRow)
		If strJOBNO <> "" then 
			SelectRtn_DTL_JOBNODEPT
			SelectRtn_DTL_ACTUALRATE
		End IF	
	end with
End Sub


'--------------------------------------------------
'추가버튼  키다운
'--------------------------------------------------
Sub sprSht_JOBNODEPT_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
	
		if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
		
		if KeyCode = meCR  Or KeyCode = meTab Then
		Else
			
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_JOBNODEPT, cint(KeyCode), cint(Shift), -1, 1)
			strRow = .sprSht_JOBNODEPT.ActiveRow
			
			mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT,false,"EMPNAME|BTN2|EMPNO|DEPTNAME|BTN|DEPTCODE",1,strRow,false
			strRow = strRow-1
			mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT,true,"SEQ|EMPNAME|BTN2|EMPNO|DEPTNAME|BTN|DEPTCODE|JOBNOSEQ",1,strRow,false
			strRow = strRow+1
			mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT,true,"SEQ|JOBNOSEQ",1,strRow,false
			Select Case intRtn
				'Case meDEL_ROW: DeleteRtn
			End Select
		End if
	end with
End Sub

Sub sprSht_ACTUALRATE_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
	
		if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
		
		if KeyCode = meCR  Or KeyCode = meTab Then
		Else
			
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_ACTUALRATE, cint(KeyCode), cint(Shift), -1, 1)
			strRow = .sprSht_ACTUALRATE.ActiveRow
			
			mobjSCGLSpr.SetCellsLock2 .sprSht_ACTUALRATE,false,"ACTDEPTNAME|ACTDEPTCD|ACTRATE",1,strRow,false
			strRow = strRow-1
			mobjSCGLSpr.SetCellsLock2 .sprSht_ACTUALRATE,true,"SEQ|ACTDEPTNAME|ACTDEPTCD|JOBNOSEQ",1,strRow,false
			strRow = strRow+1
			mobjSCGLSpr.SetCellsLock2 .sprSht_ACTUALRATE,true,"SEQ|JOBNOSEQ",1,strRow,false
			Select Case intRtn
				'Case meDEL_ROW: DeleteRtn
			End Select
		End if
	end with
End Sub


	


'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_JOBNODEPT_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
		
			
		IF Col = 6 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN") then exit Sub
		
			vntInParams = array("","","")
			vntRet = gShowModalWindow("../PDCO/PDCMDEPTPOP.aspx",vntInParams , 413,440)

				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
				
				.sprSht_JOBNODEPT.focus 
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		
		ElseIf Col = 3 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN2") then exit Sub
		
			vntInParams = array("","","","") '<< 받아오는경우
			vntRet = gShowModalWindow("../PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)

					
			IF isArray(vntRet) then

				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(2,0)			
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row						

				.sprSht_JOBNODEPT.focus 
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		end if
		.sprSht_JOBNODEPT.focus 
		
	End with
	
End Sub


Sub sprSht_ACTUALRATE_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
			
		IF Col = 3 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_ACTUALRATE,"BTN") then exit Sub
		
			vntInParams = array("","","")
			vntRet = gShowModalWindow("../PDCO/PDCMDEPTPOP.aspx",vntInParams , 413,440)

				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_ACTUALRATE,"ACTDEPTCD",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_ACTUALRATE,"ACTDEPTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht_ACTUALRATE, Col,Row
				
				.sprSht_ACTUALRATE.focus 
				mobjSCGLSpr.ActiveCell .sprSht_ACTUALRATE, Col+2,Row
			End IF
		
		end if
		.sprSht_ACTUALRATE.focus 
	End with
	
End Sub


'****************************************************************************************
' SHEET관련 끝
'****************************************************************************************
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML><XML id="xmlBind1"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="80%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD colSpan="2">
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
												<td class="TITLE">&nbsp;실적분배율&nbsp;입력</td>
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
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="40%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 17px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="1040" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand" onclick="vbscript:Call DateClean()"
													width="112">의뢰일&nbsp;검색</TD>
												<TD class="SEARCHDATA" style="WIDTH: 300px" width="300"><INPUT class="INPUT" id="txtFROM" title="의뢰일 검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
														accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="의뢰일 검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
													width="90"><FONT face="굴림">Job No</FONT></TD>
												<TD class="SEARCHDATA" style="WIDTH: 300px" width="300"><INPUT class="INPUT_L" id="txtJOBNAME" title="코드명" style="WIDTH: 160px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="21" name="txtJOBNAME"></FONT><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 56px; HEIGHT: 22px" accessKey=",M"
														type="text" maxLength="6" size="4" name="txtJOBNO"></TD>
												<td class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<tr>
						<td colSpan="2">
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="125" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="15" style="HEIGHT: 15px"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;JOB 내역</td>
											</tr>
										</table>
									</TD>
									<!--job 내역 버튼있던자리--></TR>
							</TABLE>
						</td>
					</tr>
					<!--Input End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 90%" vAlign="top" align="center" colSpan="2">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 98%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht_JOBNO" style="WIDTH: 100%; HEIGHT: 98%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="5450">
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
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%" colSpan="2"></TD>
					</TR>
					<!--2개로 나눠지는 부분-->
					<tr>
						<!--첫번째-->
						<td>
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="125" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;담당부서/담당자</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgAddRow_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
														alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
												<TD><IMG id="imgSave_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'"
														height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<td><IMG id="imgDelete_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
						</td>
						<!--두번째-->
						<td>
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="125" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;실적부서/분배율</td>
											</tr>
										</table>
									</TD>
									
									<TD style="WIDTH: 340px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton3" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgAddRow_sprSht_ACTUALRATE" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
														alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
												<TD><IMG id="imgSave_sprSht_ACTUALRATE" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'"
														height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<td><IMG id="imgDelete_sprSht_ACTUALRATE" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
						</td>
					</tr>
					<!--아랫tr 그리드 2개들어갈곳-->
					<tr height="100">
						<td>
							<table class="DATA" height="200%" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<!--첫번째-->
									<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center" colSpan="2">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_JOBNODEPT" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="13758">
												<PARAM NAME="_ExtentY" VALUE="5027">
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
							</table>
						</td>
						<td>
							<table class="DATA" height="200%" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<!--두번째-->
									<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center" colSpan="2">
										<DIV id="pnlTab3" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_ACTUALRATE" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="13758">
												<PARAM NAME="_ExtentY" VALUE="5027">
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
							</table>
						</td>
					</tr>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
