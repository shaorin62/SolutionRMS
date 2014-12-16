<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMACTUALRATEPOP.aspx.vb" Inherits="PD.PDCMACTUALRATEPOP" %>
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt 
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCOACTUALRATE '공통코드, 클래스
Dim mobjPDCOGET
Dim mstrCheck
Const meTab = 9
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
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 삭제버튼
'-----------------------------	
'부서실적 삭제버튼
Sub imgDelete_sprSht_JOBNODEPT_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn_DTL_JOBNODEPT
	gFlowWait meWAIT_OFF
End Sub
'-----------------------------
'행추가
'-----------------------------
sub imgAddRow_sprSht_JOBNODEPT_onclick ()
	With frmThis
		call sprSht_JOBNODEPT_Keydown(meINS_ROW, 0)
	End With 
end sub

Sub sprSht_JOBNODEPT_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
	
		if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
		
		if KeyCode = meCR  Or KeyCode = meTab Then
		Else
			
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_JOBNODEPT, cint(KeyCode), cint(Shift), -1, 1)
			
			Select Case intRtn
				'Case meDEL_ROW: DeleteRtn
			End Select
		End if
	end with
End Sub


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
' UI업무 프로시져 시작  - INIT,,  INITPAGEDATA ...
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 

	dim vntInParam
	dim intNo,i
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjPDCOACTUALRATE	= gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis		
		gSetSheetColor mobjSCGLSpr, .sprSht_JOBNODEPT
		mobjSCGLSpr.SpreadLayout .sprSht_JOBNODEPT, 10, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_JOBNODEPT, "CHK | SEQ | EMPNAME | BTN2 | EMPNO | DEPTNAME | BTN | DEPTCODE | JOBNOSEQ | ACTRATE"
		mobjSCGLSpr.SetHeader .sprSht_JOBNODEPT,        "선택|순번|담당자|담당자사번|담당부서|담당부서코드|JOBSEQ|부서실적입력"
		mobjSCGLSpr.SetColWidth .sprSht_JOBNODEPT, "-1","   4|   5|  10|2|        10|    28|2|          10|6     |12" 
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_JOBNODEPT, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_JOBNODEPT,"..", "BTN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_JOBNODEPT,"..", "BTN2"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_JOBNODEPT, "ACTRATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "EMPNAME | BTN2 | EMPNO | DEPTNAME | BTN | DEPTCODE | JOBNOSEQ",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT, true, "SEQ | JOBNOSEQ"
		.sprSht_JOBNODEPT.style.visibility = "visible"
		mobjSCGLSpr.SetScrollBar .sprSht_JOBNODEPT,2,True,0,-1
		mobjSCGLSpr.colhidden .sprSht_JOBNODEPT, "JOBNOSEQ | SEQ",true
    End With

	'화면 초기값 설정
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'기본값 설정
	
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)
				case 1 : .txtJOBNAME.value = vntInParam(i)
			end select
		next
		
		if .txtJOBNO.value <> "" then
			SelectRtn
		end IF
	end with
End Sub
'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		
		.sprSht_JOBNODEPT.MaxRows = 0
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""
	End with
End Sub

Sub EndPage()
	set mobjPDCOACTUALRATE = Nothing
	set mobjPDCOGET = Nothing
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
		If .txtJOBNO.value = "" Then
			gErrorMsgBox "JOB을 선택 하십시오.","조회안내"
		Else
			strJOBNO = .txtJOBNO.value 
		End If
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		
		strJOBNOSEQ = "1"
		
		vntData = mobjPDCOACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,strJOBNOSEQ)
		
		If not gDoErrorRtn ("SelectRtn_DTL_JOBNODEPT") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_JOBNODEPT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_JOBNODEPT,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht_JOBNODEPT.MaxRows = 0	
			Else
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			End If
		End If		
	END WITH
End Sub

'------------------------------------------
' 저장
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
  	Dim vntData

	'On error resume next
		
		with frmThis
  		'데이터 Validation
  		
		If .txtJOBNO.value = "" Then
			gErrorMsgBOx "우선 JOB 을 조회 하십시오.","저장안내"
			Exit Sub
		End If
		if DataValidation =false then exit sub
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_JOBNODEPT,"CHK|SEQ|EMPNAME|EMPNO|DEPTNAME|DEPTCODE|JOBNOSEQ|ACTRATE")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		intRtn = mobjPDCOACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData,.txtJOBNO.value,"1")
		if not gDoErrorRtn ("ProcessRtn_DTL_JOBNODEPT") then
			'실적분배저장
			gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내"
			SelectRtn 
  		end if
 	end with
End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols,intCnt,intCnt2,intCnt3
   	Dim intColSum
   	Dim strDupACTNAME,strDupACTNAME_CHECK
   	Dim intDupCnt,intSubCnt
   	Dim strTotalRATE
   	Dim strRATE
	'On error resume next
	with frmThis
		'필수항목검사
		for intCnt2 = 1 to .sprSht_JOBNODEPT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTCODE",intCnt2) = "" Then 
				gErrorMsgBox intCnt2 & " 번째 행의 부서코드를 선택 하지 않으셨습니다. 팝업버튼을 눌러 정확한 부서를 선택하시기 바랍니다.","입력오류"
				'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
				Exit Function
			End if
		next
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'부서중복체크
		For intDupCnt = 1 To .sprSht_JOBNODEPT.MaxRows	
			
			strDupACTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTCODE",intDupCnt)

			For intSubCnt = 1 To .sprSht_JOBNODEPT.MaxRows 
				If intDupCnt <> intSubCnt Then 
					strDupACTNAME_CHECK = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTCODE",intSubCnt)
					If strDupACTNAME = strDupACTNAME_CHECK Then
						gErrorMsgBox "담당자" & intDupCnt & "행에 중복된 부서가 있습니다.","중복안내"
						mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,intSubCnt
						'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
						Exit Function
					End IF
				End IF
			Next
		Next
		
		'실적비율 합계 100% 처리...
		strTotalRATE = 0
		For intCnt3 = 1 To .sprSht_JOBNODEPT.MaxRows 
			strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"ACTRATE",intCnt3)
			
			If strRATE = 0 Then
				gErrorMsgBox "실적분배비율을 반드시 입력하여야 합니다.","저장안내"
				'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
				Exit Function
			End IF
			
			strTotalRATE= strTotalRATE+strRATE
		Next
		
		'저장한 내역을 처음부터 끝까지 돌면서 합이 100인지 확인
		If strTotalRATE <> 100 Then 
			gErrorMsgBox "실적분배비율의 총합은 100 이어야 합니다.","저장안내"
			'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
			Exit Function
		End If
   	End with
	DataValidation = true
End Function

'------------------------------------------
' 삭제
'------------------------------------------
Sub DeleteRtn_DTL_JOBNODEPT
	Dim vntData
	Dim vntDataDEL
	Dim intSelCnt, intRtn, i , intCnt
	Dim strYEARMON
	Dim strSEQ
	Dim strRATE , strTotalRATE
	Dim strRow
	Dim strJOBNO
	Dim strJOBNOSEQ
	Dim intValCnt
	Dim strCHKCnt
	Dim intDelCount
	Dim intColSum
	
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht_JOBNODEPT,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		strCHKCnt = 0
		
		
		'삭제할 내역을뺀 나머지를 돌면서 실적분배비율의 합이 100인지 확인
		
		For intCnt = 1 To .sprSht_JOBNODEPT.MaxRows
			If mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"CHK",intCnt) <> "1" And mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"JOBNOSEQ",intCnt) <> "" Then
				strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"ACTRATE",intCnt)
				strTotalRATE= strTotalRATE+strRATE
			Else 
				strCHKCnt = 1
			End if
		Next
		
		'저장시 체크된것만 저장될수 있도록
		intColSum = 0
  		for intCnt = 1 to .sprSht_JOBNODEPT.MaxRows
				if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"CHK",intCnt) = 1  Then 
						intColSum = intColSum + 1
				End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "선택된 데이터가 없습니다.","삭제안내"
			exit Sub
		End If
		
		
		If strTotalRATE <> 100 Then
			if strTotalRATE <> 0 Then
				gErrorMsgBox "삭제할데이터를 제외한 나머지 실적분배비율을 반드시 100% 또는 0%(전체삭제) 이여야 합니다.","삭제안내"
				Exit Sub
			End If
		End IF
			
		
		intRtn = gYesNoMsgbox("선택된 자료가 삭제 됩니다." & vbcrlf & "자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		strJOBNO = .txtJOBNO.value 
		'선택된 자료를 끝에서 부터 삭제
		intDelCount = 0
		for i = .sprSht_JOBNODEPT.MaxRows to 1 step -1	
			if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"CHK",i) = 1 THEN
				If mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"SEQ",i) <> "" Then 
					strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"SEQ",i)
					strJOBNOSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"JOBNOSEQ",i) 
					intRtn = mobjPDCOACTUALRATE.DeleteRtn_DTL_JOBNODEPT_JOBNO(gstrConfigXml, strSEQ,strJOBNO,strJOBNOSEQ)
					
				End If
				IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,i
   				End IF
   				intDelCount = intDelCount + 1
   				gWriteText "",intDelCount & " 건이 삭제" & mePROC_DONE
   			END IF
		next
		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_JOBNODEPT
		'변경된 내역 재저장
		ProcessRtn_DTL_ACTUALRATE_DEL
		SelectRtn
	End with
	err.clear
End Sub

Sub ProcessRtn_DTL_ACTUALRATE_DEL()
    Dim intRtn
  	Dim vntData
	Dim strJOBNO ,strSEQFlag , strRATE ,strTotalRATE, strJOBNOSEQ
	Dim strRow , strOLDSEQ
	Dim intCnt , intCode ,intEDITCODE 
	Dim dblACTLRATE
	
	with frmThis
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht_JOBNODEPT,"CHK|SEQ|EMPNAME|EMPNO|DEPTNAME|DEPTCODE|JOBNOSEQ|ACTRATE")
		if  not IsArray(vntData) then 
			exit sub
		End If
		intRtn = mobjPDCOACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData,.txtJOBNO.value,"1")
		'건수저장 
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
  		end if
  		
 	end with
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
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
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
'****************************************************************************************
' SHEET관련 시작
'****************************************************************************************
Sub sprSht_JOBNODEPT_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_JOBNODEPT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_JOBNODEPT.MaxRows
				sprSht_JOBNODEPT_Change 1, intcnt
			next
		end if
	end with
End Sub

sub sprSht_JOBNODEPT_DblClick (ByVal Col, ByVal Row)
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_JOBNODEPT, ""
		end if
	end with
end sub

Sub sprSht_JOBNODEPT_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim strDeptCodeName
	with frmThis
	
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"DEPTNAME") Then
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTNAME",Row)
				
			If strCode = "" AND strCodeName <> "" Then		
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
					
				vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						
						mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT,Col,.sprSht_JOBNODEPT.ActiveRow			
						.txtJOBNAME.focus
						.sprSht_JOBNODEPT.focus
						If Row <> .sprSht_JOBNODEPT.MaxRows Then
							mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row -1
						Else
							mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row
						End IF
					Else
						mobjSCGLSpr_ClickProc "sprSht_JOBNODEPT", Col, .sprSht_JOBNODEPT.ActiveRow
						.txtJOBNAME.focus
						.sprSht_JOBNODEPT.focus 
						If Row <> .sprSht_JOBNODEPT.MaxRows Then
							mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row -1
						Else
							mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row
						End IF
					End If
   				End If
   			End If
	   			
	   	Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"EMPNAME") Then
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTNAME",.sprSht_JOBNODEPT.ActiveRow)
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"EMPNAME",.sprSht_JOBNODEPT.ActiveRow)
				
				vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntData(3,1)
					mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT,Col,frmThis.sprSht_JOBNODEPT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_JOBNODEPT", Col, .sprSht_JOBNODEPT.ActiveRow
				End If
				.txtJOBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다 이거수
				.sprSht_JOBNODEPT.Focus	
				If Row <> .sprSht_JOBNODEPT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+7, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+7, Row
				End IF
			
	   	Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"ACTRATE") Then
			'mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"CHK",Row, "1"	   			
   		End If
	End with
	mobjSCGLSpr.CellChanged frmThis.sprSht_JOBNODEPT, Col, Row
End Sub

'--------------------------------------------------
'쉬트 팝업버튼클릭후 프로세스
'--------------------------------------------------
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	Dim vntRet, vntInParams
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"DEPTNAME") Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
			End IF
			
			.txtJOBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_JOBNODEPT.Focus	
			If Row <> .sprSht_JOBNODEPT.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row
			End If
			
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"EMPNAME") Then
			
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"EMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
			End IF
			
			.txtJOBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_JOBNODEPT.Focus	
			If Row <> .sprSht_JOBNODEPT.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2, Row
			End If
		End if
		
	End With
End Sub

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_JOBNODEPT_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
		IF Col = 7 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN") then exit Sub
		
			vntInParams = array("","","")
			vntRet = gShowModalWindow("../PDCO/PDCMDEPTPOP.aspx",vntInParams , 413,440)

				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
				.txtJOBNAME.focus()
				.sprSht_JOBNODEPT.focus 
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		
		ElseIf Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN2") then exit Sub
		
			vntInParams = array("","","","") '<< 받아오는경우
			vntRet = gShowModalWindow("../PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)

					
			IF isArray(vntRet) then

				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(2,0)			
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row						
				.txtJOBNAME.focus()
				.sprSht_JOBNODEPT.focus 
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		end if
		.sprSht_JOBNODEPT.focus 
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
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" width="400" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
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
												<td class="TITLE">실적분배율 입력</td>
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
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="800" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="40%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 5px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="top">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 100px; CURSOR: hand" onclick="vbscript:Call DateClean()">의뢰일&nbsp;검색</TD>
												<TD class="SEARCHDATA" style="WIDTH: 225px"><INPUT class="INPUT" id="txtFROM" title="의뢰일 검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
														accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM">&nbsp;<IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0"
														name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="의뢰일 검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
														align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"><FONT face="굴림">Job 
														No</FONT></TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtJOBNAME" title="코드명" style="WIDTH: 208px; HEIGHT: 22px" type="text"
														maxLength="255" align="left" size="29" name="txtJOBNAME"></FONT> <IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgJOBNO">
													<INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 56px; HEIGHT: 22px" accessKey=",M"
														type="text" maxLength="7" size="4" name="txtJOBNO"></TD>
												<td class="SEARCHDATA"></td>
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
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left">
																<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
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
															<td class="TITLE">담당부서/담당자</td>
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
																	height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" border="0" name="imgSave_sprSht_JOBNODEPT"></TD>
															<td><IMG id="imgDelete_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																	height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete_sprSht_JOBNODEPT"></td>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
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
						<TD style="WIDTH: 100%; HEIGHT: 400px" vAlign="top" align="center" colSpan="2">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht_JOBNODEPT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31829">
									<PARAM NAME="_ExtentY" VALUE="10583">
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
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
