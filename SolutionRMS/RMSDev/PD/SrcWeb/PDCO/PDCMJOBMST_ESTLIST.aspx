<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_ESTLIST.aspx.vb" Inherits="PD.PDCMJOBMST_ESTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_ESTDTL.aspx
'기      능 : JOBMST의 두번째 탭 - 가/본 견적서를 저장 및 수정 한다. 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
		
option explicit

Dim mlngRowCnt, mlngColCnt		
Dim mobjPDCOPREESTLIST
Dim mobjPDCOGET
Dim mobjSCCOGET

Const meTab = 9

'=============================
' 이벤트 프로시져 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'조회
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub		

'엑셀버튼
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'삭제
Sub imgDelete_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF	
End Sub

'내역복사
Sub imgListcopy_onclick
	Dim vntData
	Dim i
	Dim strPREESTNO
	Dim strPREESTNAME
	Dim strJOBNO, strJOBNAME
	Dim intRtn
	Dim strNEWPREESTNO
	Dim intSaveRtn
	Dim intCnt
	Dim intEDITCODE
	Dim intCount
	
	strNEWPREESTNO = ""
	gFlowWait meWAIT_ON
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "조회된 데이터가 없습니다.","내역복사안내"
			Exit Sub
		end if
		
		'체크가 된 데이터가 있는지 없는지 체크한다.
		intCount = 0
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strPREESTNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO", i)
				strPREESTNAME	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNAME", i)
				
				intCount = intCount + 1
			end if
		next
		
		'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
		if intCount = 0 then
			gErrorMsgBox "복사할 데이터를 선택하십시오.","내역복사안내"
			Exit Sub
		elseif intCount > 1 then
			gErrorMsgBox "복사할데이터는 한행만 선택하십시오.", "내역복사안내"
			Exit Sub
		end if
		
		intRtn = gYesNoMsgbox( strPREESTNAME & " 의 내역을 복사 하시겠습니까?","내역복사 확인")
		
		IF intRtn <> vbYes then exit Sub
		
		strJOBNO   = parent.document.forms("frmThis").txtJOBNO.value
		strJOBNAME = parent.document.forms("frmThis").txtPRIJOBNAME.value 
		
		intSaveRtn = mobjPDCOPREESTLIST.ProcessRtn_DataCopy(gstrConfigXml,strPREESTNO, strNEWPREESTNO, strJOBNO)
		
		If not gDoErrorRtn ("ProcessRtn_DataCopy") Then
			'모든 플래그 클리어
			gOkMsgBox "복사되었습니다.","내역복사안내!"
			
			.txtFROM.value			= ""
			.txtTO.value			= ""
			.cmbJOBTYPE.value		= ""
			.txtCLIENTCODE1.value	= ""
			.txtCLIENTNAME1.value	= ""
			
			.txtJOBNO.value	  =  strJOBNO
			.txtJOBNAME.value =  strJOBNAME
			
			SelectRtn
			
			For intCnt = 1 To .sprSht.MaxRows 
				If strNEWPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",intCnt) Then
					intEDITCODE = intCnt 
					Exit For
				End If
			Next
			
			mobjSCGLSpr.ActiveCell .sprSht, 1,intEDITCODE
		End If
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
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
	SelectRtn
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
			
			if not gDoErrorRtn ("GetHIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
   		SelectRtn
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 날자관련 COMMAND
'-----------------------------------------------------------------------------------------
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

'-----------------------------------------------------------------------------------------
' JOB 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'실제 데이터List 가져오기
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value))
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub
			.txtJOBNO.value = trim(vntRet(0,0))
			.txtJOBNAME.value = trim(vntRet(1,0))
			SelectRtn
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
					SelectRtn
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' SpreadSheet 관련 Command
'-----------------------------------------------------------------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO, strPREESTNO
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			strJOBNO	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strPREESTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",.sprSht.ActiveRow)
			
			parent.document.forms("frmThis").txtPREESTNO.value = strPREESTNO
			
			If strJOBNO = parent.document.forms("frmThis").txtJOBNO.value Then 
				parent.document.forms("frmThis").txtSELECT.value = "T"
			Else
				parent.document.forms("frmThis").txtSELECT.value = "F"
			End If
			parent.jobMst_Call
			
			mobjSCGLSpr.ActiveCell .sprSht, strCol, strRow	
		End If
	End With
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCOLUMN
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	'키 움직일때 바인딩

	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				strCOLUMN = "SUMAMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) Then
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
END SUB

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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")  Then
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

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화
'-----------------------------------------------------------------------------------------	
Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
	
	set mobjPDCOPREESTLIST	= gCreateRemoteObject("cPDCO.ccPDCOPREESTLIST")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONFIRMGBN | CONFIRMFLAG | JOBNO | JOBNAME | PREESTNAME | SUMAMT | MEMO | PREESTNO | ENDFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|가/본|견적일|JOBNO|JOB명|견적명|견적금액|비고|견적번호|청구구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|   10|    10|    9|   30|    30|      15|  30|      10|      10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONFIRMFLAG", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONFIRMGBN | JOBNO | JOBNAME | PREESTNAME | MEMO | PREESTNO | ENDFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CONFIRMGBN | CONFIRMFLAG | JOBNO | JOBNAME | PREESTNAME | SUMAMT | MEMO | PREESTNO | ENDFLAG"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMGBN | JOBNO | PREESTNO | ENDFLAG",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0

		InitPageData	
		
		SelectRtn
	End With
End Sub

Sub InitPageData
	'초기 데이터 설정
	with frmThis
		'날자관련 전체조회 사용자요청시 취소
		DateClean
		.txtFROM.value = ""
		
		.txtJOBNO.value	  =  parent.document.forms("frmThis").txtJOBNO.value 
		.txtJOBNAME.value =  parent.document.forms("frmThis").txtPRIJOBNAME.value 
		
		Call SEARCHCOMBO_TYPE()
	End with
End Sub

'페이지닫기
Sub EndPage()
	set mobjPDCOPREESTLIST = Nothing
	set mobjPDCOGET = Nothing
	set mobjSCCOGET = Nothing
	
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

'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()'
	Dim vntJOBTYPE
  
   With frmThis   
	'On error resume next
	'Long Type의 ByRef 변수의 초기화
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	
	vntJOBTYPE = mobjPDCOPREESTLIST.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt)  'JOB종류 호출
	
	if not gDoErrorRtn ("COMBO_TYPE") then 
		mobjSCGLSpr.TypeComboBox = True 
		gLoadComboBox .cmbJOBTYPE,  vntJOBTYPE, False
   	end if    				   		
   end with     
End Sub

'-----------------------------------------------------------------------------------------
' 조회
'-----------------------------------------------------------------------------------------
Sub SelectRtn
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim intCnt
   	Dim strJOBNAME
	On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strFROM = REPLACE(.txtFROM.value, "-", "")
		strTO	= REPLACE(.txtTO.value, "-", "")
		
		strJOBNAME = REPLACE(.txtJOBNAME.value,"[","[[]")
		
		
		vntData = mobjPDCOPREESTLIST.SelectRtn_List(gstrConfigXml, mlngRowCnt, mlngColCnt, _
													strFROM, strTO, strJOBNAME, Trim(.txtJOBNO.value), _
													.cmbJOBTYPE.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value)
		If not gDoErrorRtn ("SelectRtn_List") then
			'조회한 데이터를 바인딩
			mobjSCGLSpr.SetClipBinding .sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht.MaxRows '조회된 내역을 처음부터 끝까지 돌면서
					'본견적일 경우 녹색
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) ="본견적" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HD3FED7, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
				Next
			ELSE
				.sprSht.MaxRows = 0	
			End If
			
			gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
		End If	
		
		window.setTimeout "AMT_SUM",1	
		.txtSELECTAMT.value = 0
	END WITH
End Sub

'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		
		If .sprSht.MaxRows > 0 Then
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		else
			.txtSUMAMT.value = 0
		End If
	End With
End Sub

'JOBMST 에서 호출 청구견적의 Est_Copy 의 영향을 받아 여기서 재조회한다.  
Sub PreSelectData
	with frmThis
		.txtJOBNO.value = parent.document.forms("frmThis").txtJOBNOVIEW.value    
		.txtJOBNAME.value = parent.document.forms("frmThis").txtPRIJOBVIEW.value   
		SelectRtn
	End with
End Sub

'-----------------------------------------------------------------------------------------
' 삭제
'-----------------------------------------------------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCount, intRtn, i,intRtn2,lngCnt
	Dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim intChk
	Dim strJOBNO
	Dim intRntChFlag
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "조회된 데이터가 없습니다.","내역복사안내"
			Exit Sub
		end if
		
		'체크가 된 데이터가 있는지 없는지 체크한다.
		intCount = 0
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				intCount = intCount + 1
			end if
		next
		
		'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
		if intCount = 0 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, "삭제안내"
			Exit Sub
		end if
		
		for i = 1 to .sprSht.MaxRows
			strJOBNO = "" : strPREESTNO = ""
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strJOBNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO", i)
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO", i)	
				
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				
				vntData = mobjPDCOPREESTLIST.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strJOBNO, strPREESTNO) 
				
				If mlngRowCnt > 0  Then
					gOkMsgBox i & "행의 견적은 승인상태 또는 청구가 진행되었습니다. 삭제할수 없습니다","삭제안내!"
					Exit Sub
				End if
			end if
		Next
	
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		lngCnt =0
		intRtn2 = 0

		for i = .sprSht.MaxRows to 1 step -1
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)				
				intRtn2 = mobjPDCOPREESTLIST.DeleteRtn(gstrConfigXml, strPREESTNO)
				
				IF not gDoErrorRtn ("DeleteRtn") then
					lngCnt = lngCnt +1
					mobjSCGLSpr.DeleteRow .sprSht, i
   				End IF
			End If
		next
		
		If lngCnt <> 0 Then
			gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
			If .sprSht.MaxRows = 0 Then
				strJOBNO	 = parent.document.forms("frmThis").txtJOBNO.value 
				
				parent.document.forms("frmThis").txtPREESTNO.value  = ""
				
				intRntChFlag = mobjPDCOPREESTLIST.FlagUpdateRtn(gstrConfigXml, strJOBNO)
			End If
		End If
		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		
		SelectRtn
		
		parent.jobMst_Tab2Search
		parent.jobMst_Tab5Search
	End with
	err.clear
End Sub

		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle1" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD0" align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="300" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="65" background="../../../images/back_p.gIF"
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
											<td class="TITLE">견적리스트</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start--></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<TABLE class="SEARCHDATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" align="left"
							border="0">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
									width="60">견적일</TD>
								<TD class="SEARCHDATA" width="214"><INPUT class="INPUT" id="txtFROM" title="기간검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
										accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"> <IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
										border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="기간검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
									width="50">JOB명</TD>
								<TD class="SEARCHDATA" width="235"><INPUT class="INPUT_L" id="txtJOBNAME" title="제작관리명 조회" style="WIDTH: 145px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="29" name="txtJOBNAME"> <IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgJOBNO">
									<INPUT class="INPUT" id="txtJOBNO" title="제작관리코드 조회" style="WIDTH: 60px; HEIGHT: 22px"
										type="text" maxLength="7" align="left" size="3" name="txtJOBNO"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbJOBTYPE,'')"
									width="45">유형</TD>
								<TD class="SEARCHDATA" width="100"><select id="cmbJOBTYPE" style="WIDTH: 100px">
									</select></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">광고주</TD>
								<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="조회용광고주명" style="WIDTH: 140px; HEIGHT: 22px"
										type="text" maxLength="100" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="조회용광고주코드" style="WIDTH: 57px; HEIGHT: 22px"
										type="text" maxLength="7" size="4" name="txtCLIENTCODE1">
								</TD>
								<TD class="SEARCHDATA" align="right" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
										align="right" border="0" name="imgQuery"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" id="spacebar" style="WIDTH: 100%; HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20">
									<table style="WIDTH: 640px; HEIGHT: 20px" cellSpacing="0" cellPadding="0" width="640" border="0">
										<tr>
											<td class="TITLE">견적리스트 합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">&nbsp;&nbsp; 
												선택합계 : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgListcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													height="20" alt="선택한행의복사를 합니다." src="../../../images/imglistcopy.gIF" border="0"
													name="imgListcopy"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="13679">
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
				</tr>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
