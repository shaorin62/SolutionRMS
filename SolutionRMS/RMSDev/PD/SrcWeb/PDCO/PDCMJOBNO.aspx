<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBNO.aspx.vb" Inherits="PD.PDCMJOBNO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>제작관리번호 등록</title>
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
Dim mlngRowCnt, mlngColCnt		'조회시 로우,컬럼 갯수를 반환
Dim mlngRowCnt2,mlngColCnt2		'조회시 로우,컬럼 갯수를 반환
Dim mobjPDCMJOBNO, mobjPDCMACTUALRATE, mobjSCCOGET, mobjPDCMGET '모듈(JOBNO-CRUD, ACTUALRATE-CRUD, 전체공통, 제작공통)
Dim mstrFlag					'입력시 Insert,Update 구분
Dim mstrBindCHK					'COMBO 의 SUBCOMBO 호출필요성 체크
Dim mstrHIDDEN					'입력필드의 숨기기
Dim mstrNoClick					'전체체크시 발동 하나, 해당 화면에서 미사용


Const meTab = 9
mstrFlag = "SELECT"
mstrNoClick = False
mstrBindCHK = False
mstrHIDDEN = 0

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
' 명령버튼
'=========================================================================================
'조회버튼
Sub imgQuery_onclick
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
	
End Sub

'신규버튼
Sub imgNew_onclick
	NewRegNo
End Sub

'저장버튼
Sub imgSave_onclick ()
	If frmThis.cmbENDFLAG.value = "PF01" Or  frmThis.cmbENDFLAG.value = "PF02" Then
	Else
		gErrorMsgBox "진행상태가 의뢰 및 진행 이 아닌건은 수정될수 없습니다.","저장안내"
		SelectRtn
		exit Sub
	End If
			
	gFlowWait meWAIT_ON
	ProcessRtn
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

'닫기버튼
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'삭제버튼
Sub imgDelete_onclick()
	with frmThis 
	
	End with 
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'추가버튼
sub imgAddRow_onclick ()
	Dim strREG_NUM
	
	With frmThis
	
		strREG_NUM	= .txtREG_NUM.value
		
		IF strREG_NUM = "" THEN
			gErrorMsgBox "추가할 수 없습니다.","추가안내!"
			Exit Sub
		Else
			call sprSht_Keydown(meINS_ROW, 0)
		End if
	End With 
end sub

'부서실적등록버튼 최초 클릭
Sub ImgDivamtPop_onclick
	Call ACTUALRATE_POP()
	SelectRtn
End Sub

'부서실적등록 버튼 실제동작
Sub ACTUALRATE_POP
	Dim vntRet, vntInParams
	Dim strJOBNO , strJOBNAME
	with frmThis
		
		strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
		strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow)
		
		If .sprSht.MaxRows = 0  Then
			gErrorMsgBox "선택된데이터가 없습니다.","처리안내"
			Exit Sub
		End If
		
		If  mstrFlag = "NEW" Then
			gErrorMsgBox "등록후 실적분배율을 입력할수 있습니다.","처리안내"
			Exit Sub
		End If
		
		vntInParams = array(trim(strJOBNO),trim(strJOBNAME))
		vntRet = gShowModalWindow("PDCMACTUALRATEPOP.aspx",vntInParams , 1060,800)
		
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


'신규 등록
Sub NewRegNo
	mstrNoClick = True
	mstrFlag = "NEW"
	Dim vntRet
	Dim vntInParams
	dim intRtn
	DataClean
	
	with frmThis
	
		vntInParams = array(trim(.txtPROJECTNO.value), trim(.txtPROJECTNM.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			.cmbJOBGUBN.disabled = false
			.cmbCREPART.disabled = false
			
			if .txtPROJECTNO.value = vntRet(0,0) and .txtPROJECTNM.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM.value = trim(vntRet(1,0))  ' 코드명 표시
			.txtCLIENTNAME.value = trim(vntRet(2,0))  ' 코드명 표시
			.txtSUBSEQNAME.value = trim(vntRet(4,0))  ' 코드명 표시
			.txtGROUPGBN.value = trim(vntRet(5,0))  ' 코드명 표시	
			.txtCREDAY.value = trim(vntRet(6,0))  ' 코드명 표시
			.txtCPDEPTNAME.value = trim(vntRet(7,0))  ' 코드명 표시
			.txtCPEMPNAME.value = trim(vntRet(8,0))  ' 코드명 표시
			.txtCLIENTTEAMNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			.txtMEMO.value = trim(vntRet(9,0))
			If mstrHIDDEN = 0 Then
				.txtJOBNAME.focus()					' 포커스 이동
			End If
			call sprSht_Keydown(meINS_ROW, 0)
			DataFill
			 mobjSCGLSpr.SetCellsLock2 .sprSht,false,"JOBNAME|JOBGUBN|REQDAY|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|BIGO|CREPART|CREDEPTNAME|CREEMPNAME|EXCLIENTCODE|EXCLIENTNAME",1,1,false
     	end if
	End with
	
End Sub
'=========================================================================================
' 칼렌더버튼 및 기타 버튼
'=========================================================================================

'상위입력 필드 숨기기
Sub Set_SELECTTBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblSelectBody").style.display = "inline"
		Else
			document.getElementById("tblSelectBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub


' 하위입력 필드 숨기기
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='입력필드 숨김' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblBody2").style.display = "inline"
			document.getElementById("spacebar1").style.display = "inline"
			document.getElementById("spacebar2").style.display = "inline"
			
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='입력필드 노출' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblBody2").style.display = "none"
			document.getElementById("spacebar1").style.display = "none"
			document.getElementById("spacebar2").style.display = "none"
	
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

'=========================================================================================
' Sub 에서 호출하는 Sub
'=========================================================================================
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		
		.txtHOPEENDDAY.value = gNowDate
		.txtREQDAY.value = gNowDate
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""
	
	Call COMBO_TYPE()
	Call SUBCOMBO_TYPE()
	Call SEARCHCOMBO_TYPE()
	End with
	DataNewClean
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub EndPage()
	set mobjPDCMJOBNO = Nothing
	set mobjPDCMGET = Nothing
	set mobjPDCMACTUALRATE = Nothing
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

Sub DataNewClean
	with frmThis
		.cmbJOBGUBN.selectedIndex = -1
		.cmbCREPART.selectedIndex = -1
		.cmbCREGUBN.selectedIndex = -1
		.cmbENDFLAG.selectedIndex = -1 
		.cmbJOBBASE.selectedIndex = -1
		.txtREQDAY.value = ""
		.txtHOPEENDDAY.value = ""	
	End with
End Sub

Sub imgEndChange_onclick

Dim vntData2
Dim strCODE
Dim intRtnSave
Dim intRtn
Dim intCnt
Dim intCode
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCODE = .txtJOBNO.value

		If strCODE = "" Then
		gErrorMsgBox "우선 JOBNO 를 조회하십시오.","처리안내"
		Exit Sub
		End If

		If .cmbENDFLAG.value <> "PF02" Then
		gErrorMsgBox "완료구분 변경은 '진행' 일경우만 가능합니다.","처리안내"
		Exit Sub
		End If
		vntData2 = mobjPDCMJOBNO.GetJOBNOSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
		
		If mlngRowCnt = 0 Then
			intRtnSave = gYesNoMsgbox("완료구분을 '의뢰'상태 로 변경하시겠습니까?","처리안내")
			IF intRtnSave <> vbYes then exit Sub
			intRtn = mobjPDCMJOBNO.ProcessRtn_ENDFLAG(gstrConfigXml,strCODE)
			if not gDoErrorRtn ("ProcessRtn_ENDFLAG") then
				gErrorMsgBox "JOBNO [" & strCODE & " ]완료구분이 '의뢰' 상태로 변경되었습니다.","처리안내" 
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
						intCode = intCnt 
						Exit For
					End If
				Next
				mobjSCGLSpr.ActiveCell .sprSht, 1,intCode
			end if
		Else
			gErrorMsgBox "해당 JOBNO 의 외주정산내역을 확인하십시오","처리안내"
		End If

	End with
End Sub


Sub DataFill
	with frmThis
	mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTNO",.sprSht.ActiveRow, .txtPROJECTNO.value
	mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTNM",.sprSht.ActiveRow, .txtPROJECTNM.value
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBBASE",.sprSht.ActiveRow, .cmbJOBBASE.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"JOBBASENAME",.sprSht.ActiveRow, .cmbJOBBASE(.cmbJOBBASE.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CREGUBN",.sprSht.ActiveRow, .cmbCREGUBN.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"CREGUBNNAME",.sprSht.ActiveRow, .cmbCREGUBN(.cmbCREGUBN.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CREPART",.sprSht.ActiveRow, .cmbCREPART.value
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBGUBN",.sprSht.ActiveRow, .cmbJOBGUBN.value
	mobjSCGLSpr.SetTextBinding .sprSht,"ENDFLAG",.sprSht.ActiveRow, .cmbENDFLAG.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"ENDFLAGNAME",.sprSht.ActiveRow, .cmbENDFLAG(.cmbENDFLAG.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
	mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, .txtSUBSEQNAME.value
	mobjSCGLSpr.SetTextBinding .sprSht,"REQDAY",.sprSht.ActiveRow, gNowDATE
	mobjSCGLSpr.SetTextBinding .sprSht,"HOPEENDDAY",.sprSht.ActiveRow, gNowDATE
	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTTEAMNAME.value 
	End with
End Sub


'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT", lngCnt)
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

'클리어
Sub CleanField (objField1, objField2)
	If frmThis.sprSht.MaxRows > 0 Then
			if isobject(objField1) then 
				objField1.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField1.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			end if
			if isobject(objField2) then 
				objField2.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField2.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			End If
	End If
End Sub

'화면클리어
Sub DataClean
	with frmThis
		
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = "" 
		.txtCLIENTNAME.value =  ""
		.txtCPDEPTNAME.value =  ""
		.txtCREDAY.value = ""
		.txtCPEMPNAME.value =  ""
		.txtGROUPGBN.value = ""
		.txtSUBSEQNAME.value =  ""
		.txtCLIENTTEAMNAME.value = ""
		.txtMEMO.value =  ""
		.txtJOBNAME.value =  ""
		.txtJOBNO.value =  ""
		.cmbJOBGUBN.selectedIndex = 0 
		.txtDEPTNAME.value =  ""
		.txtDEPTCD.value =  ""
		.txtREQDAY.value = gNowDate
		'.cmbCREPART.selectedIndex = 0  '여기서 0이아닌 구분에 따른걸 가져옴
		SUBCOMBO_TYPE
		.txtEMPNAME.value =  "" 
		.txtEMPNO.value =  ""
		.txtHOPEENDDAY.value = gNowDate 
		.cmbCREGUBN.selectedIndex = 0 
		.cmbJOBBASE.selectedIndex = 0 
		.txtCREDEPTNAME.value = "" 
		.txtCREDEPTCD.value =  ""
		.cmbENDFLAG.selectedIndex = 0 
		.txtCREEMPNAME.value =  ""
		.txtCREEMPNO.value =  ""
		.txtAGREEYEARMON.value = "" 
		.txtDEMANDYEARMON.value =  ""
		.txtSETYEARMON.value =  ""
		.txtBUDGETAMT.value =  ""
		.txtBIGO.value =  ""
		.txtEXCLIENTCODE.value = ""
		.txtEXCLIENTNAME.value = ""
		
		.sprSht.MaxRows = 0
	End With
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
Sub InitPage()
	'서버업무객체 생성	
	
	set mobjPDCMJOBNO = gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjPDCMACTUALRATE = gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
   
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 40, 0, 3, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 15, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 22, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 24, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 27, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "JOBNAME|JOBNO|CLIENTNAME|CLIENTTEAMNAME|SUBSEQNAME|JOBGUBN|REQDAY|ENDFLAG|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|BTN_DEPT|EMPNAME|BTN_EMP|HOPEENDDAY|BUDGETAMT|CREPART|CREGUBN|JOBBASE|CREDEPTNAME|BTN_CDEPT|CREEMPNAME|BTN_CEMP|EXCLIENTCODE|EXCLIENTNAME|BTN_EXCLIENTCODE|BIGO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CREDAY|CLIENTTEAMCODE|EMPNO|CREDEPTCD|CREEMPNO"
		mobjSCGLSpr.SetHeader .sprSht,        "JOB명|JOBNO|광고주|팀|브랜드|매체부문|의뢰일|완료구분|합의월|청구월|결산월|부서코드|담당팀|담당자|완료예정일|예산금액|매체분류|신규구분|정산대상|제작담당팀|제작담당자|크리코드|크리조직|비고|프로젝트코드|프로젝트명|그룹구분|CP부서|CP담당자|메모|등록일|팀코드|의뢰자NO|제작부서코드|제작사번"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   20|    9|    20|20|    20|      12|    10|       8|     8|     8|     8|       0|    12|2|   8|2|       9|      11|      10|       8|      10|      10|2|       8|2|0       |15|    2|  10|           0|         0|       0|    0|        0|   0|     0|     0|       0|           0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "EXCLIENTCODE|EXCLIENTNAME|BIGO", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|HOPEENDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "BUDGETAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_DEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_EMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_CDEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_CEMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_EXCLIENTCODE"
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "JOBNO|JOBNAME|CLIENTNAME|SUBSEQNAME|JOBGUBN|REQDAY|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|ENDFLAG|CREDEPTCD|CREDEPTNAME|CREEMPNO|CREEMPNAME|BIGO|CREPART|CLIENTTEAMNAME|EXCLIENTCODE|EXCLIENTNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|CLIENTTEAMNAME|SUBSEQNAME|DEPTNAME|EMPNAME|CREDEPTNAME|CREEMPNAME|BIGO|EXCLIENTNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|EXCLIENTCODE",-1,-1,2,2,false '가운데
		mobjSCGLSpr.colhidden .sprSht, "DEPTCD|EMPNO|CREDEPTCD|CREEMPNO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CLIENTTEAMCODE|EXCLIENTCODE",true
		.sprSht.style.visibility = "visible"
		'If .cmbENDFLAG.value = "PF01" Or .cmbENDFLAG.value = "PF02" Then 
		'	.cmbENDFLAG.disabled = false
		'Else 
			.cmbENDFLAG.disabled = true
		'End If
		.cmbPOPUPTYPE.value=2
		


    End With

	'화면 초기값 설정
	InitPageData
	
End Sub
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()

	mstrNoClick = False
	mstrFlag = "SELECT"
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
    Dim intCnt
    Dim strCODE
    Dim vntDataSubCombo
	'On error resume next
	with frmThis
	
	
	
		'Sheet초기화
		.sprSht.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
	
		
		vntData = mobjPDCMJOBNO.SelectRtn_PROJECTORJOB(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtSEARCHCLIENTSUBCODE.value),Trim(.txtSEARCHCLIENTSUBNAME.value),Trim(.txtSEARCHCLIENTCODE.value),Trim(.txtSEARCHCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value,Trim(.txtPROJECTNO1.value),Trim(.txtPROJECTNM1.value),Trim(.cmbPOPUPTYPE.value))
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
				DATACLEAN	
				DataNewClean
			Else
				If .txtBUDGETAMT.value <> "" Then
				txtBUDGETAMT_onblur
				End If
				'SetCellsLock 설정
				for intCnt = 1 to .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",intCnt) = "PF01" Or  mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",intCnt) = "PF02" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"JOBNAME|JOBGUBN|REQDAY|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|BIGO|CREPART|CREDEPTNAME|CREEMPNAME|EXCLIENTCODE|EXCLIENTNAME",intCnt,intCnt,false
					End If
					strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",intCnt)
					'Call Get_SUBCOMBO_VALUE(strCODE)
					mlngRowCnt2=clng(0)
					mlngColCnt2=clng(0)
					'무한루프 걸리기 쉬운 Combo Setting [만약 조회시 매체부분의 의 해당분류만 보고 싶다면 (매체부분을 움직이지 않고 해당분류만 보고싶다면 아래의 로직을 추가하여야 한다.- 속도가 굉장히 느려짐]
					'vntDataSubCombo = mobjPDCMJOBNO.GetDataType_SubCode(gstrConfigXml, mlngRowCnt2, mlngColCnt2, strCODE)					
									
					'mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",intCnt,intCnt,vntDataSubCombo,,77			
					'mobjSCGLSpr.TypeComboBox = True 			
   				
				Next
			End If
			
			gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			sprShtToFieldBinding 1,1
			
			
	
			
		End If		
		.cmbJOBGUBN.disabled = true
		.cmbCREPART.disabled = true
		.txtSELECTAMT.value = 0

	'조회완료메세지
	AMT_SUM
	
	End With
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub


'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
  	Dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strJOBNO
	Dim intCnt
	Dim intCode,intEDITCODE
	Dim strEDITJOBNO
	Dim intRtnSave
	with frmThis
	'On error resume next
		strJOBNO = ""
		
  		'데이터 Validation
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNAME|JOBNO|SUBSEQNAME|JOBGUBN|REQDAY|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREPART|CREGUBN|JOBBASE|ENDFLAG|CREDEPTCD|CREDEPTNAME|CREEMPNO|CREEMPNAME|BIGO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CLIENTNAME|CREDAY|EXCLIENTCODE")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		if DataValidation =false then exit sub
		'strCODE = .txtPROJECTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
		
		
	
		'처리 업무객체 호출
		strMasterData = gXMLGetBindingData (xmlBind)
		
		if .txtJOBNO.value = "" then
			strSEQFlag = "new"
			intRtn = mobjPDCMJOBNO.ProcessRtn(gstrConfigXml,strMasterData, strSEQFlag,strJOBNO)
		else
			intRtn = mobjPDCMJOBNO.ProcessRtnSheet(gstrConfigXml,vntData)
		end if
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if strSEQFlag = "new" then
				
				'gErrorMsgBox " 자료가 신규저장" & mePROC_DONE,"저장안내"
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
					intCode = intCnt 
					Exit For
					End If
				Next
				mobjSCGLSpr.ActiveCell .sprSht, 1,intCode
				sprShtToFieldBinding 1,intCode		
				intRtnSave = gYesNoMsgbox("자료가 저장 되었습니다. 실적분배율을 입력 하시겠습니까?","처리안내")
				IF intRtnSave <> vbYes then exit Sub
				Call ImgDivamtPop_onclick()	
				sprSht_Click 1,intCode					

			else
				gErrorMsgBox " 자료가" & intRtn & " 건 수정저장" & mePROC_DONE,"저장안내" 
				strEDITJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.activeRow)
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strEDITJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
						intEDITCODE = intCnt 
						Exit For
					End If
				Next

				mobjSCGLSpr.ActiveCell .sprSht, 1,intEDITCODE
				sprShtToFieldBinding .sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
			
  		end if
 	end with
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
   		If .cmbCREPART.value = "PC01" Or .cmbCREPART.value = "PR01" Then
			If .txtEXCLIENTCODE.value = "" Then
				gErrorMsgBox "매체분류가 TV-CF 또는 Radio-CM 일때 크리조직 입력은 필수 입니다.","입력안내"
				.txtEXCLIENTNAME.focus()
				Exit Function
			End If
		End If
   	
   	End with
	DataValidation = true
End Function


'------------------------------------------
' 데이터 입력시 SHEET BINDING onchange EVENT
'------------------------------------------
'생략부분
Sub txtJOBNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBNAME",frmThis.sprSht.ActiveRow, frmThis.txtJOBNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtJOBNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBNO",frmThis.sprSht.ActiveRow, frmThis.txtJOBNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbJOBGUBN_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		SUBCOMBO_TYPE
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBGUBN",frmThis.sprSht.ActiveRow, frmThis.cmbJOBGUBN.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBGUBNNAME",frmThis.sprSht.ActiveRow, frmThis.cmbJOBGUBN(frmThis.cmbJOBGUBN.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtDEPTCD_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPTCD",frmThis.sprSht.ActiveRow, frmThis.txtDEPTCD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtREQDAY_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REQDAY",frmThis.sprSht.ActiveRow, frmThis.txtREQDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbCREPART_onchange
	if frmThis.sprSht.ActiveRow >0 AND mstrBindCHK = False Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREPART",frmThis.sprSht.ActiveRow, frmThis.cmbCREPART.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREPARTNAME",frmThis.sprSht.ActiveRow, frmThis.cmbCREPART(frmThis.cmbCREPART.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtEMPNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EMPNAME",frmThis.sprSht.ActiveRow, frmThis.txtEMPNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtEMPNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EMPNO",frmThis.sprSht.ActiveRow, frmThis.txtEMPNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtHOPEENDDAY_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"HOPEENDDAY",frmThis.sprSht.ActiveRow, frmThis.txtHOPEENDDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'여기부터
Sub cmbCREGUBN_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREGUBN",frmThis.sprSht.ActiveRow, frmThis.cmbCREGUBN.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREGUBNNAME",frmThis.sprSht.ActiveRow, frmThis.cmbCREGUBN(frmThis.cmbCREGUBN.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'cmbJOBBASE
Sub cmbJOBBASE_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBBASE",frmThis.sprSht.ActiveRow, frmThis.cmbJOBBASE.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBBASENAME",frmThis.sprSht.ActiveRow, frmThis.cmbJOBBASE(frmThis.cmbJOBBASE.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtCREDEPTNAME
Sub txtCREDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREDEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCREDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCREDEPTCD_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREDEPTCD",frmThis.sprSht.ActiveRow, frmThis.txtCREDEPTCD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbENDFLAG_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDFLAG",frmThis.sprSht.ActiveRow, frmThis.cmbENDFLAG.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDFLAGNAME",frmThis.sprSht.ActiveRow, frmThis.cmbENDFLAG(frmThis.cmbENDFLAG.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtCREEMPNAME
Sub txtCREEMPNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREEMPNAME",frmThis.sprSht.ActiveRow, frmThis.txtCREEMPNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCREEMPNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREEMPNO",frmThis.sprSht.ActiveRow, frmThis.txtCREEMPNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtBUDGETAMT_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUDGETAMT",frmThis.sprSht.ActiveRow, frmThis.txtBUDGETAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtBIGO
Sub txtBIGO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BIGO",frmThis.sprSht.ActiveRow, frmThis.txtBIGO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'PROJECT,JOBNO 선택
Sub cmbPOPUPTYPE_onchange
	with frmThis
		.txtPROJECTNM1.value = ""
		.txtPROJECTNO1.value = ""
	End with
	gSetChange
End Sub


Sub txtEXCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEXCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub




'****************************************************************************************
' UI 시작
'****************************************************************************************
'입력용
'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
	Dim strCODE
    With frmThis   

		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB종류 호출
		vntJOBBASE = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBBASE")  '청구기준 호출	
		vntCREGUBN = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREGUBN")  '신규/기존 호출
		vntENDFLAG = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '제작상태 호출
		vntCREPART = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREPART")  
		if not gDoErrorRtn ("COMBO_TYPE") then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "JOBGUBN",,,vntJOBGUBN,,95 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "JOBBASE",,,vntJOBBASE,,77
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREGUBN",,,vntCREGUBN,,65 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "ENDFLAG",,,vntENDFLAG,,65 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",,,vntCREPART,,77
			
			mobjSCGLSpr.TypeComboBox = True 
			 gLoadComboBox .cmbENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbJOBGUBN, vntJOBGUBN, False
			 gLoadComboBox .cmbJOBBASE, vntJOBBASE, False
			 gLoadComboBox .cmbCREGUBN, vntCREGUBN, False 
			 gLoadComboBox .cmbCREPART, vntCREPART, False 
			 'strCODE = vntJOBGUBN(0,1)
			 'Call Get_SUBCOMBO_VALUE(strCODE)	
   		end if    
   		'cmbJOBGUBN_onchange   		
   	end with     	
End Sub
'조회용

'Dynamic Combo
Sub Get_SUBCOMBO_VALUE(strCODE,strPos)							
	Dim vntData					
	With frmThis   					
		On error resume Next				
		'Long Type의 ByRef 변수의 초기화				
		mlngRowCnt=clng(0)				
		mlngColCnt=clng(0)				

       	vntData = mobjPDCMJOBNO.GetDataType_SubCode(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)					
		If not gDoErrorRtn ("GetDataType_SubCode") Then 				
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",strPos,strPos,vntData,,77		
			
			mobjSCGLSpr.TypeComboBox = True 			
   		End If  				
   		gSetChange				
   	end With   					
End Sub		
'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
    With frmThis   

		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB종류 호출
		vntENDFLAG = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '제작상태 호출
		if not gDoErrorRtn ("SEARCHCOMBO_TYPE") then 
			 gLoadComboBox .cmbSEARCHENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbSEARCHJOBGUBN, vntJOBGUBN, False
			 
			' mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",,,vntCREPART,,77
   		end if    				   		
   	end with     
   		
End Sub
'-----------------------------------------------------------------------------------------
' SUBCOMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()

	Dim vntCREPART
   	Dim vntCREGUBN
   
	With frmThis   

		'On error resume next
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
       	
	       	vntCREPART = mobjPDCMJOBNO.GetDataTypeChange(gstrConfigXml, mlngRowCnt, mlngColCnt,.cmbJOBGUBN.value,"K")  '제작종류 호출	

		if not gDoErrorRtn ("SUBCOMBO_TYPE") then 
			 gLoadComboBox .cmbCREPART, vntCREPART, False
   		end if  
   		cmbCREPART_onchange		   		
   	end with   
End Sub



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
Sub txtBUDGETAMT_onfocus
	with frmThis
		.txtBUDGETAMT.value = Replace(.txtBUDGETAMT.value,",","")
	end with
End Sub
Sub txtBUDGETAMT_onblur
	with frmThis
		call gFormatNumber(.txtBUDGETAMT,0,true)
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub







Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtHOPEENDDAY,frmThis.imgCalEndar,"txtHOPEENDDAY_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarREQ_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtREQDAY,frmThis.imgCalEndar,"txtREQDAY_onchange()"
		gSetChange
	end with
End Sub

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
		vntInParams = array( trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' 코드명 표시
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
' 사업부코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------

'이미지버튼 클릭시
Sub ImgSEARCHCLIENTSUBCODE_onclick
	with frmThis
			Call SEARCHTEAM_POP()
	End with
	
End Sub
'광고팀 조회팝업
Sub SEARCHTEAM_POP
	Dim vntRet, vntInParams
	with frmThis
		'광고주코드,광고주명,팀코드,팀명
		
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value) , trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value) , trim(.txtSEARCHCLIENTSUBNAME.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
			.txtSEARCHCLIENTSUBCODE.value = trim(vntRet(0,0))
			.txtSEARCHCLIENTSUBNAME.value = trim(vntRet(1,0))
			.txtSEARCHCLIENTCODE.value = trim(vntRet(4,0))	'Code값 저장
			.txtSEARCHCLIENTNAME.value = trim(vntRet(5,0))	'코드명 표시
			
		 
			.txtSEARCHCLIENTNAME.focus()
		end if
	end with

End SUb

'실제 데이터List 가져오기
Sub SEARCHCLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value), trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value), trim(.txtSEARCHCLIENTSUBNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtSEARCHCLIENTSUBCODE.value = vntRet(0,0) and .txtSEARCHCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtSEARCHCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtSEARCHCLIENTSUBNAME.value = trim(vntRet(1,0))  ' 코드명 표시
			.txtSEARCHCLIENTCODE.value = trim(vntRet(3,0))
			.txtSEARCHCLIENTNAME.value = trim(vntRet(4,0))
			
			.txtSEARCHCLIENTNAME.focus()					' 포커스 이동
			'gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtSEARCHCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
				vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSEARCHCLIENTCODE.value),trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value),trim(.txtSEARCHCLIENTSUBNAME.value))
			
				if not gDoErrorRtn ("txtSEARCHCLIENTSUBNAME_onkeydown") then
					If mlngRowCnt = 1 Then
						.txtSEARCHCLIENTSUBCODE.value = trim(vntData(0,1))
						.txtSEARCHCLIENTSUBNAME.value = trim(vntData(1,1))
						.txtSEARCHCLIENTCODE.value = trim(vntData(4,1))
						.txtSEARCHCLIENTNAME.value = trim(vntData(5,1))
						.txtSEARCHCLIENTNAME.focus()
					Else
						Call SEARCHTEAM_POP()
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
Sub ImgSEARCHCLIENTCODE_onclick
	Call SEARCHCLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub SEARCHCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value), trim(.txtSEARCHCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtSEARCHCLIENTCODE.value = vntRet(0,0) and .txtSEARCHCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtSEARCHCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtSEARCHCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시	
			.txtPROJECTNM1.focus()					' 포커스 이동
			'gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
     	
	End with
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtSEARCHCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSEARCHCLIENTCODE.value),trim(.txtSEARCHCLIENTNAME.value),"A")
			
			if not gDoErrorRtn ("txtSEARCHCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtSEARCHCLIENTCODE.value = trim(vntData(0,1))
					.txtSEARCHCLIENTNAME.value = trim(vntData(1,1))
					
				Else
					Call SEARCHCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------
' 담당부서 조회 
'-----------------------------
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
' Project 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'ProjectNO 조회팝업
Sub ImgPROJECTNO1_onclick
	with frmThis
		'1은 PROJECT 조회   2는 JOBNO조회
		IF .cmbPOPUPTYPE.value = "1" then
			Call PONO_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	
	End with
End Sub
'실제 데이터List 가져오기
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' 코드명 표시
			'.txtCLIENTNAME1.focus()					' 포커스 이동
     	end if
	End with
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
		if .cmbPOPUPTYPE.value = "1" Then '프로젝트 코드 라면
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
		Else
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
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
			
			.txtCREDEPTNAME.focus()
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
			gSetChangeFlag .txtDEPTCD
			gSetChangeFlag .txtDEPTNAME
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
					.txtCREDEPTNAME.focus()
					'.txtMEMO.focus()
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
'-----------------------------
' 제작부서 조회 
'-----------------------------
Sub ImgCREDEPTCD_onclick
	Call CREDEPT_POP()
End Sub

Sub CREDEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC 코드/명,optional(현재사용여부,사용검사일,추가조회 필드,Key Like여부)
		vntInParams = array(trim(.txtCREDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCREDEPTCD.value = trim(vntRet(0,0))	'Code값 저장
			.txtCREDEPTNAME.value = trim(vntRet(1,0))	'코드명 표시
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtCREEMPNAME.focus()
			gSetChangeFlag .txtCREDEPTCD
		end if
	end with
End Sub

Sub txtCREDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREDEPTNAME.value)
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = trim(vntData(0,0))
					.txtCREDEPTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtCREEMPNAME.focus()
				Else
					Call CREDEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 제작사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCREEMPNO_onclick
	Call CREEMP_POP()
End Sub

'실제 데이터List 가져오기
Sub CREEMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCREDEPTCD.value), trim(.txtCREDEPTNAME.value), trim(.txtCREEMPNO.value), trim(.txtCREEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCREEMPNO.value = vntRet(0,0) and .txtCREEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCREDEPTCD.value = trim(vntRet(2,0))  ' Code값 저장
			.txtCREDEPTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			.txtCREEMPNO.value = trim(vntRet(0,0))
			.txtCREEMPNAME.value = trim(vntRet(1,0))
			
			if .sprSht.ActiveRow >0 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtBUDGETAMT.focus()					' 포커스 이동
			gSetChangeFlag .txtCREEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtCREEMPNAME
			gSetChangeFlag .txtCREDEPTCD
			gSetChangeFlag .txtCREDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCREEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREEMPNO.value, .txtCREEMPNAME.value,"A",.txtCREDEPTCD.value,.txtCREDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCREEMPNO.value = trim(vntData(0,1))
					.txtCREEMPNAME.value = trim(vntData(1,1))
					.txtCREDEPTCD.value = trim(vntData(2,1))
					.txtCREDEPTNAME.value = trim(vntData(3,1))
					
					if .sprSht.ActiveRow >0 Then
			
						mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
						
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtBUDGETAMT.focus()
					gSetChangeFlag .txtCREEMPNO
				Else
					Call CREEMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 대행사 코드팝업 버튼
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value), "") '<< 받아오는경우
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEXCLIENTCODE.value = trim(vntRet(1,0))  ' Code값 저장
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))  ' 코드명 표시
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, .txtEXCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, .txtEXCLIENTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtBIGO.focus() 
			gSetChangeFlag .txtCREEMPNO	
     	End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value), "")
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))
					.txtEXCLIENTNAME.value = trim(vntData(2,1))
					if .sprSht.ActiveRow >0 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, .txtEXCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, .txtEXCLIENTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtBIGO.focus()
					gSetChangeFlag .txtEXCLIENTCODE
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub


Sub JOBGUBNClean
	with frmThis
		.cmbSEARCHJOBGUBN.selectedIndex = 0
	End with
End Sub
Sub ENDFLAGClean
	with frmThis
		.cmbSEARCHENDFLAG.selectedIndex = 0
	End With
End Sub

Sub DeleteRtn
	Dim vntData
	Dim intSelCnt, intRtn, i , intCnt
	Dim strCODE , strENDFLAG
	Dim intSubRtn
	
	with frmThis
		
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		

		If intSelCnt = 0 Or .sprSht.MaxRows = 0 Then
			gErrorMsgBox "삭제할 자료가 없습니다.","삭제안내"
			Exit Sub
		End If
		
		'실적분배율 있으면 삭제 불가능
		'If .chkDEPT.checked = TRUE or .chkRATE.checked =TRUE Then
		'	gErrorMsgBox "실적분배비율이 있습니다.","삭제안내"
		'	Exit Sub
		'End if
		
		for i = intSelCnt-1 to 0 step -1
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strENDFLAG = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",vntData(i))
		
				If strENDFLAG <> "PF01" Then
					gErrorMsgBox "[" & i & "행] 의진행상태가 의뢰가 아닌건은 삭제하실수 없습니다.","삭제안내!"
					Exit Sub
				End If
			End IF
		next
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",vntData(i))
				'자료 삭제
				intRtn = mobjPDCMJOBNO.DeleteRtn(gstrConfigXml,strCODE)
				'실적분배율까지 바로삭제
				intSubRtn =mobjPDCMACTUALRATE.DeleteRtn_DTL_JOBNODEPT_JOBNO(gstrConfigXml,strCODE)
				intSubRtn =mobjPDCMACTUALRATE.DeleteRtn_DTL_ACTUALRATE_JOBNO(gstrConfigXml,strCODE)
			End IF
		next
		

		IF not gDoErrorRtn ("DeleteRtn_DTL_ACTUALRATE_JOBNO") then
			'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			gWriteText "", "선택한 데이터" & intSelCnt & "건이 삭제" & mePROC_DONE
   		End IF
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear
End Sub



'------------------------------------------
' SHEET EVENT
'------------------------------------------

Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
	
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
	If mstrNoClick = True Then Exit Sub
		if Row > 0 and Col > 0 then		
			
			sprShtToFieldBinding Col,Row
		End If
		'if Col = 20 Then
		'msgbox "된다."
		'End If
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim strJOBGUBN
	Dim strCode
	Dim strCodeName
	Dim strDeptCodeName
	Dim vntData
	with frmThis
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		.txtBIGO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row)
		.txtBUDGETAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT",Row)
		.cmbCREPART.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREPART",Row)
		.cmbCREGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREGUBN",Row)
		.cmbENDFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",Row)
		.cmbJOBBASE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBBASE",Row)
		.cmbJOBGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row)
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		'매체부문 변경시 Subtype 설정
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"JOBGUBN") Then 
			strJOBGUBN = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row)
			
			Call Get_SUBCOMBO_VALUE(strJOBGUBN,Row)
			SUBCOMBO_TYPE
			
		'매체분류 Subtype 적용 Validation	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREPART") Then '20
			If .cmbCREPART.value = "" Then
			gErrorMsgBox "선택하신 분류는 해당부문의 분류사항이 아닙니다.","처리안내!"
			sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"JOBGUBN"),.sprSht.activeRow 
			End If
		
		'제작사원			
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREEMPNAME") Then '25
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CREEMPNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, ""
				.txtCREEMPNO.value = ""
				.txtCREEMPNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow)
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREEMPNAME",.sprSht.ActiveRow)
				
				vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = vntData(0,1)  ' Code값 저장
					.txtCREDEPTNAME.value = vntData(1,1)  ' 코드명 표시
					.txtCREEMPNO.value = vntData(2,1)
					.txtCREEMPNAME.value = vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntData(3,1)
					mobjSCGLSpr.CellChanged .sprSht,38,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다 이거수
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
			
		'제작부서	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '23
			 
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, ""
				.txtCREDEPTCD.value = ""
				.txtCREDEPTNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow)
				
				
				vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = vntData(0,0)  ' Code값 저장
					.txtCREDEPTNAME.value = vntData(1,0)  ' 코드명 표시
					'msgbox vntData(0,0) 
					
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht,37,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
			
		'담당부서	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPTNAME") Then '14
		
			If mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, ""
				.txtDEPTCD.value = ""
				.txtDEPTNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow)
				
				
				vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,0)  ' Code값 저장
					.txtDEPTNAME.value = vntData(1,0)  ' 코드명 표시
					'msgbox vntData(0,0) 
					
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht,13,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If	
		'담당자		
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EMPNAME") Then '16
		
			If mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, ""
				.txtEMPNAME.value = ""
				.txtEMPNO.value = ""
				
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow)
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNAME",.sprSht.ActiveRow)
				
				vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,1)  ' Code값 저장
					.txtDEPTNAME.value = vntData(1,1)  ' 코드명 표시
					.txtEMPNO.value = vntData(2,1)
					.txtEMPNAME.value = vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntData(3,1)
					mobjSCGLSpr.CellChanged .sprSht,36,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다 이거수
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
		
		'크리조직
		ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			If mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, ""
				.txtEXCLIENTNAME.value = ""
				.txtEXCLIENTCODE.value = ""
				
			Else
				strCode		= ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
				
				If strCode = "" AND strCodeName <> "" Then			
					vntData = mobjSCCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "")

					If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
						If mlngRowCnt = 1 Then
							.txtEXCLIENTCODE.value = vntData(1,1)
							.txtEXCLIENTNAME.value = vntData(2,1)	
							mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(1,1)
							mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(2,1)			
							.txtFROM.focus
							.sprSht.focus
							If Row <> .sprSht.MaxRows Then
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
							Else
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
							End IF
						Else
							mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
							.txtFROM.focus
							.sprSht.focus 
							If Row <> .sprSht.MaxRows Then
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
							Else
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
							End IF
						End If
   					End If
   				End If
   			End If
		End If
	End with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
	
		'제작담당자
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '25
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"CREDEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(sprSht,"CREEMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtCREDEPTCD.value = vntRet(2,0)  ' Code값 저장
				.txtCREDEPTNAME.value = vntRet(3,0)  ' 코드명 표시
				.txtCREEMPNO.value = vntRet(0,0)
				.txtCREEMPNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
		
		'제작담당부서
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '23
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding(sprSht,"CREDEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtCREDEPTCD.value = vntRet(0,0)
				.txtCREDEPTNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
		
		'담당부서
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPTNAME") Then '14
			vntInParams = array(mobjSCGLSpr.GetTextBinding(sprSht,"DEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtDEPTCD.value = vntRet(0,0)
				.txtDEPTNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		'담당자
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EMPNAME") Then '16
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"DEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(sprSht,"EMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtDEPTCD.value = vntRet(2,0)  ' Code값 저장
				.txtDEPTNAME.value = vntRet(3,0)  ' 코드명 표시
				.txtEMPNO.value = vntRet(0,0)
				.txtEMPNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		'크리조직
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then '28
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				.txtEXCLIENTCODE.value = vntRet(1,0)
				.txtEXCLIENTNAME.value = vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)	
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		End if
		
	End With
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
	
		
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_DEPT")  Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_DEPT") then exit Sub
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
				.txtFROM.focus()
				gSetChangeFlag .txtDEPTCD
			end if
		
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EMP") then exit Sub
		
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
				
				.txtFROM.focus()
				'.txtMEMO.focus()					' 포커스 이동
				gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
				gSetChangeFlag .txtEMPNAME
				gSetChangeFlag .txtDEPTCD
				gSetChangeFlag .txtDEPTNAME
     		end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CDEPT") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CDEPT") then exit Sub
		
				vntInParams = array(trim(.txtCREDEPTNAME.value))
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
			if isArray(vntRet) then
				.txtCREDEPTCD.value = trim(vntRet(0,0))	'Code값 저장
				.txtCREDEPTNAME.value = trim(vntRet(1,0))	'코드명 표시
				if .sprSht.ActiveRow >0 Then	
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				.txtFROM.focus()
				gSetChangeFlag .txtCREDEPTCD
			end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CEMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CEMP") then exit Sub
		
			vntInParams = array(trim(.txtCREDEPTCD.value), trim(.txtCREDEPTNAME.value), trim(.txtCREEMPNO.value), trim(.txtCREEMPNAME.value)) '<< 받아오는경우
		
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if .txtCREEMPNO.value = vntRet(0,0) and .txtCREEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				.txtCREDEPTCD.value = trim(vntRet(2,0))  ' Code값 저장
				.txtCREDEPTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
				.txtCREEMPNO.value = trim(vntRet(0,0))
				.txtCREEMPNAME.value = trim(vntRet(1,0))
				
				if .sprSht.ActiveRow >0 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
					
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
					
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				
				.txtFROM.focus()					' 포커스 이동
				gSetChangeFlag .txtCREEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
				gSetChangeFlag .txtCREEMPNAME
				gSetChangeFlag .txtCREDEPTCD
				gSetChangeFlag .txtCREDEPTNAME
     		end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EXCLIENTCODE") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				.txtEXCLIENTCODE.value = vntRet(1,0)	
				.txtEXCLIENTNAME.value = vntRet(2,0)	
				if .sprSht.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)	
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				End If		
			
			End If
			.txtFROM.focus()					' 포커스 이동
			gSetChangeFlag .txtEXCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEXCLIENTNAME
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		End if
	
		
		
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
	
	'키가 움질일때 바인딩
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT") Then
				strCOLUMN = "BUDGETAMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT")  Then
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


Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	Dim vntData_DEPT , vntData_RATE
   	Dim strJOBNO , strRow
   	
	mstrBindCHK = True
	with frmThis
	
		if .sprSht.MaxRows = 0 then exit function '그리드 데이터가 없으면 나간다.
		
		.txtPROJECTNM.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNM",Row)
		.txtPROJECTNO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNO",Row) 
		.txtCLIENTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtCPDEPTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CPDEPTNAME",Row)
		.txtCREDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREDAY",Row)
		.txtCPEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CPEMPNAME",Row)
		.txtGROUPGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"GROUPGBN",Row)
		.txtSUBSEQNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtCLIENTTEAMNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTTEAMNAME",Row)
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		.txtJOBNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",Row)
		.cmbJOBGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row) 
		'.txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row) 
		Call SUBCOMBO_TYPE()
		.txtDEPTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEPTNAME",Row)
		.txtDEPTCD.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEPTCD",Row)
		.txtREQDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",Row)
		.cmbCREPART.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREPART",Row) 
		'Call Get_SUBCOMBO_VALUE(mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row))
		 
		.txtEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNAME",Row) 
		.txtEMPNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",Row)
		.txtHOPEENDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"HOPEENDDAY",Row)
		.cmbCREGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREGUBN",Row) 
		.cmbJOBBASE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBBASE",Row) 
		.txtCREDEPTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREDEPTNAME",Row) 
		.txtCREDEPTCD.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREDEPTCD",Row)
		.cmbENDFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",Row) 
		.txtCREEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREEMPNAME",Row)
		.txtCREEMPNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREEMPNO",Row)
		.txtAGREEYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AGREEYEARMON",Row) 
		.txtDEMANDYEARMON.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDYEARMON",Row)
		.txtSETYEARMON.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"SETYEARMON",Row)
		.txtBUDGETAMT.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT",Row)
		.txtBIGO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row)
		.txtEXCLIENTCODE.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		.txtEXCLIENTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
  		If .txtBUDGETAMT.value <> "" Then
			txtBUDGETAMT_onblur
		End If
		'If .cmbENDFLAG.value = "PF01" Or .cmbENDFLAG.value = "PF02" Then 
		'	.cmbENDFLAG.disabled = false
		'Else 
			.cmbENDFLAG.disabled = true
		'End If
		
		
		'분배비율 여부를 미리 알려준다.
		strJOBNO = .txtJOBNO.value		
		.chkDEPT.checked = FALSE
		.chkRATE.checked = FALSE
		
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData_DEPT = mobjPDCMACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,"")
		IF mlngRowCnt > 0 then
			.chkDEPT.checked = TRUE
		end If
		
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData_RATE = mobjPDCMACTUALRATE.SelectRtn_DTL_ACTUALRATE(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,"")
		If mlngRowCnt > 0  then
			.chkRATE.checked = TRUE
		end IF
		
		.txtFROM.focus()
		.sprSht.focus()	
			
   	end with
  
	mstrBindCHK = False
End Function
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML><XML id="xmlBind1"></XML>
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
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;제작&nbsp;관리 <!--<span id="spnSELECTHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_SELECTTBL_HIDDEN ()">
														(숨기기)</span>--></td>
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
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="화면을 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblSelectBody" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="top">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand; HEIGHT: 22px" onclick="vbscript:Call DateClean()"
													width="112">의뢰일&nbsp;검색</TD>
												<TD class="SEARCHDATA" style="WIDTH: 230px; HEIGHT: 16.72pt" width="230"><INPUT class="INPUT" id="txtFROM" title="의뢰일 검색(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
														accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="의뢰일 검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 22px" onclick="vbscript:Call gCleanField(txtSEARCHCLIENTSUBNAME, txtSEARCHCLIENTSUBCODE)"
													width="90">팀</TD>
												<TD class="SEARCHDATA" style="WIDTH: 229px; HEIGHT: 16.72pt" width="229"><INPUT class="INPUT_L" id="txtSEARCHCLIENTSUBNAME" title="사업부명 조회" style="WIDTH: 149px; HEIGHT: 22px"
														type="text" maxLength="100" size="18" name="txtSEARCHCLIENTSUBNAME"><IMG id="ImgSEARCHCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgSEARCHCLIENTSUBCODE"><INPUT class="INPUT" id="txtSEARCHCLIENTSUBCODE" title="사업부코드 조회" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtSEARCHCLIENTSUBCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 88px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtSEARCHCLIENTCODE, txtSEARCHCLIENTNAME)"
													width="88">광고주</TD>
												<TD class="SEARCHDATA" style="HEIGHT: 18.24pt"><INPUT class="INPUT_L" id="txtSEARCHCLIENTNAME" title="광고주명 조회" style="WIDTH: 131px; HEIGHT: 22px"
														type="text" maxLength="100" size="16" name="txtSEARCHCLIENTNAME"><IMG id="ImgSEARCHCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgSEARCHCLIENTCODE"><INPUT class="INPUT" id="txtSEARCHCLIENTCODE" title="광고주코드 조회" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtSEARCHCLIENTCODE"></TD>
												<td class="SEARCHDATA" style="HEIGHT: 18.24pt" width="53"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
														align="right" border="0" name="imgQuery"></td>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand" onclick="vbscript:Call JOBGUBNClean()"
													width="112">매체부문</TD>
												<TD class="SEARCHDATA" style="WIDTH: 230px" width="230"><SELECT id="cmbSEARCHJOBGUBN" title="매체부문조회" style="WIDTH: 223px" name="cmbSEARCHJOBGUBN"></SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call ENDFLAGClean()"
													width="90">완료구분</TD>
												<TD class="SEARCHDATA" style="WIDTH: 229px" width="229"><SELECT id="cmbSEARCHENDFLAG" title="완료구분" style="WIDTH: 227px" name="cmbSEARCHENDFLAG"></SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 88px; CURSOR: hand" width="88"><SELECT id="cmbPOPUPTYPE" title="프로젝트,JOBNO선택" style="WIDTH: 88px" name="cmbPOPUPTYPE">
														<OPTION value="1" selected>PROJECT</OPTION>
														<OPTION value="2">JOBNO</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtPROJECTNM1" title="제작건명 조회" style="WIDTH: 192px; HEIGHT: 22px"
														type="text" maxLength="100" size="26" name="txtPROJECTNM1"><IMG id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgPROJECTNO1"><INPUT class="INPUT" id="txtPROJECTNO1" title="제작번호 조회" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="7" size="4" name="txtPROJECTNO1"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<tr>
						<td>
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE style="WIDTH: 100%; HEIGHT: 8px" height="8" cellSpacing="0" cellPadding="0" width="100%"
								background="../../../images/TitleBG.gIF" border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table style="WIDTH: 640px; HEIGHT: 26px" cellSpacing="0" cellPadding="0" width="640" border="0">
											<tr>
												<td class="TITLE"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id="imgTableUp" style="CURSOR: hand" alt="자료를 검색합니다." src="../../../images/imgTableUp.gif"
															align="absMiddle" border="0" name="imgTableUp"></span> &nbsp;JOB 
													관리&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
														accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">&nbsp;&nbsp; 
													선택합계 : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
														readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
														src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
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
						<TD class="BODYSPLIT" id="spacebar1" style="WIDTH: 100%; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%" vAlign="top" align="left">
							<TABLE class="DATA" id="tblBody1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="GROUP" width="20" rowSpan="3">기<BR>
										본<BR>
										정<BR>
										보
									</TD>
									<TD class="LABEL" width="90">프로젝트명</TD>
									<TD class="DATA" width="230"><INPUT dataFld="PROJECTNM" class="NOINPUTB_L" id="txtPROJECTNM" title="프로젝트명" style="WIDTH: 160px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="21" name="txtPROJECTNM">&nbsp;<INPUT dataFld="PROJECTNO" class="NOINPUTB" id="txtPROJECTNO" title="프로젝트건" style="WIDTH: 65px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="6" name="txtPROJECTNO"></TD>
									<TD class="LABEL" width="90">브랜드</TD>
									<TD class="DATA" width="230"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_L" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="24" name="txtSUBSEQNAME"></TD>
									<TD class="LABEL" width="90">담당부서 [CP]</TD>
									<TD class="DATA"><INPUT dataFld="CPDEPTNAME" class="NOINPUTB_L" id="txtCPDEPTNAME" title="담당부서 CP" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCPDEPTNAME"></TD>
								<TR>
									<TD class="LABEL">등록일</TD>
									<TD class="DATA"><INPUT dataFld="CREDAY" class="NOINPUTB" id="txtCREDAY" title="등록일" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCREDAY"></TD>
									<TD class="LABEL">팀</TD>
									<TD class="DATA"><INPUT dataFld="CLIENTTEAMNAME" class="NOINPUTB_L" id="txtCLIENTTEAMNAME" title="팀" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTTEAMNAME"></TD>
									<TD class="LABEL">담당자 [CP]</TD>
									<TD class="DATA"><INPUT dataFld="CPEMPNAME" class="NOINPUTB_L" id="txtCPEMPNAME" title="담당자CP" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCPEMPNAME"></TD>
								</TR>
								<TR>
									<TD class="LABEL">그룹구분</TD>
									<TD class="DATA"><INPUT dataFld="GROUPGBN" class="NOINPUTB_L" id="txtGROUPGBN" title="그룹구분" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtGROUPGBN"></TD>
									<TD class="LABEL">광고주</TD>
									<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
									<TD class="LABEL">비고</TD>
									<TD class="DATA"><INPUT dataFld="MEMO" class="NOINPUTB_L" id="txtMEMO" title="메모" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtMEMO"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" id="spacebar2" style="WIDTH: 100%; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%" vAlign="top" align="left">
							<TABLE class="DATA" id="tblBody2" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="GROUP" width="20" rowSpan="7"><BR>
										부<BR>
										가<BR>
										정<BR>
										보<BR>
									</TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtJOBNAME, '')"
										width="90">JOB명</TD>
									<TD class="DATA" width="230"><INPUT dataFld="JOBNAME" id="txtJOBNAME" title="제작건명" style="WIDTH: 164px; HEIGHT: 21px"
											accessKey=",M" dataSrc="#xmlBind" type="text" size="21" name="txtJOBNAME"><INPUT dataFld="JOBNO" class="NOINPUT" id="txtJOBNO" title="제작관번호코드" style="WIDTH: 65px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBNO"></TD>
									<TD class="LABEL" width="90">매체부문</TD>
									<TD class="DATA" width="230"><SELECT dataFld="JOBGUBN" id="cmbJOBGUBN" title="매체구분" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBGUBN"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPTNAME, txtDEPTCD)"
										width="90">담당팀</TD>
									<TD class="DATA"><INPUT dataFld="DEPTNAME" class="INPUT_L" id="txtDEPTNAME" title="담당부서명" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtDEPTNAME"><IMG id="ImgDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgDEPTCD"><INPUT dataFld="DEPTCD" class="INPUT_L" id="txtDEPTCD" title="담당부서코드" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="6" name="txtDEPTCD"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREQDAY, '')">의뢰일</TD>
									<TD class="DATA"><INPUT dataFld="REQDAY" class="INPUT" id="txtREQDAY" title="의뢰일" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtREQDAY"><IMG id="imgCalEndarREQ" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
											name="imgCalEndarREQ"></TD>
									<TD class="LABEL">매체분류</TD>
									<TD class="DATA"><SELECT dataFld="CREPART" id="cmbCREPART" title="매체분류" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREPART"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEMPNAME, txtEMPNO)">담당자</TD>
									<TD class="DATA"><INPUT dataFld="EMPNAME" class="INPUT_L" id="txtEMPNAME" title="담당자명" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtEMPNAME"><IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgEMPNO"><INPUT dataFld="EMPNO" class="INPUT_L" id="txtEMPNO" title="담당자사번" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtEMPNO"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtHOPEENDDAY, '')">완료예정일</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><INPUT dataFld="HOPEENDDAY" class="INPUT" id="txtHOPEENDDAY" title="완료예정일" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtHOPEENDDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndar"></TD>
									<TD class="LABEL" style="HEIGHT: 24px">신규/수정 구분</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><SELECT dataFld="CREGUBN" id="cmbCREGUBN" title="신규/수정 구분" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREGUBN"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtCREDEPTNAME,txtCREDEPTCD)">제작담당팀</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><INPUT dataFld="CREDEPTNAME" class="INPUT_L" id="txtCREDEPTNAME" title="제작담당자부서명" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtCREDEPTNAME"><IMG id="ImgCREDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
											name="ImgCREDEPTCD"><INPUT dataFld="CREDEPTCD" class="INPUT_L" id="txtCREDEPTCD" title="제작담당자부서코드" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCREDEPTCD"></TD>
								</TR>
								<TR>
									<TD class="LABEL">완료구분</TD>
									<TD class="DATA"><SELECT dataFld="ENDFLAG" id="cmbENDFLAG" title="완료구분" style="WIDTH: 112px" dataSrc="#xmlBind"
											name="cmbENDFLAG"></SELECT>&nbsp;&nbsp;<IMG id="imgEndChange" onmouseover="JavaScript:this.src='../../../images/imgEndChangeOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgEndChange.gIF'" height="20" alt="진행상태 를 의뢰상태로 변경합니다."
											src="../../../images/imgEndChange.gIF" align="absMiddle" border="0" name="imgEndChange"></TD>
									<TD class="LABEL">정산대상</TD>
									<TD class="DATA"><SELECT dataFld="JOBBASE" id="cmbJOBBASE" title="정산대상" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBBASE"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCREEMPNAME, txtCREEMPNO)">제작담당자</TD>
									<TD class="DATA"><INPUT dataFld="CREEMPNAME" class="INPUT_L" id="txtCREEMPNAME" title="제작담당자" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtCREEMPNAME"><IMG id="ImgCREEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
											name="ImgCREEMPNO"><INPUT dataFld="CREEMPNO" class="INPUT_L" id="txtCREEMPNO" title="제작담당사번" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCREEMPNO"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtBUDGETAMT, '')">예산</TD>
									<TD class="DATA"><INPUT dataFld="BUDGETAMT" class="INPUT_R" id="txtBUDGETAMT" title="예산금액" style="WIDTH: 224px; HEIGHT: 21px"
											accessKey="NUM" dataSrc="#xmlBind" type="text" size="32" name="txtBUDGETAMT"></TD>
									<TD class="LABEL" style="CURSOR: hand">실적분배율</TD>
									<TD class="DATA">
										<TABLE class="NOINPUTB" id="Table1" style="WIDTH: 224px; HEIGHT: 27px" cellSpacing="1"
											cellPadding="0">
											<TR>
												<TD class="DATA" style="WIDTH: 54px" align="left" width="54">담당<INPUT id="chkDEPT" disabled tabIndex="0" type="checkbox" name="chkDEPT"></TD>
												<TD class="DATA" style="WIDTH: 46px" align="left" width="46">분배<INPUT id="chkRATE" disabled type="checkbox" name="chkRATE"></TD>
												<TD class="DATA"><IMG id="ImgDivamtPop" onmouseover="JavaScript:this.src='../../../images/ImgDivamtPopOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgDivamtPop.gIF'" height="20"
														alt="신규자료를 작성합니다." src="../../../images/ImgDivamtPop.gif" border="0" name="ImgDivamtPop"></TD>
											</TR>
										</TABLE>
									</TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)">크리조직</TD>
									<TD class="DATA"><INPUT dataFld="EXCLIENTNAME" class="INPUT_L" id="txtEXCLIENTNAME" title="대대행사명" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="24" name="txtEXCLIENTNAME"><IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE"><INPUT dataFld="EXCLIENTCODE" class="INPUT_L" id="txtEXCLIENTCODE" title="대대행사코드" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="8" size="6" name="txtEXCLIENTCODE">
									</TD>
								</TR>
								<TR>
									<TD class="LABEL">합의월</TD>
									<TD class="DATA"><INPUT dataFld="AGREEYEARMON" class="NOINPUTB" id="txtAGREEYEARMON" title="합의월" style="WIDTH: 224px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="32" name="txtAGREEYEARMON"></TD>
									<TD class="LABEL">청구월</TD>
									<TD class="DATA"><INPUT dataFld="DEMANDYEARMON" class="NOINPUTB" id="txtDEMANDYEARMON" title="청구월" style="WIDTH: 224px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="32" name="txtDEMANDYEARMON"></TD>
									<TD class="LABEL">결산월</TD>
									<TD class="DATA"><INPUT dataFld="SETYEARMON" class="NOINPUTB" id="txtSETYEARMON" title="결산월" style="WIDTH: 264px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="38" name="txtSETYEARMON"></TD>
								</TR>
								<TR>
									<TD class="LABEL" onclick="vbscript:Call CleanField(txtBIGO, '')">비고</TD>
									<TD class="DATA" colSpan="5"><INPUT dataFld="BIGO" id="txtBIGO" title="부가정보 비고" style="WIDTH: 920px; HEIGHT: 21px" dataSrc="#xmlBind"
											type="text" size="148" name="txtBIGO"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--Input End-->
					<!--BodySplit Start-->
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 98%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 98%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="42413">
									<PARAM NAME="_ExtentY" VALUE="6826">
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
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
