<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDEMANDCONFIRM_USER.aspx.vb" Inherits="PD.PDCMDEMANDCONFIRM_USER" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>청구요청</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMDEMAND.aspx
'기      능 : SpreadSheet를 이용한 청구요청/JOB분할/조회 의 기능을 가진다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/10 By KimTH
'****************************************************************************************
-->
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet 정보 --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt			'데이터의 로우및 컬럼 반환
Dim mobjPDCODEMAND					'청구요청 의 Control Class
Dim mobjPDCOGET						'제작공통 Control Class
Dim mobjSCCOGET						'전체공통 Control Class
Dim mstrCheck						'전체 선택 및 해제 구분자
Dim mstrSelect						'조회구분 (저장후 이력조회 Or 최초 입력대상 조회)
Dim mlngRowChk						'하단그리드 저장발리데이션 사용
Dim mstrDEPTCD						'로그인사용자부서


Dim mlngTaxRowCnt
Dim mlngTaxColCnt
Const meTab = 9
mstrCheck = True					'전체선택은 최초 해제	
mstrSelect = false					'조회구분 Default Value: 입력대상 조회

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgDivDemand_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' 명령버튼
'=========================================================================================
Sub imgQuery_onclick
	with frmThis
		
	End with
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'인쇄 - 해당사항 없음
Sub imgPrint_onclick ()
	
End Sub	
Sub imgAgree_onclick
	Data_Confirm("3")
	SelectRtn
	
End Sub

Sub imgAgreeCanCel_onclick
	Data_Confirm("2")
	SelectRtn
End Sub

Sub imgBackProc_onclick
	Data_Confirm("0")
	SelectRtn
End Sub

Sub Chk_False
	Dim intCnt
	with frmThis
		If .sprSht.MaxRows <> 0 Then
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt, "0"	
		Next
		End If
	End with
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDelUp_onclick
	gFlowWait meWAIT_ON
	DeleteRtnProc
	gFlowWait meWAIT_OFF

End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' SpreadSheet 이벤트 
'=========================================================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intRtn
	Dim dblChk
	Dim dblChkSum
	Dim vntData
	Dim intRtnChk
	'mlngRowChk
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if
		'전체클릭시 셀데이터 변경 반영
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
		
		
	end with	
End Sub


Sub sprSht_Change(ByVal Col, ByVal Row)	
	
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
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
		'필드 To 바인딩 존재시 기입
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
				strCOLUMN = "DIVAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
				strCOLUMN = "CHARGE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
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
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	Dim lngChCnt,IntChAMT,IntChAMTSUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtDIVAMT.value = 0
		else
			.txtDIVAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtDIVAMT,0,True)
		End If
		
		IntEXEAMTSUM = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtADJAMT.value = 0
		else
			.txtADJAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtADJAMT,0,True)
		End If
		
		IntChAMTSUM = 0
		For lngChCnt = 1 To .sprSht.MaxRows
			IntChAMT = 0	
			IntChAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHARGE", lngChCnt)
			IntChAMTSUM = IntChAMTSUM + IntChAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtCHARGE.value = 0
		else
			.txtCHARGE.value = IntChAMTSUM
			Call gFormatNumber(frmThis.txtCHARGE,0,True)
		End If
	End With
End Sub




sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO,strPREESTNO,strPRIJOBNAME,strPROJECTNM,strJOBNAME
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	Dim strCLIENTCODE,strCLIENTNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strTIMCODE,strTIMNAME,strSUBSEQ,strSUBSEQNAME,strJOBGUBN,strJOBGUBNNAME,lngCOMMITIONVALUE
	Dim vntInParams
	Dim vntRet
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
		
			strWith =  Screen.width
			strHeight =  Screen.height - 100
			
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strSUBNO= mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",.sprSht.ActiveRow)
			strJOBNAME	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",.sprSht.ActiveRow)	
			strPREESTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",.sprSht.ActiveRow)		
			strPRIJOBNAME = mobjSCGLSpr.GetTextBinding( .sprSht,"PRIJOBNAME",.sprSht.ActiveRow)	
			strPROJECTNM = mobjSCGLSpr.GetTextBinding( .sprSht,"PROJECTNM",.sprSht.ActiveRow) 
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow) 
			strJOBGUBNNAME  = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBNNAME",.sprSht.ActiveRow) 
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow)	 
			strTIMCODE =  mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",.sprSht.ActiveRow)	
			strSUBSEQ =  mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",.sprSht.ActiveRow)
			strJOBGUBN =  mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBN",.sprSht.ActiveRow)	
			
			vntInParams = array(strJOBNO,strSUBNO,strJOBNAME,strPREESTNO,strPRIJOBNAME,strPROJECTNM,strCLIENTNAME,strJOBGUBNNAME,strCLIENTCODE,strTIMCODE,strSUBSEQ,strJOBGUBN)
			vntRet = gShowModalWindow("PDCMJOBMST.aspx",vntInParams , strWith,strHeight)
			
		End If
	End With
end sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)

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
		sprSht_Click frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================

' 페이지 화면 디자인 및 초기화 
Sub InitPage()
	'서버업무객체 생성	
	set mobjPDCODEMAND	= gCreateRemoteObject("cPDCO.ccPDCODEMAND")
	set mobjPDCOGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
		'=========================================================================================
		'청구요청 SHEET 'CHK|YEARMON|JOBNO|SEQ|PREESTNO
		'=========================================================================================
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 30, 0
		mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|DATAYEARMON|YEARMON|JOBNAME|JOBNO|SEQ|PREESTNO|CLIENTNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|DEPTNAME|EMPNAME|DEMANDPERSON|MANAGERNAME|RANKDIV|USENO|MANAGER|CHARGEHISTORY|DELCHK|CLIENTCODE|TIMCODE|SUBSEQ|JOBGUBN|PRIJOBNAME|PROJECTNM|JOBGUBNNAME"
		mobjSCGLSpr.SetHeader .sprSht,		  "선택|승인요청월|청구요청월|JOB명|JOBno.|SUBno.|견적번호|광고주명|견적금액|청구금액|차액|청구기준|내역|청구방법|담당부서|담당자|요청자|승인자|색구분|요청자사번|승인자사번|요청SEQ|반려구분|광고주코드|팀코드|브랜드코드|매체구분|대표좝명|프로젝트명|매체구분명"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|10        |10        |20   |8     |6     |0       |22      |11      |11      |11  |10      |10  |10      |12      |8     |8     |8     |0     |0         |0         |10     |10      |0         |0     |0         |0       |0       |0         |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT|CHARGE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "DATAYEARMON|YEARMON|JOBNAME|JOBNO|SEQ|PREESTNO|CLIENTNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|DEPTNAME|EMPNAME|DEMANDPERSON|MANAGERNAME|MANAGER|CHARGEHISTORY|DELCHK|CLIENTCODE|TIMCODE|SUBSEQ|JOBGUBN|PRIJOBNAME|PROJECTNM|JOBGUBNNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DATAYEARMON|YEARMON|JOBNO|SEQ|DEMANDFLAG|TAXCODE|EMPNAME|DEMANDPERSON|MANAGERNAME|MEMO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|DEPTNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|RANKDIV|USENO|MANAGER", true
		.sprSht.style.visibility = "visible"
		
		.rdT.style.display = "none"
		.rdF.style.display = "none"
		.imgBackProc.style.display = "none"
    End With
	
	InitPageData	
	'SelectRtn
End Sub

Sub EndPage()
	set mobjPDCODEMAND = Nothing
	set mobjPDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


' 화면의 초기상태 데이터 설정

Sub InitPageData
	'모든 데이터 클리어
	Dim vntData
	
	gClearAllObject frmThis
	'초기 데이터 설정
	with frmThis
		.sprSht.maxrows = 0
		.txtYEARMON.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2) '추후 이것으로 대처 임시로 테스트값 연결 하였음
		'.txtYEARMON.value = "200910"

	vntData = mobjPDCODEMAND.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
		if mlngRowCnt > 0 Then
		mstrDEPTCD = vntData(0,1)
		end if
   	end if	
	
	rdChecked
	End with
	'새로운 XML 바인딩을 생성
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub rdChecked
	with frmThis
		If .rdT.checked = True Then
			.imgAgreeCanCel.style.display = "none"
			.imgAgree.style.display = "inline"
			'.imgBackProc.style.display = "inline"
		Else
			.imgAgree.style.display = "none"
			.imgBackProc.style.display = "none"
			.imgAgreeCanCel.style.display = "inline"
		End If
	End with

End Sub
Sub rdT_onclick
	rdChecked
	SelectRtn
End Sub
Sub rdF_onclick
	rdChecked
	SelectRtn
End Sub

' 그리드콤보
'자동콤보 변경
Sub Get_COMBO_PVALUE (ByVal blnRow)		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",blnRow,blnRow,vntData_TaxCode,,80,,true
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	
'상단그리드 콤보
Sub Get_COMBO_UPVALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DEMANDFLAG",,,vntData_Demand,,80	
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",,,vntData_TaxCode,,80						
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		
'하단그리드콤보
Sub Get_COMBO_VALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "DEMANDFLAG",,,vntData_Demand,,80		
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "TAXCODE",,,vntData_TaxCode,,80

			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

'****************************************************************************************
' 데이터 처리 
'****************************************************************************************
'행추가
Sub imgRowAdd_onclick ()
	with frmThis
		If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI04" Then 
			call sprSht1_Keydown(meINS_ROW, 0)
			mlngRowChk = .sprSht.ActiveRow
		Else 
			gErrorMsgBox "상단선택된 데이터 청구기준이 분할내역 대상이 아닙니다." & vbcrlf & "청구기준을 확인하십시오.","행추가처리안내"
		End If
	End with
End Sub


Sub sprSht1_Keydown(KeyCode, Shift)

	Dim intRtn
	
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	'if KeyCode = meCR  Or KeyCode = meTab Then
	'	if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht1,"SAVEFLAG")  Then ' 예제 frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DETAIL_BTN")
	'		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(13), cint(Shift), -1, 1)
	'		DefaultValue
	'	End If
	'Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select

	'End If
End Sub

'신규디펄드 값을 생성
Sub DefaultValue
	
	
End Sub
'조회
Sub SelectRtn
	Dim vntData
	Dim intCnt
	Dim strGbn

	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		mlngTaxRowCnt=clng(0)
		mlngTaxColCnt=clng(0)
		
		If .rdT.checked = True Then
			strGbn = "2"
		Else
			strGbn = "3"
		End If
		
		vntData = mobjPDCODEMAND.SelectRtn_DEMANDPRECONFIRM(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,strGbn)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows 
					'JOB별 컬러 통일
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKDIV",intCnt) Mod 2 = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
				Next
				
				Chk_False
   			Else
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		window.setTimeout "AMT_SUM",1	
		.txtSELECTAMT.value = 0
   	end with
	
End Sub


' 저장
Sub ProcessRtn ()
	
End Sub

Sub Data_Confirm(byVal strConfirmFlag)
	Dim vntData
	Dim intRtn
	Dim intCnt
	Dim strMSG
	Dim intSaveRtn
	Dim intCnt3
	Dim intCnt4
	Dim strJOBNAME
	Dim strSEQ
	with frmThis
		If strConfirmFlag = "3" Then
			For intCnt3 = 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt3) = "1" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"MANAGER",intCnt3,gstrUsrID
				End If
			Next
		End If
		'반려 일경우 반려 불가능 항목은 PD_DEMANDRETURN_FUN 으로 인해 Y 로 표기되며,거래명세서 또는 세금계산서, 전표 삭제를 해야 반려가 가능 하다.
		If strConfirmFlag = "0" Then
			For intCnt4= 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt4) = "1" Then 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DELCHK",intCnt4) = "Y" Then
						strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt4)
						strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",intCnt4)
						gErrorMsgBox "JOBNO [" & strJOBNAME & "-" & strSEQ & "] 데이터 는 반려대상이 아닙니다." & vbcrlf & "관련된 차액이월 분 의 거래명세서 를 확인 하십시오.","반려처리안내!"
						Exit Sub
					End If
				End If
			Next
		End If
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|JOBNO|SEQ|PREESTNO|USENO|MANAGER")
		
		
		
		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 데이터가 없습니다.","승인처리 안내"
		End If
		if  not IsArray(vntData)  then
			gErrorMsgBox "변경된 " & meNO_DATA,"승인처리 안내"
			Exit Sub
		End If
		
		select case strConfirmFlag
			case "2": strMSG = "승인취소"
			case "3": strMSG = "승인"
			case "0": strMSG = "반려"
		end select
		
		intSaveRtn = gYesNoMsgbox("해당데이터를 " & strMSG & "하시겠습니까?","청구요청 확인")
		
		IF intSaveRtn <> vbYes then exit Sub
		
		
		intRtn = mobjPDCODEMAND.Data_Confirm(gstrConfigXml,vntData,strConfirmFlag)
		'저장후 sms 발송
		Call SMS_SEND()
		
		if not gDoErrorRtn ("Data_Confirm") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "자료가 " & strMSG & mePROC_DONE,"처리안내" 		
		End If
	End with

End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim dblSumAmt
   	Dim dblAMT
	'On error resume next
	with frmThis
  	
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		dblSumAmt = 0
		
   		for intCnt = 1 to .sprSht1.MaxRows
   			'Sheet 필수 입력사항
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTCODE",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"YEARMON",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 내용 기입여부 를 확인하십시오","저장오류"
				Exit Function
			End if
			dblAMT = 0
			dblAMT = mobjSCGLSpr.getTextBinding(.sprSht1,"DIVAMT",intCnt) 
			dblSumAmt = dblSumAmt + dblAMT
			'금액 오류사항
		next
   		If mobjSCGLSpr.getTextBinding(.sprSht,"DIVAMT",.sprSht.ActiveRow) < dblSumAmt Then
   			gErrorMsgBox "분할대상금액의 합은 견적금액 을 초과할수 없습니다","저장오류"
   			Exit Function
   		End If
   	End with
   	
	DataValidation = true
End Function



-->
		</script>
		<script language="javascript">
		//SMS 발송
		function SMS_SEND(){
			frmSMS.location.href = "PD_SMS.asp"; 
		}
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0">
				<TR valign="top">
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="76" background="../../../images/back_p.gIF"
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
											<td class="TITLE">청구요청접수&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 326px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="청구요청 월 을 삭제 합니다." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="80">승인요청 월</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtYEARMON" title="등록월" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHDATA">&nbsp;<INPUT id="rdT" title="요청내역조회" type="radio" CHECKED value="rdT" name="rdGBN">
												&nbsp; <INPUT id="rdF" title="승인내역조회" type="radio" value="rdF" name="rdGBN"></TD>
											<td align="right" ><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery">
												<IMG id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
													height="20" alt="선택한 행을 승인합니다." src="../../../images/imgAgree.gIF" align="absMiddle"
													border="0" name="imgAgree"> <IMG id="imgBackProc" onmouseover="JavaScript:this.src='../../../images/imgBackProcOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgBackProc.gIF'" height="20" alt="선택한 행을 반려합니다."
													src="../../../images/imgBackProc.gIF" align="absMiddle" border="0" name="imgBackProc">
												<IMG id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'"
													height="20" alt="선택한 행을 승인취소 합니다." src="../../../images/imgAgreeCanCel.gIF" align="absMiddle"
													border="0" name="imgAgreeCanCel"> <IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"
													align="absMiddle" border="0" name="imgExcel">
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR valign="top">
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="굴림"></FONT></TD>
				</TR>
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td class="TITLE">합 계 : <INPUT class="NOINPUTB_R" id="txtDIVAMT" title="견적금액합계" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtDIVAMT"> <INPUT class="NOINPUTB_R" id="txtADJAMT" title="청구금액합계" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtADJAMT">&nbsp;<INPUT class="NOINPUTB_R" id="txtCHARGE" title="잔액합계" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtCHARGE">&nbsp;<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
							VIEWASTEXT>
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="40323">
							<PARAM NAME="_ExtentY" VALUE="16325">
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
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus"></TD>
				</TR>
			</TABLE>
			</TD></TR></TBODY></TABLE></form>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 0px;HEIGHT: 0px" name="frmSMS"></iframe> <!--DISPLAY: none; -->
	</body>
</HTML>
