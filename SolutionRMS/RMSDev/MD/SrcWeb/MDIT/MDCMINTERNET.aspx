<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNET.aspx.vb" Inherits="MD.MDCMINTERNET" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>인터넷 청약관리</title>
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDITINTERNETREG, mobjMDCOGET
Dim mstrCheck, mstrCheck1
CONST meTAB = 9
mstrCheck=True
mstrCheck1=True
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
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

sub ImgAddRow_onclick ()
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "상단의 캠페인 정보가 없으면 추가할 수 없습니다.","저장안내"
			Exit Sub
		End If
		
		call sprSht_DTL_Keydown(meINS_ROW, 0)
		.txtCAMPAIGN_CODE1.focus
		.sprSht_DTL.focus
	End With 
End sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
Sub ImgSave_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "상세항목 이 없습니다.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' 캠페인 팝업[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCAMPAIGN_CODE1_onclick
	Call CAMPAIGN_POP()
End Sub

'실제 데이터List 가져오기
Sub CAMPAIGN_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtYEARMON1.value, .txtCAMPAIGN_CODE1.value, .txtCAMPAIGN_NAME1.value)
			
		vntRet = gShowModalWindow("MDCMINTERNETCAMPAIGNPOP.aspx",vntInParams , 520,525)
			
		if isArray(vntRet) then
			.txtYEARMON1.value = vntRet(0,0)	
			.txtCAMPAIGN_CODE1.value = vntRet(1,0)
			.txtCAMPAIGN_NAME1.value = vntRet(2,0)
			SelectRtn
		end if
		
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCAMPAIGN_NAME1_onkeydown
	if window.event.keyCode = meEnter Or window.event.keyCode = meTab then
		Dim vntData
   		Dim i, strCols
		
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCAMPAIGN_INFO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtCAMPAIGN_CODE1.value,.txtCAMPAIGN_NAME1.value)

			if not gDoErrorRtn ("GetCAMPAIGN_INFO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtCAMPAIGN_CODE1.value = vntData(1,1)
					.txtCAMPAIGN_NAME1.value = vntData(2,1)
					selectRtn           
				Else
					Call CAMPAIGN_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if  Row > 0 AND Col > 1 then
			SelectRtn_DTL Col, Row
		end if
	end with
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
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
	end with
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
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
End Sub

sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	With frmThis
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
		if KeyCode = meCR  Or KeyCode = meTab Then
			if .sprSht_DTL.ActiveRow = .sprSht_DTL.MaxRows and .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"TRU_TAX_FLAG") Then
				intRtn = mobjSCGLSpr.InsDelRow(.sprSht_DTL, cint(13), cint(Shift), -1, 1)
				if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"YEARMON",.sprSht_DTL.ActiveRow-1) <> "" and .sprSht_DTL.MaxRows > 1 then
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"YEARMON",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"YEARMON",.sprSht_DTL.ActiveRow-1)
				else
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"YEARMON",.sprSht_DTL.ActiveRow, Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
				end if 
				
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"VOCH_TYPE",.sprSht_DTL.ActiveRow, "0"
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTNAME",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSURATE",.sprSht_DTL.ActiveRow, 100-mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MCCOMMI_RATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSURATE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MCCOMMI_RATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMI_RATE",.sprSht_DTL.ActiveRow, 20
				
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CAMPAIGN_CODE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CAMPAIGN_CODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TIMCODE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TIMCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"SUBSEQ",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SUBSEQ",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TBRDSTDATE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDSTDATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TBRDEDDATE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDEDDATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"DEPT_CD",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_CD",.sprSht_HDR.ActiveRow)
			End if
		ElseIf KeyCode = meINS_ROW  Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_DTL, meINS_ROW, 0, -1, 1)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"VOCH_TYPE",.sprSht_DTL.ActiveRow, "0"
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CAMPAIGN_CODE",.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CAMPAIGN_CODE",.sprSht_HDR.ActiveRow)
		End if
		
	End With
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
		If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") or _
			.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MCSUSU") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXSUSU") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MCSUSU")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXSUSU")) Then
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
	End With
End Sub

Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_DTL.MaxRows >0 Then
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") or _
			.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MCSUSU") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXSUSU") Then
						
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
	End With
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
   	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim lngPrice
	Dim lngVALUE
	Dim lngVALUE1
	Dim lngVALUE2
	Dim strAMT
	Dim strCOMMI_RATE
	Dim strCOMMISSION
	Dim lngMCSUSU
		
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"YEARMON") Then	
			Dim strdate
			strdate = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"YEARMON",Row)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"YEARMON",Row, strdate
			DateClean_SHEET strdate, Row
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MEDNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMEDCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", _
													  strCode, strCodeName)

				If not gDoErrorRtn ("GetMEDCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",Row, vntData(4,1)
						
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MEDNAME"), Row
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_LOWNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",.sprSht_DTL.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",.sprSht_DTL.ActiveRow, trim(vntData(1,1))
						
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_LOWNAME"), Row
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus 
					End If
   				End If
   			End If
		END IF	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_CODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_CODE",.sprSht_DTL.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_NAME",.sprSht_DTL.ActiveRow, trim(vntData(1,1))
						
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_NAME"), Row
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus 
					End If
   				End If
   			End If
		END IF	
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTCODE",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row)
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",Row, vntData(2,1)			
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME"), Row
						.txtCAMPAIGN_NAME1.focus
						.sprSht_DTL.focus 
					End If
   				End If
   			End If
		End If
		
		
		'금액 컬럼 수수료/수수료율계산   		
   		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
   			strAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",Row)
   			strCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMI_RATE",Row)
   			strCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMISSION",Row)
   			If strAMT <> ""  And strCOMMI_RATE <> "" Then
   				lngVALUE = strAMT * strCOMMI_RATE /100
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMISSION",Row, lngVALUE
   				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, lngVALUE - lngMCSUSU
   			ElseIf strAMT <> "" And strCOMMISSION <> "" Then
   				lngVALUE1 = gRound((strCOMMISSION /  strAMT * 100),2)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMI_RATE",Row, lngVALUE1
   			End IF
   			AMT_SUM
   		'수수료율컬럼 수수료/금액계산
   		elseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMI_RATE") Then
   			strAMT		= mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow)
			strCOMMI_RATE   = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMI_RATE",.sprSht_DTL.ActiveRow)
			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow, strAMT	
			If strAMT <> "" And strCOMMI_RATE <> "" Then
   				lngVALUE = strAMT * strCOMMI_RATE /100
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMISSION",Row, lngVALUE
   				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, lngVALUE - lngMCSUSU
   			ElseIf strCOMMISSION <> "" Then
   				if strCOMMI_RATE = 0 then 
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",Row, 0
   					lngMCSUSU = strCOMMISSION * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, strCOMMISSION - lngMCSUSU
   				else
   					lngVALUE1 = gRound((strCOMMISSION / strCOMMI_RATE * 100),0)
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",Row, lngVALUE1
   				end if
   			End IF
			AMT_SUM
		'수수료 컬럼 수수료율 계산
   		elseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") Then
   			strAMT		= mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow)
			strCOMMI_RATE   = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMI_RATE",.sprSht_DTL.ActiveRow)
			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
   			If strAMT <> "" and strAMT <> 0 Then
   				lngVALUE1 = gRound((strCOMMISSION /  strAMT * 100),2)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMI_RATE",Row, lngVALUE1
   				
   				lngMCSUSU = strCOMMISSION * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, strCOMMISSION - lngMCSUSU
   			ELSEIF strAMT = 0 and strCOMMI_RATE = 0 and strCOMMISSION <> "" and strCOMMISSION <> 0 Then
   				lngVALUE1 = strCOMMISSION
  				
   				lngMCSUSU = lngVALUE1 * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, lngVALUE1 - lngMCSUSU
   			End IF
   			AMT_SUM		
   		elseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MCSUSURATE") Then
   			if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MCSUSURATE",.sprSht_DTL.ActiveRow) > 100 THEN
   				gErrorMsgbox "수수료율은 100%를 넘을 수 없습니다.","수정오류"
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSURATE",Row, .txtMCCOMMI_RATE.value
   			END IF
   			strAMT		= mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow)
			strCOMMI_RATE   = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMI_RATE",.sprSht_DTL.ActiveRow)
			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow, strAMT	
			
			If strAMT <> "" and strAMT <> 0 And strCOMMI_RATE <> "" Then
   				lngVALUE = strAMT * strCOMMI_RATE /100
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMISSION",Row, lngVALUE
   				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   			ELSEIF strAMT = 0 and strCOMMI_RATE = 0 and strCOMMISSION <> "" and strCOMMISSION <> 0 Then
   				lngVALUE = strCOMMISSION
  				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"MCSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSU",Row, lngMCSUSU
   			ElseIf strCOMMISSION <> "" AND strCOMMISSION <> 0 AND strCOMMI_RATE <> "" AND strCOMMI_RATE <> 0 Then
   				lngVALUE1 = gRound((strCOMMISSION / strCOMMI_RATE * 100),0)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",Row, lngVALUE1
   			End IF
			AMT_SUM	
		elseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MCSUSU") Then
 			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
 			lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MCSUSU",.sprSht_DTL.ActiveRow)
   			If strCOMMISSION <> "" and strCOMMISSION <> 0 and lngMCSUSU <> "" and lngMCSUSU <> 0 Then
   				lngVALUE1 = gRound((lngMCSUSU /  strCOMMISSION * 100),2)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSURATE",Row, lngVALUE1
   			elseif lngMCSUSU = 0 then
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSURATE",Row, 0
   			End IF
   			AMT_SUM
		elseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXSUSURATE") Then
			if mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXSUSURATE",.sprSht_DTL.ActiveRow) > 100 THEN
   				gErrorMsgbox "수수료율은 100%를 넘을 수 없습니다.","수정오류"
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSURATE",Row, 100 - .txtMCCOMMI_RATE.value
   			END IF
   			strAMT		= mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow)
			strCOMMI_RATE   = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMI_RATE",.sprSht_DTL.ActiveRow)
			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",.sprSht_DTL.ActiveRow, strAMT	
			
			If strAMT <> "" and strAMT <> 0 And strCOMMI_RATE <> "" Then
   				lngVALUE = strAMT * strCOMMI_RATE /100
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMISSION",Row, lngVALUE
   				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EXSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, lngMCSUSU
   			ELSEIF strAMT = 0 and strCOMMI_RATE = 0 and strCOMMISSION <> "" and strCOMMISSION <> 0 Then
   				lngVALUE = strCOMMISSION
  				
   				lngMCSUSU = lngVALUE * mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"EXSUSURATE",Row)/100
   				
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSU",Row, lngMCSUSU
   			ElseIf strCOMMISSION <> "" AND strCOMMISSION <> 0 AND strCOMMI_RATE <> "" AND strCOMMI_RATE <> 0 Then
   				lngVALUE1 = gRound((strCOMMISSION / strCOMMI_RATE * 100),0)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"AMT",Row, lngVALUE1
   			End IF
			AMT_SUM	
		elseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXSUSU") Then
 			strCOMMISSION = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMMISSION",.sprSht_DTL.ActiveRow)
 			lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXSUSU",.sprSht_DTL.ActiveRow)
   			If strCOMMISSION <> "" and strCOMMISSION <> 0 and lngMCSUSU <> "" and lngMCSUSU <> 0 Then
   				lngVALUE1 = gRound((lngMCSUSU /  strCOMMISSION * 100),2)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSURATE",Row, lngVALUE1
   			elseif lngMCSUSU = 0 then
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSURATE",Row, 0
   			End IF
   			AMT_SUM
		end if
		mobjSCGLSpr.CellChanged .sprSht_DTL, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MEDNAME") Then		
			vntInParams = array("","" , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MEDNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_LOWNAME") Then		
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",Row, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"REAL_MED_NAME") Then		
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_NAME",Row, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"EXCLIENTNAME") Then			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				.txtCAMPAIGN_NAME1.focus
				.sprSht_DTL.focus 
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		
		'sprShtToFieldBinding Col, Row
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.sprSht_DTL.Focus
	End With
End Sub


Sub sprSht_DTL_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNMED") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MEDCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MEDNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
			.txtCAMPAIGN_CODE1.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_DTL.Focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2, Row
		elseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNLOW") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_LOWNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWCODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_LOWNAME",Row, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
			.txtCAMPAIGN_CODE1.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_DTL.Focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2, Row
		elseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNREAL") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"REAL_MED_NAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"REAL_MED_NAME",Row, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
			.txtCAMPAIGN_CODE1.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_DTL.Focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2, Row
		ElseIF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNEX") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
			.txtCAMPAIGN_NAME1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_DTL.Focus
			mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2, Row
		End If
	End with
End Sub

Sub DateClean_SHEET (strYEARMON, Row)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		mobjSCGLSpr.SetTextBinding .sprSht_DTL,"DEMANDDAY",Row, date2
	End With
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
	set mobjMDITINTERNETREG	= gCreateRemoteObject("cMDIT.ccMDITINTERNETREG")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
		'캠페인 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 16, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CAMPAIGN_CODE | CAMPAIGN_NAME | CLIENTNAME | TIMNAME | SUBSEQNAME | TBRDSTDATE | TBRDEDDATE | DEPT_NAME | EXCLIENTNAME | MCCOMMI_RATE | MEMO | CLIENTCODE | TIMCODE | SUBSEQ | DEPT_CD | EXCLIENTCODE"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "캠페인코드|캠페인|광고주|팀명|브랜드|시작일|종료일|부서명|대대행사|M&C수수료율|비고|광고주코드|팀코드|브랜드코드|부서코드|대행사코드"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "		9|    20|    14|  13|    13|     9|     9|    10|      12|          9|  20          0|     0|         0|       0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "MCCOMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CAMPAIGN_CODE | CAMPAIGN_NAME | CLIENTNAME | TIMNAME | SUBSEQNAME | DEPT_NAME | EXCLIENTNAME | MEMO ", -1, -1, 200
		mobjSCGLSpr.ColHidden .sprSht_HDR, "CLIENTCODE | TIMCODE | SUBSEQ | DEPT_CD | EXCLIENTCODE", True
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CAMPAIGN_CODE | CAMPAIGN_NAME | CLIENTNAME | TIMNAME | SUBSEQNAME | TBRDSTDATE | TBRDEDDATE | DEPT_NAME | EXCLIENTNAME | MCCOMMI_RATE | MEMO "
		
		'개별청약 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 39, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 10, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 23, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "VOCH_TYPE | CHK | YEARMON | SEQ | DEMANDDAY | MEDNAME | BTNMED | MEDCODE | REAL_MED_LOWNAME | BTNLOW | REAL_MED_LOWCODE | REAL_MED_NAME | BTNREAL | REAL_MED_CODE |  AMT | COMMI_RATE | COMMISSION | MCSUSU | MCSUSURATE | EXSUSU | EXSUSURATE | EXCLIENTNAME | BTNEX | EXCLIENTCODE | MATTERNAME | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | TRU_TAX_FLAG | CAMPAIGN_CODE | CLIENTCODE | TIMCODE | SUBSEQ | TBRDSTDATE | TBRDEDDATE | DEPT_CD | GFLAG | OLDYEARMON | OLDSEQ"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "구분|선택|년월|순번|청구일자|매체명|매체코드|매체사|매체사코드|랩사|랩사코드|광고비|수수료율|수수료|MC수수료|M&C수수료율|대대행수수료|대대행수수료율|대대행사명|대대행사코드|소재명|비고|위수탁거래번호|수수료거래번호|VAT|캠페인코드|광고주코드|팀코드|브랜드코드|광고시작일|광고종료일|부서코드|발행구분|기존년월|기존번호"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 8|   4|   8|   4|       9|    12|     2|5|    12|       2|7|  13|     2|5|    10|       4|    10|      10|          6|          10|             6|        10|         2|8|    15|  15|            12|            12|  4|         0|         0|     0|         0|         0|         0|       0|       0|       0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK | TRU_TAX_FLAG "
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_DTL,"..", "BTNMED | BTNLOW | BTNREAL | BTNEX"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "YEARMON | MEDNAME | MEDCODE | REAL_MED_LOWNAME | REAL_MED_LOWCODE | REAL_MED_NAME | REAL_MED_CODE | EXCLIENTNAME | EXCLIENTCODE | MATTERNAME | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | CAMPAIGN_CODE | CLIENTCODE | TIMCODE | SUBSEQ | DEPT_CD", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE | MCSUSURATE | EXSUSURATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "SEQ | AMT | COMMISSION | MCSUSU | EXSUSU", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, True, "SEQ | TRU_TRANS_NO | COMMI_TRANS_NO"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "CAMPAIGN_CODE | CLIENTCODE | TIMCODE | SUBSEQ | TBRDSTDATE | TBRDEDDATE | DEPT_CD | GFLAG | OLDYEARMON | OLDSEQ", True
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CHK | YEARMON | TRU_TRANS_NO | COMMI_TRANS_NO",-1,-1,2,2,False
		
		.sprSht_HDR.style.visibility = "visible"
		.sprSht_DTL.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDITINTERNETREG = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 그리드 콤보박스 설정
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData_VOCH
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_VOCH = mobjMDITINTERNETREG.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBOVOCH_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_DTL, "VOCH_TYPE",,,vntData_VOCH,,60 
			mobjSCGLSpr.TypeComboBox = True 
   		End If
   	End With
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

		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		
		CALL Get_COMBO_VALUE ()
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON
	Dim strCAMPAIGN_CODE
   	Dim i, strCols
    
	'On error resume next
	with frmThis
	
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "조회시 년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value
		strCAMPAIGN_CODE= .txtCAMPAIGN_CODE1.value
		
		vntData = mobjMDITINTERNETREG.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strCAMPAIGN_CODE)
													
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				SelectRtn_DTL 1,1
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strCAMPAIGN_CODE
   	Dim i, strCols
    Dim intCnt, intCnt2, strRows
    
	with frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0
		intCnt2 = 1

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCAMPAIGN_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CAMPAIGN_CODE",Row)
				
		vntData = mobjMDITINTERNETREG.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strCAMPAIGN_CODE)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
				For intCnt = 1 To .sprSht_DTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRU_TRANS_NO",intCnt) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMI_TRANS_NO",intCnt) <> ""  Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,True,strRows,1,39,True
				
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			else
   				mobjSCGLSpr.SetMaxRows .sprSht_DTL, 10 
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"YEARMON",-1, Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
   				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"VOCH_TYPE",-1, "0"
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTCODE",-1, mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EXCLIENTCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXCLIENTNAME",-1, mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EXCLIENTNAME",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"EXSUSURATE",-1, 100-mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"MCCOMMI_RATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MCSUSURATE",-1, mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"MCCOMMI_RATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMMI_RATE",-1, 20
				
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CAMPAIGN_CODE",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CAMPAIGN_CODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTCODE",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TIMCODE",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TIMCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"SUBSEQ",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SUBSEQ",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TBRDSTDATE",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDSTDATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TBRDEDDATE",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDEDDATE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"DEPT_CD",-1, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_CD",.sprSht_HDR.ActiveRow)
			
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"TRU_TAX_FLAG",-1, "1"
				DateClean_SHEET Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2), -1
   				gWriteText lblStatusDTR, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				
   			End If
   			AMT_SUM
   		End If
   	end with
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
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' 저장로직
'****************************************************************************************
Sub ProcessRtn ()
	Dim intRtn
   	Dim vntData, i
   	Dim intCnt
	Dim strYEARMON
	Dim strSEQ
	Dim strDataCHK
	Dim lngCol, lngRow
	Dim strCAMPAIGN_CODE, strCLIENTCODE, strTIMCODE, strSUBSEQ
	Dim strTBRDSTDATE, strTBRDEDDATE, strDEPT_CD
	With frmThis
	
		if .sprSht_DTL.MaxRows = 0 Then
			gErrorMsgBox "청약 데이터를 추가하십시오.","저장안내!"
			Exit Sub
		end if
		
		'데이터 Validation
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_DTL, "YEARMON | DEMANDDAY | REAL_MED_CODE",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 년월/청구일/랩사코드는 필수 입력사항입니다.","저장안내"
			Exit Sub
		End If
		
		
		For intCnt = 1 to .sprSht_DTL.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"YEARMON",intCnt) = "" OR mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"REAL_MED_CODE",intCnt) = ""  Then
				mobjSCGLSpr.DeleteRow .sprSht_DTL,intCnt
			ELSE
				mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, 1, intCnt
			End If
		Next
		
		strCAMPAIGN_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CAMPAIGN_CODE",.sprSht_HDR.ActiveRow)
		strCLIENTCODE	 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTCODE",.sprSht_HDR.ActiveRow)
		strTIMCODE		 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TIMCODE",.sprSht_HDR.ActiveRow)
		strSUBSEQ		 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SUBSEQ",.sprSht_HDR.ActiveRow)
		strTBRDSTDATE	 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDSTDATE",.sprSht_HDR.ActiveRow)
		strTBRDEDDATE	 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TBRDEDDATE",.sprSht_HDR.ActiveRow)
		strDEPT_CD		 = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_CD",.sprSht_HDR.ActiveRow)

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"VOCH_TYPE | CHK | YEARMON | SEQ | DEMANDDAY | MEDNAME | BTNMED | MEDCODE | REAL_MED_LOWNAME | BTNLOW | REAL_MED_LOWCODE | REAL_MED_NAME | BTNREAL | REAL_MED_CODE |  AMT | COMMI_RATE | COMMISSION | MCSUSU | MCSUSURATE | EXSUSU | EXSUSURATE | EXCLIENTNAME | BTNEX | EXCLIENTCODE | MATTERNAME | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | TRU_TAX_FLAG | CAMPAIGN_CODE | CLIENTCODE | TIMCODE | SUBSEQ | TBRDSTDATE | TBRDEDDATE | DEPT_CD | GFLAG | OLDYEARMON | OLDSEQ")
				
		'처리 업무객체 호출
		intRtn = mobjMDITINTERNETREG.ProcessRtn(gstrConfigXml, vntData, strCAMPAIGN_CODE, strCLIENTCODE, strTIMCODE, strSUBSEQ, _
												strTBRDSTDATE, strTBRDEDDATE, strDEPT_CD)
	
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gErrorMsgBox "자료가 저장" & mePROC_DONE,"저장안내" 
			
			SelectRtn
   		end if
   	end with
End Sub

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '실제데이터여부 플레
	Dim lngchkCnt
	Dim strSEQ	
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		for i = 1 to .sprSht_DTL.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRU_TRANS_NO",i) <> "" Then
					gErrorMsgBox "선택하신 " & i & "행의 자료는 거래명세표가 존재 합니다." & vbcrlf & "먼저 거래명세표를 삭제 하십시오!","삭제안내!"
					exit Sub
				else 
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"GFLAG",i) = "1" Then
						gErrorMsgBox "선택하신 " & i & "행의 자료는 승인된 자료입니다." & vbcrlf & "먼저 승인취소처리 하십시오!","삭제안내!"
						exit Sub
					End If
					lngchkCnt = lngchkCnt +1
				End If
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_DTL.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"YEARMON",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_DTL,i
				else
					intRtn = mobjMDITINTERNETREG.DeleteRtn(gstrConfigXml, strYEARMON, strSEQ)
					
					IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht_DTL,i
   					End IF
   					
   					strSEQFLAG = TRUE
				end if				
   				intCnt = intCnt + 1
   			END IF
		next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_DTL
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
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
													<TABLE cellSpacing="0" cellPadding="0" width="118" background="../../../images/back_p.gIF"
														border="0">
														<TR>
															<TD align="left" width="100%" height="2"></TD>
														</TR>
													</TABLE>
												</td>
											</tr>
											<tr>
												<td class="TITLE">청약관리 - 개별청약</td>
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
												<TD class="SEARCHDATA" width="130"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
													width="60">캠페인</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCAMPAIGN_NAME1" title="코드명" style="WIDTH: 217px; HEIGHT: 22px"
														type="text" maxLength="100" size="30" name="txtCAMPAIGN_NAME1"> <IMG id="ImgCAMPAIGN_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCAMPAIGN_CODE1"> <INPUT class="INPUT" id="txtCAMPAIGN_CODE1" title="코드입력" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtCAMPAIGN_CODE1"></TD>
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
															<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="화면을 초기화 합니다."
																	src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
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
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 30.18%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="4339">
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
												<TD class="TITLE" align="left" width="400" height="22" vAlign="absmiddle">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
																<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="22">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
															<TD><IMG id="ImgSave" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/ImgSave.gIF'" height="20" alt="매체사별로 거래명세서를 생성합니다.."
																	src="../../../images/ImgSave.gIF" border="0" name="ImgSave"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="개별 거래명세서를 출력합니다.." src="../../../images/imgDelete.gIF" border="0"
																	name="imgDelete"></TD>
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
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="8916">
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
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
