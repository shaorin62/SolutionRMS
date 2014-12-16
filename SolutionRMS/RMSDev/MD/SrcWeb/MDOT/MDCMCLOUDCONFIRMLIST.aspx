<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUDCONFIRMLIST.aspx.vb" Inherits="MD.MDCMCLOUDCONFIRMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>개별청약 승인/조회</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'시스템구분 : MD/OUTDOORLIST 청약승인화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMOUTDOORLIST.aspx
'기      능 : 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/23 By Hwang Duck su
			:2) 2009/09/28 By Kim Tae Yub
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script id="clientEventHandlersVBS" language="vbscript">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjCLOUDCONFIRM
Dim mobjMDCMGET
Dim mstrCheck

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


Sub imgClose_onclick ()
	'Window_OnUnload
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------

'승인버튼
Sub imgSetting_onclick
	Call ProcessRtn_ConfirmOK()
End Sub

'승인취소
Sub imgSettingCancel_onclick
	Call ProcessRtn_ConfirmCancel()
End Sub

'조회버튼 클릭시
Sub imgQuery_onclick
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'엑셀버튼 클릭시
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK")  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			NEXT
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

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
	
	End If
	
	With frmThis 
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") OR _
		   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
		   
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") Then
				strCOLUMN = "EXSUSU"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
				strCOLUMN = "CGV_AMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU")) OR _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") or  _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
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
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성									
	set mobjCLOUDCONFIRM	= gCreateRemoteObject("cMDOT.ccMDOTCLOUDCONFIRM")
	set mobjMDCMGET		    = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		'상단 청약 내용
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 17, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|승인유무|구분|년월|번호|계약코드|계약명|광고주|담당부서|월청구액|대행사코드|대행사명|대행수수료|지점별지급액|지점명|청구지명|비고"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|       5|   4|   5|   4|       8|    15|    15|      15|       8|         8|      12|         8|           8|    10|      12|  10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT | EXSUSU | CGV_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONFIRMGBN | GBN | YEARMON | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | CGV_NAME | CUSTNAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | CGV_NAME | CUSTNAME | MEMO",-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"CONT_CODE"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjCLOUDCONFIRM = Nothing
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
		.txtYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		.txtCLIENTNAME.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON
	Dim strCLIENTNAME
	Dim strCLIENTCODE
	Dim strCONT_CODE
	Dim strCONT_NAME

	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCONT_CODE	= .txtCONT_CODE.value
		strCONT_NAME	= .txtCONT_NAME.value
		

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjCLOUDCONFIRM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											strYEARMON, strCLIENTCODE, strCLIENTNAME , _
											strCONT_CODE, strCONT_NAME)

		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE	
   		end if
   		AMT_SUM
   	end with
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
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
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

'------------------------------------------
'정약 데이터 확정
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strUSER
	
	with frmThis
			strUSER = gstrEmpNo

   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 확정이 불가능 합니다.","확정안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) = "확정" THEN
   					gErrorMsgBox "미확정상태인 데이터만 확정이 가능합니다.","저장안내!"
					Exit Sub
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "확정할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If


		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO")
		
		intRtn = mobjCLOUDCONFIRM.ProcessRtn(gstrConfigXml,vntData,strUSER,"CONFIRM")
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOK") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 확정" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 정약 데이터 확정 취소
'------------------------------------------
Sub ProcessRtn_ConfirmCancel
    Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strUSER

	
	with frmThis
		strUSER = gstrEmpNo
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 승인취소이 불가능 합니다.","승인취소안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) = "미확정" THEN
   					gErrorMsgBox "확정상태인 데이터만 확정취소가 가능합니다.","확정 취소 안내!"
					Exit Sub					
   				END IF
   				
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "승인취소할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		'On error resume next
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO")
		
		intRtn = mobjCLOUDCONFIRM.ProcessRtn(gstrConfigXml,vntData,strUSER,"CANCEL")
		
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 확정취소" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<td style="WIDTH: 400px" height="28" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="160" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td id="tblTitleName" class="TITLE">CGV 클라우드 청약확정</td>
										</tr>
									</table>
								</td>
								<TD height="28" vAlign="middle" width="640" align="right">
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 350px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="처리중입니다."
													src="../../../images/Waiting.GIF" height="23"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1024" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
							<TR>
								<TD style="HEIGHT: 10px" class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="left">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="45">년월</TD>
											<TD style="WIDTH: 100px" class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtYEARMON" class="INPUT"
													title="년월조회" maxLength="6" size="10" name="txtYEARMON"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="50">광고주</TD>
											<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 170px; HEIGHT: 22px" id="txtCLIENTNAME" class="INPUT_L" title="코드명"
													maxLength="100" align="left" size="22" name="txtCLIENTNAME"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE" class="INPUT_L" title="코드조회"
													maxLength="6" align="left" name="txtCLIENTCODE"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONT_NAME,'')"
												width="50">계약명</TD>
											<TD style="WIDTH: 130px" class="SEARCHDATA"><INPUT style="WIDTH: 160px; HEIGHT: 22px" id="txtCONT_NAME" class="INPUT_L" title="계약명"
													maxLength="100" size="22" name="txtCONT_NAME">
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONT_CODE,'')"
												width="60">계약 코드</TD>
											<TD style="WIDTH: 86px" class="SEARCHDATA"><INPUT style="WIDTH: 80px; HEIGHT: 22px" id="txtCONT_CODE" class="INPUT_L" title="계약코드"
													maxLength="6" size="6" name="txtCONT_CODE"></TD>
											<TD align="right"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="자료를 조회합니다."
													src="../../../images/imgQuery.gIF" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<tr>
								<td>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="13">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 25px" class="TOPSPLIT"></TD>
										</TR>
									</TABLE>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28">
										<TR>
											<td style="WIDTH: 100%" class="DATA">합계 : <INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 20px" id="txtSUMAMT" class="NOINPUTB_R"
													title="합계금액" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT style="WIDTH: 120px; HEIGHT: 20px" id="txtSELECTAMT" class="NOINPUTB_R" title="선택금액"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
											<TD style="WIDTH: 100%" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'" border="0" name="imgSetting"
																alt="자료를확정처리합니다." src="../../../images/imgSetting.gIF" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgSettingCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gIF'" border="0"
																name="ImgConfirmCancel" alt="확정을 취소합니다." src="../../../images/ImgConfirmCancel.gif"
																height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
																alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></td>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							<!--BodySplit Start-->
							<TR>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="17012">
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
							<!--List End-->
							<!--Bottom Split Start-->
							<TR>
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
