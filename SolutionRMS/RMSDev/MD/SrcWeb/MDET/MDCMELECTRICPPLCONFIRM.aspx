<%@ Page CodeBehind="MDCMELECTRICPPLCONFIRM.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRICPPLCONFIRM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 가상/간접광고 승인화면</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/표준샘플/스프레드쉬트
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : SpreadSheet를 이용한 조회/입력/수정/삭제/인쇄 처리 표준 샘플
'파라  메터 : 
'특이  사항 : 표준샘플을 위해 만든 것임
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
'****************************************************************************************
-->
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<!-- StyleSheet 정보 --><LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script id="clientEventHandlersVBS" language="vbscript">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET
Dim mobjMDETELECTRICPPLLIST 
Dim mstrCheck
Dim mstrConfirmGBN

mstrConfirmGBN = "Y"
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
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오",""
		exit Sub
	end if
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'승인버튼 클릭
Sub imgAgree_onclick
	gFlowWait meWAIT_ON
		mstrConfirmGBN = "Y"
		ProcessRtn_Confirm(mstrConfirmGBN)
	gFlowWait meWAIT_OFF
End Sub

'승인취소 버튼 클릭
Sub imgAgreeCancel_onclick
	gFlowWait meWAIT_ON
		mstrConfirmGBN = "N"
		ProcessRtn_Confirm(mstrConfirmGBN)
	gFlowWait meWAIT_OFF
End Sub

'엑셀버튼 클릭
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'닫기버튼 클릭
Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					SELECTRTN
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'제작 대행사 파트너 팝업 
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code값 저장
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'코드명 표시
			gSetChangeFlag .txtEXCLIENTCODE
		End If
	end With
End Sub


Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)

			vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code값 저장
					.txtEXCLIENTNAME.value = trim(vntData(2,1))	'코드명 표시
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub



Sub rdT_onclick
	rdChecked
	SelectRtn
End Sub

Sub rdF_onclick
	rdChecked
	SelectRtn
End Sub

Sub rdChecked
	with frmThis
		If .rdT.checked = True Then
			.imgAgreeCanCel.style.display = "none"
			.imgAgree.style.display = "inline"
		Else
			.imgAgree.style.display = "none"
			.imgAgreeCanCel.style.display = "inline"
		End If
	End with
End sub
'****************************************************************************************
' 쉬트 더불클릭 이벤트
'****************************************************************************************

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
	
		If Row = 0 and Col = 1  then 'AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "N"
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) <> "" Then
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
				
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
				
				for intcnt = 1 to .sprSht.MaxRows
					sprSht_Change 1, intcnt
					
				next
				For intCnt = 1 To .sprSht.MaxRows
					If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
						'스태틱
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					'Else
						'체크
					'	mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
					End If			
				Next
			End IF
		end if
	end with
End Sub  	


sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	
	'서버업무객체 생성	
	set mobjMDETELECTRICPPLLIST = gCreateRemoteObject("cMDET.ccMDETELECTRICPPLLIST")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
       gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 26, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO "
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|년월|순번|광고주코드|광고주명|담당부서코드|담당부서명|채널|프로그램|요일|청약방송시작일|청약방송종료일|총횟수|당월횟수|잔여횟수|매체비단가|월총매체비|월총수수료|파트너코드|파트너명|파트너회당매체수익|월총매체청구비|전표번호|승인자|승일일|비고"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   4|         8|      18|           0|        10|	8|      15|   5|            12|            12|     7|       7|       7|        10|        10|        10|         8|      15|                15|            13|       0|     0|     0|  20"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "TBRDDAY", -1, -1, "월" & vbTab & "화" & vbTab & "수" & vbTab & "목" & vbTab & "금" & vbTab & "토" & vbTab & "일"  , 10, 40, False, False
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | CNT_AMT | EXSUSU ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDFDATE | TBRDTDATE | CONFIRM_DATE ", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, " CLIENTCODE | CLIENTNAME | DEPT_CD | MEDNAME | PROGRAM | EXCLIENTCODE | EXCLIENTNAME | VOCHNO | CONFIRM_USER ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE |  EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO "
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | DEPT_CD | VOCHNO | EXCLIENTCODE |  CONFIRM_USER | CONFIRM_DATE", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | YEARMON | SEQ | PROGRAM ",-1,-1,2,2,False  '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEMO",-1,-1,0,2,false
		.sprSht.style.visibility = "visible"
    End With

	InitPageData
	'SelectRtn	
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	set mobjMDETELECTRICPPLLIST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	with frmThis
		.txtYEARMON.value =  Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
		
		.imgAgreeCanCel.style.display = "none"
		.imgAgree.style.display = "inline"
			
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	dim vntData
	Dim strYEARMON, strCLIENTCODE, strEXCLIENTCODE
   	Dim intCnt, strGBN, strEMPNO
	on error resume next
	with frmThis
		
		.sprSht.MaxRows = 0
		
		strYEARMON			= .txtYEARMON.value
		strCLIENTCODE		= .txtCLIENTCODE.value
		strEXCLIENTCODE		= .txtEXCLIENTCODE.value
		strEMPNO			= gstrEmpNo
		
		if .rdF.checked = TRUE then
			strGBN = "Y"
		ELSE
			strGBN = "N"
		end if
		

		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		vntData = mobjMDETELECTRICPPLLIST.SelectRtn_confirm(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE , strEXCLIENTCODE,strGBN,strEMPNO)
		
		IF not gDoErrorRtn ("SelectRtn_confirm") then
			'조회한 데이터를 바인딩
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				
			Next
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If 	mlngRowCnt < 1 Then
			.sprSht.MaxRows= 0
			End If
		End IF
		gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE	
		
	end with
End Sub

'------------------------------------------
' 승인/취소 저장로직
'------------------------------------------
Sub ProcessRtn_Confirm(strCONFIRMFLAG)
	Dim intRtn, intRtnChk
   	Dim vntData
   	Dim lngCHK , intCnt
	
	with frmThis
		'On error resume next
   		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = lngCHK + 1
			END IF
		Next
		
		If lngCHK = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		IF strCONFIRMFLAG = "Y" THEN
			intRtnChk = gYesNoMsgbox("선택하신 자료를 승인 하시겠습니까?","승인안내")
			If intRtnChk <> vbYes then 
				exit sub
			End If
		ELSE 
			intRtnChk = gYesNoMsgbox("선택하신 자료를 승인취소 하시겠습니까?" & vbcrlf & "승인취소하시면 데이터가 반려됩니다.","승인 취소 안내")
			If intRtnChk <> vbYes then 
				exit sub
			End If
		END IF 
	
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO")
		
		intRtn = mobjMDETELECTRICPPLLIST.ProcessRtn_ConfirmOK(gstrConfigXml,vntData,strCONFIRMFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOK") then 
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			If strCONFIRMFLAG = "Y" Then
				msgbox lngCHK & " 건의 자료가 승인" & mePROC_DONE
			else
				msgbox lngCHK & " 건의 자료가 승인 취소" & mePROC_DONE
			end if
			
			SelectRtn
   		end if
   	end with
End Sub


-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE style="WIDTH: 100%; HEIGHT: 98%" id="tblForm" border="0" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<TD height="20" width="400" align="left">
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
											<td class="TITLE">공중파 가상간접광고 승인</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="처리중입니다."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--테이블이 무너지는것을 막아준다-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%">
							<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 75px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')">년 
												월</TD>
											<TD style="WIDTH: 87px" class="SEARCHDATA"><INPUT accessKey="NUM" id="txtYEARMON" class="INPUT" title="년월조회" maxLength="6" size="10"
													name="txtYEARMON"></TD>
											<TD style="WIDTH: 51px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">광고주
											</TD>
											<TD style="WIDTH: 239px" class="SEARCHDATA" width="239"><INPUT style="WIDTH: 150px; HEIGHT: 22px" id="txtCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="광고주명" dataFld="CLIENTNAME" size="32" name="txtCLIENTNAME"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 64px; HEIGHT: 22px" id="txtCLIENTCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="광고주코드" dataFld="CLIENTCODE" size="5" name="txtCLIENTCODE"></TD>
											<TD style="HEIGHT: 22px; CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="70">파트너</TD>
											<TD class="DATA"><INPUT style="WIDTH: 150px; HEIGHT: 22px" id="txtEXCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="제작사명" dataFld="EXCLIENTNAME" maxLength="100" size="30" name="txtEXCLIENTNAME">
												<IMG style="CURSOR: hand" id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgEXCLIENTCODE"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT style="WIDTH: 55px; HEIGHT: 22px" id="txtEXCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="제작사코드" dataFld="EXCLIENTCODE" maxLength="10" size="4" name="txtEXCLIENTCODE"></TD>
											<TD class="SEARCHDATA"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="자료를 검색합니다."
													align="right" src="../../../images/imgQuery.gIF" height="20"></TD>
										</TR>
										<tr>
											<TD style="CURSOR: hand" class="SEARCHLABEL" title="자료를 승인 하거나 취소합니다." width="75">작업선택</TD>
											<TD class="SEARCHDATA" colSpan="6">&nbsp;<INPUT id="rdT" title="요청내역조회" value="rdT" CHECKED type="radio" name="rdGBN">&nbsp;요청내역 
												조회 <INPUT id="rdF" title="승인내역조회" value="rdF" type="radio" name="rdGBN">&nbsp;승인내역조회&nbsp;</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 25px" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%" class="BODYSPLIT">
						<!--테스트 시작-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" vAlign="middle" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton1" border="0" cellSpacing="0" cellPadding="2"
										width="50">
										<TR>
											<TD><IMG style="CURSOR: hand" id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'" border="0" name="imgAgree"
													alt="선택한 행을 승인합니다." align="absMiddle" src="../../../images/imgAgree.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'" border="0"
													name="imgAgreeCanCel" alt="선택한 행을 승인취소 합니다." align="absMiddle" src="../../../images/imgAgreeCanCel.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!-- 추가 디자인끝--></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 1040px; HEIGHT: 3px" class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" class="LISTFRAME" vAlign="top" align="center">
						<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%" id="pnlTab1" ms_positioning="GridLayout">
							<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31802">
								<PARAM NAME="_ExtentY" VALUE="11853">
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
								<PARAM NAME="EditEnterAction" VALUE="5">
								<PARAM NAME="EditModePermanent" VALUE="0">
								<PARAM NAME="EditModeReplace" VALUE="0">
								<PARAM NAME="FormulaSync" VALUE="-1">
								<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
								<PARAM NAME="GridColor" VALUE="12632256">
								<PARAM NAME="GridShowHoriz" VALUE="1">
								<PARAM NAME="GridShowVert" VALUE="1">
								<PARAM NAME="GridSolid" VALUE="1">
								<PARAM NAME="MaxCols" VALUE="19">
								<PARAM NAME="MaxRows" VALUE="0">
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
					<TD id="lblStatus" class="BOTTOMSPLIT"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
