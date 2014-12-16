<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLIENTYEARBRANSLIST.aspx.vb" Inherits="MD.MDCMCLIENTYEARBRANSLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>브랜드별 집행실적</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/그룹광고 분담금 입력/조회 화면(MDCMGROUP)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMGROUP.aspx.aspx
'기      능 : 그룹광고 분담금 을 조회/입력 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/01/09 By Kim Tae Yub
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
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET, mobjMDSRREPORTLIST'공통코드, 클래스

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
	if frmThis.txtYEAR.value = "" or frmThis.txtCLIENTCODE.value = ""  then
		gErrorMsgBox "년도와 광고주를 입력하시오","조회안내"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
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
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
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
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'체크체인지
Sub chkALL_onclick
	CheckCleanField
End Sub

Sub chkMEDFLAG1_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG2_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG3_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG4_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG5_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG6_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG7_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG8_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG9_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chk_chk
	with frmThis
		If .chkMEDFLAG1.checked = false and .chkMEDFLAG2.checked = false and .chkMEDFLAG3.checked = false and  _ 
			.chkMEDFLAG4.checked = false and .chkMEDFLAG5.checked = false and .chkMEDFLAG6.checked = false and  _ 
			.chkMEDFLAG7.checked = false and .chkMEDFLAG8.checked = false and .chkMEDFLAG9.checked = false then
		.chkALL.checked = True
		Else
		.chkALL.checked = false
		end If
	end with
End Sub

Sub CheckCleanField
	with frmThis
		.chkALL.checked = True
		.chkMEDFLAG1.checked = False
		.chkMEDFLAG2.checked = False
		.chkMEDFLAG3.checked = False
		.chkMEDFLAG4.checked = False
		.chkMEDFLAG5.checked = False
		.chkMEDFLAG6.checked = False
		.chkMEDFLAG7.checked = False
		.chkMEDFLAG8.checked = False
		.chkMEDFLAG9.checked = False
	End with
End Sub

'매체구분 화체크하면 변경한다.
sub chkMED_FLAG_onclick
	with frmThis
		if .chkMED_FLAG.checked then
			MED_FLAG.style.visibility = "visible" 
		else
			MED_FLAG.style.visibility = "hidden" 
		end if
	end with
end sub

'광고주 클릭
sub chkCLIENT_onclick
end sub

'브랜드 클릭
sub chkSUBSEQ_onclick
end sub

'소재 클릭
sub chkMATTER_onclick
end sub



'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
    
		gSetSheetDefaultColor
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* 년월을 광고주 조회조건을 입력하시고 조회버튼을 눌러주세요."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "123"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

'조회 조건에 따라서 시트를 새로 그린다.
Sub makePageData(strGUBUN)
     Dim strFIELD
     Dim strFIELDNAME
     Dim i, j
     Dim intCNT,intCNT2
     Dim strGUBUNCODE
     
     With frmThis
        .sprSht.MaxRows = 0
        
        strGUBUNCODE = split(strGUBUN,",")
        
        intCNT = ubound(strGUBUNCODE,1)
        intCNT2 = 0
        
        for i = 0 to intCNT
			if i = 0 then
				if strGUBUNCODE(i) = "1" then
					strFIELD = "CLIENTNAME"
					strFIELDNAME = "광고주명"
					intCNT2 = intCNT2 + 1
				else
					strFIELD = ""
				end if
			end if
			
			if i = 1 then
				if strGUBUNCODE(i) = "1" then
					if strFIELD = "" then
						strFIELD = strFIELD & " SUBSEQNAME"
						strFIELDNAME = strFIELDNAME & "브랜드명"
						intCNT2 = intCNT2 + 1
					else
						strFIELD = strFIELD & " | SUBSEQNAME"
						strFIELDNAME = strFIELDNAME & " | 브랜드명"
						intCNT2 = intCNT2 + 1
					end if
				end if
			end if
			
			if i = 2 then
				if strGUBUNCODE(i) = "1" then
					if strFIELD = "" then
						strFIELD = strFIELD & "MATTERNAME"
						strFIELDNAME = strFIELDNAME & "소재명"
						intCNT2 = intCNT2 + 1
					else
						strFIELD = strFIELD & " | MATTERNAME"
						strFIELDNAME = strFIELDNAME & " | 소재명"
						intCNT2 = intCNT2 + 1
					end if
				end if
			end if
			if i = 3 then
				if strGUBUNCODE(i) = "1" then
					if strFIELD = "" then
						strFIELD = strFIELD & "MED_FLAG"
						strFIELDNAME = strFIELDNAME & "매체구분"
						intCNT2 = intCNT2 + 1
					else
						strFIELD = strFIELD & " | MED_FLAG"
						strFIELDNAME = strFIELDNAME & " | 매체구분"
						intCNT2 = intCNT2 + 1
					end if
				end if
			end if			
        next

		intCNT2 = intCNT2 + 13
        
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, intCNT2, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, strFIELD & "|A1 | A2 |  A3 |  A4 |  A5 |  A6 |  A7 |  A8 |  A9 |  A10 |  A11 |  A12 | SUMAMT"
		mobjSCGLSpr.SetHeader .sprSht,        strFIELDNAME & "|1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월|총합계"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 15| 12| 15|10| 10| 10| 10| 10| 10| 10| 10| 10| 10|  10|  10|  10|    12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "A1 | A2 |  A3 |  A4 |  A5 |  A6 |  A7 |  A8 |  A9 |  A10 |  A11 |  A12 | SUMAMT", -1, -1,0
		
		mobjSCGLSpr.SetCellAlign2 .sprSht, strFIELD ,-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht, strFIELD
		
		
    End With
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	set mobjMDSRREPORTLIST = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtYEAR.value = Mid(gNowDate,1,4)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtYEAR.focus()
	End with
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTCODE
   	Dim strGUBUN

   	'매체구분에 따른 변수 
   	Dim strMEDFLAGALL,strMEDFLAG1,strMEDFLAG2,strMEDFLAG3,strMEDFLAG4,strMEDFLAG5,strMEDFLAG6,strMEDFLAG7,strMEDFLAG8,strMEDFLAG9
   	
	'On error resume next
	with frmThis
		
		if .chkCLIENT.checked = false and .chkSUBSEQ.checked =false and .chkMATTER.checked = false and .chkMED_FLAG.checked = false then
			gErrorMsgBox "조회구분중에 체크를 해야 합니다..",""
			exit sub	
		end if		
		
		'체크를 조립한다.
		if .chkCLIENT.checked then
			strGUBUN = "1"
		else
			strGUBUN = "0"
		end if
		
		if .chkSUBSEQ.checked then
			strGUBUN = strGUBUN & ",1"
		else
			strGUBUN = strGUBUN & ",0"
		end if
		
		if .chkMATTER.checked then
			strGUBUN = strGUBUN & ",1"
		else
			strGUBUN = strGUBUN & ",0"
		end if
		if .chkMED_FLAG.checked then
			strGUBUN = strGUBUN & ",1"
		else
			strGUBUN = strGUBUN & ",0"
		end if
		
		'시트 다시 그리기 
		makePageData strGUBUN
		
		
		'Sheet초기화
		.sprSht.MaxRows = 0
	
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		If .chkALL.checked = True Then strMEDFLAGALL = "1" Else strMEDFLAGALL = "0" 
		If .chkMEDFLAG1.checked = True Then strMEDFLAG1 = "1" Else strMEDFLAG1 = "0"  
		If .chkMEDFLAG2.checked = True Then strMEDFLAG2 = "1" Else strMEDFLAG2 = "0"
		If .chkMEDFLAG3.checked = True Then strMEDFLAG3 = "1" Else strMEDFLAG3 = "0"
		If .chkMEDFLAG4.checked = True Then strMEDFLAG4 = "1" Else strMEDFLAG4 = "0"
		If .chkMEDFLAG5.checked = True Then strMEDFLAG5 = "1" Else strMEDFLAG5 = "0"
		If .chkMEDFLAG6.checked = True Then strMEDFLAG6 = "1" Else strMEDFLAG6 = "0"
		If .chkMEDFLAG7.checked = True Then strMEDFLAG7 = "1" Else strMEDFLAG7 = "0"
		If .chkMEDFLAG8.checked = True Then strMEDFLAG8 = "1" Else strMEDFLAG8 = "0"
		If .chkMEDFLAG9.checked = True Then strMEDFLAG9 = "1" Else strMEDFLAG9 = "0"

		strCLIENTCODE = .txtCLIENTCODE.value

		vntData = mobjMDSRREPORTLIST.SelectRtn_CLIENTYEARBRANDLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEAR.value, strCLIENTCODE, _
																	strMEDFLAGALL, strMEDFLAG1, strMEDFLAG2, strMEDFLAG3, strMEDFLAG4, _
																	strMEDFLAG5, strMEDFLAG6, strMEDFLAG7, strMEDFLAG8, strMEDFLAG9, strGUBUN)

		if not gDoErrorRtn ("SelectRtn_CLIENTYEARBRANDLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		
   		'총계쪽 시트 색변환
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
		
		mobjSCGLSpr.SetTextBinding .sprSht,1,.sprSht.maxRows, "총계"
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, .sprSht.maxRows, .sprSht.maxRows,&HCCFFFF, &H000000,False
	
	End With
End Sub

		</script>
	</HEAD>
	<body class="base">
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
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
												<TABLE cellSpacing="0" cellPadding="0" width="146" background="../../../images/back_p.gIF"
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
											<td class="TITLE">브랜드별 집행 실적</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
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
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="년도을삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEAR,'')">년&nbsp; 
												도
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 130px"><INPUT class="INPUT" id="txtYEAR" title="년도을입력하세요" style="WIDTH: 100px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="4" size="14" name="txtYEAR">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="80">광고주
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 192px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="26" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
											</TD>
											<td style="WIDTH: 400px"></td>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" height="25">조회 구분
											</TD>
											<TD class="SEARCHDATA" colSpan="4">
												<INPUT id="chkCLIENT" type="checkbox" value="CLIENT" checked name="chkCLIENT">&nbsp;1.광고주
												<INPUT id="chkSUBSEQ" type="checkbox" value="SUBSEQ" checked name="chkSUBSEQ">&nbsp;2.브랜드
												<INPUT id="chkMATTER" type="checkbox" value="MATTER" checked name="chkMATTER">&nbsp;3.소재
												<INPUT id="chkMED_FLAG" type="checkbox" value="MED_FLAG" checked name="chkMED_FLAG">&nbsp;4.매체별
											</TD>
										</tr>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" height="25">매체구분
											</TD>
											<TD class="SEARCHDATA" colSpan="4">
												<div id="MED_FLAG" style="VISIBILITY: visible"><INPUT id="chkALL" type="checkbox" CHECKED name="chkALL">&nbsp;전체&nbsp;
													<INPUT id="chkMEDFLAG3" type="checkbox" name="chkMEDFLAG3">&nbsp;TV <INPUT id="chkMEDFLAG4" type="checkbox" name="chkMEDFLAG4">&nbsp;Radio
													<INPUT id="chkMEDFLAG5" type="checkbox" name="chkMEDFLAG5">&nbsp;지상파DMB <INPUT id="chkMEDFLAG6" type="checkbox" name="chkMEDFLAG6">&nbsp;CATV
													<INPUT id="chkMEDFLAG9" type="checkbox" name="chkMEDFLAG9">&nbsp;종합편성방송 <INPUT id="chkMEDFLAG1" type="checkbox" name="chkMEDFLAG1">&nbsp;신문
													<INPUT id="chkMEDFLAG2" type="checkbox" name="chkMEDFLAG2">&nbsp;잡지 <INPUT id="chkMEDFLAG7" type="checkbox" name="chkMEDFLAG7">&nbsp;인터넷
													<INPUT id="chkMEDFLAG8" type="checkbox" name="chkMEDFLAG8">&nbsp;옥외
												</div>
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="굴림"></FONT></TD>
							</TR> <!--BodySplit End--> <!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="16193">
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
							</TR> <!--List End--> <!--Bottom Split Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
								</TD>
							</TR> <!--Bottom Split End--></TABLE> <!--Input Define Table End--></TD>
				</TR> <!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
