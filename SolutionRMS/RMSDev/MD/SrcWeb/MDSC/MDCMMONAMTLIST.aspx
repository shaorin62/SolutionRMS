<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMONAMTLIST.aspx.vb" Inherits="MD.MDCMMONAMTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>월별 매체별 광고비</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987"> <!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET, mobjEXECUTE, mobjMDSRREPORTLIST'공통코드, 클래스
Dim mClientsubcode

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
	
	if frmThis.txtFYEARMON.value = "" or frmThis.txtTYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	if frmThis.txtCLIENTCODE.value = ""  then
		ImgCLIENTCODE_onclick
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
	Dim strYEARMON
	Dim strCLIENTNAME
	Dim strSUBLIST
	Dim strCLIENTSUBLIST
	Dim intSUBRow
	Dim chkflag
	Dim strCLIENTCODE
	Dim strFYEARMON
	Dim strTYEARMON
	Dim Con1 
	Dim Con2
	Dim Con3
	Dim Con4
	
	with frmThis
		Con1 = ""
		Con2 = ""
		Con3 = ""
		Con4 = ""
		
		if frmThis.sprSht.MaxRows = 0 then
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
			Exit Sub
		end if
		
		
		ModuleDir = "MD"
		ReportName = "MDCMMONAMTLIST.rpt"
		
		strFYEARMON		= .txtFYEARMON.value
		strTYEARMON		= .txtTYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		
		If strFYEARMON <> "" Then Con1 = " AND (YEARMON >= '" & strFYEARMON & "')"
		If strTYEARMON <> "" Then Con2 = " AND (YEARMON <= '" & strTYEARMON & "')"
		If strCLIENTCODE <> "" Then Con3 = " AND (CLIENTCODE = '" & strCLIENTCODE & "')"
		
		strCLIENTSUBLIST=""
		strSUBLIST = ""
		chkflag = 1
		strCLIENTSUBLIST = 	split(mClientsubcode,"♥")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
	
		if strSUBLIST <> "" then Con4 = " AND (CLIENTSUBCODE IN(" & strSUBLIST & "))"
		strCLIENTNAME = .txtCLIENTNAME.value
        
		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & strCLIENTNAME & ":" & Con4
		
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

    
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
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
			Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
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
					Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub GetCLIENTSUBLIST (strCLIENTCODE)
	Dim vntData
   	Dim i, strCols
   	Dim strHTML
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strHTML = "" 
		mClientsubcode = ""
		vntData = mobjMDSRREPORTLIST.GetCLIENTSUBLIST2(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE, .txtFYEARMON.value, .txtTYEARMON.value)
		if not gDoErrorRtn ("GetCLIENTSUBLIST2") then
			If mlngRowCnt > 0 Then
				For i = 0 to mlngRowCnt-1
					strHTML = strHTML & "<INPUT id='"& vntData(0,i) & "' type='checkbox' name='"&  vntData(0,i) & "' checked>" & vntData(1,i) & "&nbsp;&nbsp;"
					IF i = 0 THEN
						mClientsubcode = mClientsubcode & vntData(0,i)
					ELSE
						mClientsubcode = mClientsubcode & "♥" & vntData(0,i)
					END IF
				next
			Else
				strHTML = ""
			End If
			document.getElementById("tdCLIENTSUB").innerHTML = strHTML
   		end if
   	end with
End Sub

Sub txtFYEARMON_onchange
	if frmThis.txtCLIENTCODE.value <> "" then
		Call GetCLIENTSUBLIST (frmThis.txtCLIENTCODE.value)
	else
		document.getElementById("tdCLIENTSUB").innerHTML = ""
	end if
End Sub

Sub txtTYEARMON_onchange
	if frmThis.txtCLIENTCODE.value <> "" then
		Call GetCLIENTSUBLIST (frmThis.txtCLIENTCODE.value)
	else
		document.getElementById("tdCLIENTSUB").innerHTML = ""
	end if
End Sub

Sub txtCLIENTCODE_onchange
	if frmThis.txtCLIENTCODE.value <> "" then
		Call GetCLIENTSUBLIST (frmThis.txtCLIENTCODE.value)
	else
		document.getElementById("tdCLIENTSUB").innerHTML = ""
	end if
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST")
	set mobjEXECUTE	= gCreateRemoteObject("cMDSC.ccMDSCEXECUTE")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
         gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 2, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "YEARMONFLAG | CLIENTNAME | CLIENTSUBNAME | TV | RD | CATV | DMB | TOTAL | MP01 | MP02 | INTERNET | OUTDOOR | SUMAMT"
		mobjSCGLSpr.SetHeader .sprSht,			"년월|광고주|구분|TV|RD|CATV|지상파DMB|종합편성방송|신문|잡지|인터넷|옥외|계"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "    7|    15|  12|11|11|  11|       11|          11|  11|  11|   11|  11| 12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TV | RD | CATV | DMB | TOTAL | MP01 | MP02 | INTERNET | OUTDOOR | SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMONFLAG | CLIENTNAME | CLIENTSUBNAME | TV | RD | CATV | DMB | TOTAL | MP01 | MP02 | INTERNET | OUTDOOR | SUMAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMONFLAG | CLIENTNAME | CLIENTSUBNAME",-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMONFLAG"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSRREPORTLIST = Nothing
	set mobjMDCOGET = Nothing
	set mobjEXECUTE = Nothing
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
		.txtFYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		.txtTYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtFYEARMON.focus()
		
	End with	
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	Dim chkflag
   	Dim strSUBLIST
   	Dim strCLIENTSUBLIST
   	Dim intSUBRow
   	Dim strMONCNT
   	Dim strLIST
   	Dim tmon, fmon
	Dim strYEARMONLAST
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		strSUBLIST = ""
		chkflag = 1
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTSUBLIST=""
		
		strCLIENTSUBLIST = 	split(mClientsubcode,"♥")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
		
		IF .txtFYEARMON.value > .txtTYEARMON.value THEN
			gErrMsgBox "시작월이 더클수 없습니다. 다시입력하세요.",""
			exit sub
		END IF
		
		strYEARMONLAST = ""
		fmon = mid(.txtFYEARMON.value,1,4) & "-" & mid(.txtFYEARMON.value,5,2) & "-" & "01"
		tmon = mid(.txtTYEARMON.value,1,4) & "-" & mid(.txtTYEARMON.value,5,2) & "-" & "01"
		'종료월과 시작월의 달수를 가져옴
		strMONCNT = datediff("m",fmon,tmon) 
		
		'시작년도와 종료년도가 같다면
		if mid(.txtFYEARMON.value,1,4) = mid(.txtTYEARMON.value,1,4) then
			FOR i=0 TO strMONCNT
				IF i=0 THEN
					strLIST = .txtFYEARMON.value + i
				ELSE
					strLIST = strLIST & "|" & .txtFYEARMON.value + i
				END IF
			NEXT
		else
			for i=0 to strMONCNT
				IF i=0 THEN
					strLIST = .txtFYEARMON.value + i
					strYEARMONLAST = .txtFYEARMON.value + i
				ELSE
					'붙여지는 마지막 년월이 12월이면 년도에 1을 더한후 1월달로 세팅
					if mid(strYEARMONLAST,5,2) = "12" then
						strLIST = strLIST & "|" & mid(strYEARMONLAST,1,4) + 1 & "01"
						strYEARMONLAST = mid(strYEARMONLAST,1,4) + 1 & "01"
					ELSE
						strLIST = strLIST & "|" & strYEARMONLAST + 1
						strYEARMONLAST = strYEARMONLAST + 1
					END IF
				End if
			Next
		END IF
		
		
		strSUBLIST = replace(strSUBLIST,"없음","")
		vntData = mobjEXECUTE.SelectRtn_MONAMTLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strLIST, .txtCLIENTCODE.value, strSUBLIST)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",intCnt) = "총합계" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		End If
	Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<FORM id="frmThis" method="post" runat="server"> <!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--Top TR Start-->
				<TBODY>
					<TR>
						<TD> <!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="115" background="../../../images/back_p.gIF"
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
												<td class="TITLE">월별 매체별 광고비&nbsp;</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="28"> <!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE> <!--Wait Button End--> <!--Common Button Start-->
										<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
											</TR>
										</TABLE> <!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE> <!--Top Define Table End--> <!--Input Define Table End-->
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
								</TR> <!--TopSplit End--> <!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
										<TABLE class="searchDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" title="년도을삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFYEARMON,txtTYEARMON)">
													년월 기간
												</TD>
												<TD class="SEARCHDATA" width="424" style="WIDTH: 424px"><INPUT class="INPUT" id="txtFYEARMON" title="년도을입력하세요" style="WIDTH: 88px; HEIGHT: 22px"
														type="text" maxLength="6" size="9" name="txtFYEARMON" accessKey="NUM">&nbsp;~
													<INPUT class="INPUT" id="txtTYEARMON" title="년도을입력하세요" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="9" name="txtTYEARMON">
												</TD>
												<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">광고주
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 207px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"  src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
												</TD>
											</TR>
											<tr>
												<TD class="SEARCHLABEL" style="WIDTH: 80px">팀
												</TD>
												<TD id="tdCLIENTSUB" class="SEARCHDATA" colspan="3">
												</TD>
											</tr>
										</TABLE>										
									</TD>
								</TR> <!--Input End--> <!--BodySplit Start-->
								<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="굴림"></FONT></TD>
							</TR> <!--BodySplit End--> <!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="16722">
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
