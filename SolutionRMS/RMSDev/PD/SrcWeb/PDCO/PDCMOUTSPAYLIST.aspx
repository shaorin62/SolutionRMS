<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMOUTSPAYLIST.aspx.vb" Inherits="PD.PDCMOUTSPAYLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>외주비 지급현황</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 외주비 지급현황
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMOUTSPAYLIST.aspx
'기      능 : 외주비 내역을 조회한다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2010/01/20 By Taeho Kim
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
		<script language="vbscript" id="clientEventHandlersVBS">
'전역변수 설정
Dim mobjPDCMSEARCH
Dim mobjPDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mobjPDCMJOBNO
Dim mobjSCCOGET

CONST meTAB = 9
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'서버업무객체 생성	
	Set mobjPDCMSEARCH = gCreateRemoteObject("cPDCO.ccPDCOSEARCH")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	Set mobjPDCMJOBNO = gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'화면 초기값 설정
	
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0, 5
		mobjSCGLSpr.SpreadDataField .sprSht,    "OUTSCODE | OUTSNAME | JOBNO | JOBNAME | CLIENTCODE | CLIENTNAME | ADJAMT | CONTRACTNO | TAXFLAG | TAXDATE | GBN | AMTFLAG | VOCHNO | DEMANDDAY | DOCUMENTDATE | PAYCODE"
		mobjSCGLSpr.SetHeader .sprSht,		    "외주처코드|외주처명|JOBNO|JOB명|광고주코드|광고주명|정산금액|계약번호|발행여부|발행날짜|SORT|하도급|전표번호|지급기일|증빙일|지급방법"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "         0|      23|    8|   25|         0|      15|      13|      13|       5|       9|   0|     0|       0|       0|     0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "AMTFLAG "
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TAXDATE | DEMANDDAY | DOCUMENTDATE"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		'mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO | CONTRACTNO | TAXFLAG | VOCHNO",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME | OUTSNAME | JOBNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"OUTSCODE | OUTSNAME | JOBNO | JOBNAME | CLIENTCODE | CLIENTNAME | ADJAMT | CONTRACTNO | TAXFLAG | TAXDATE | AMTFLAG | VOCHNO | DEMANDDAY | DOCUMENTDATE | PAYCODE"
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE | CLIENTCODE | GBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSCODE | OUTSNAME"
	End with
	DateClean
	pnlTab1.style.visibility = "visible" 
End Sub
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
'-----------------------------------------------------------------------------------------
' 광고주팝업(조회)
'-----------------------------------------------------------------------------------------
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)	
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),"A")
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
				
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub imgOUTSCODE_onclick
	Call SEARCHOUT_POP()
End Sub
'-----------------------------------------------------------------------------------------
' 외주처팝업(조회)
'-----------------------------------------------------------------------------------------
'실제 데이터List 가져오기
Sub SEARCHOUT_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtOUTSNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
				Else
				
					Call SEARCHOUT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'검색조건 시작일
Sub imgFrom_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtFrom,.imgFrom,"txtFrom_onchange()"
		gSetChange
	end with
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub

'검색조건 종료일
Sub imgTo_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtTo,.imgTo,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtTo_onchange
	gSetChange
End Sub


Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strFROM
    Dim strTO
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		vntData = mobjPDCMSEARCH.SelectRtn_PAYLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,.txtCLIENTCODE.value,.txtCLIENTNAME.value,.txtOUTSCODE.value,.txtOUTSNAME.value,.cmbTAXFLAG.value)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 1 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht.MaxRows 
					
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
					If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "2"  Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					end if
				
				Next
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate2
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFrom.value = date1
		.txtTo.value = date2
	End With
End Sub
Sub EndPage()
	set mobjPDCMSEARCH = Nothing
	set mobjPDCMGET = Nothing
	set mobjPDCMJOBNO = Nothing
	set mobjSCCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0
	End with
End Sub


		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" HEIGHT="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--1단-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="97" background="../../../images/back_p.gIF"
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
											<td class="TITLE">외주비 지급현황</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="50" border="0">
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
								</TD>
							</TR>
						</TABLE>
						<!--1단끝-->
						<!--2단-->
						<TABLE id="tblBody" style="WIDTH: 100%;HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()" width="60">&nbsp;등록월
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 209px"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="8" size="6" name="txtFROM"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
														align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="8" size="6" name="txtTO"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
														align="absMiddle" border="0" name="imgTo"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
													width="60">광고주
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 240px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 155px; HEIGHT: 22px"
														type="text" maxLength="255" align="left" size="26" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
													<INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
												<TD class="SEARCHLABEL" width="60" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME,txtOUTSCODE)">외주처</TD>
												<TD class="SEARCHDATA" style="WIDTH: 240px"><INPUT class="INPUT_L" id="txtOUTSNAME" title="코드명" style="WIDTH: 155px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="38" name="txtOUTSNAME"> <IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" width="12" align="absMiddle" border="0"
														name="ImgOUTSCODE"> <INPUT class="INPUT_L" id="txtOUTSCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtOUTSCODE"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbENDFLAG1,'')"
													width="50">발행유무</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbTAXFLAG" title="발행유무" style="WIDTH: 70px" name="cmbTAXFLAG">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="Y">발행</OPTION>
														<OPTION value="N">미발행</OPTION>
													</SELECT></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR vAlign="top" align="left">
								<TD height="100%">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 96%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="42466">
											<PARAM NAME="_ExtentY" VALUE="19315">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
						</TABLE>
						<!--2단끝--></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" style="WIDTH: 100%" id="lblstatus"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
