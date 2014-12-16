<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCOTEST.aspx.vb" Inherits="PD.PDCOTEST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CATV 수수료 거래명세표 발행</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mobjPDCO_TEST
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode


'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	'EndPage
End Sub



'엑셀버튼 클릭시 
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub


'프린트버튼 클릭시
SUB imgPrint_onclick()
	gFlowWait meWAIT_ON
	With frmThis
'-----------------------------크리스탈 레포트연결
'-----------------------------크리스탈 레포트연결
'-----------------------------크리스탈 레포트연결
	End With
	gFlowWait meWAIT_OFF
END SUB

'조회버튼 클릭시
sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
end sub

'신규버튼 클릭시
sub imgNew_onclick ()
	InitPageData
end sub


'저장버튼 클릭시
sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
end sub



'부서코드 버튼 클릭시
Sub ImgDEPTCODE_onclick ()
	Call DEPTCODE_POP()
End Sub


'부서팝업 띄우기
Sub DEPTCODE_POP ()
	Dim vntRet, vntInParams
	with frmThis
		vntInParams = array(trim(.txtDEPTCD.value))
		vntRet = gShowModalWindow("../PDCO/PDCMDEPTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
		    .txtDEPTCD.value = vntRet(0,0)	'Code값 저장
			.txtDEPTNAME.value = vntRet(1,0)	'코드명 표시
			'txtDEPTCD_onchange
			'txtDEPTNAME_onchange
			'.txtATTR02.focus()
		gSetChangeFlag .txtDEPTCD
		end if
	End with
END SUB


'사원코드버튼 클릭시
SUB ImgEMPNOCODE_onclick ()
	call EMPNOCODE_POP ()
END SUB

'사원팝업 띄우기
SUB EMPNOCODE_POP ()
	Dim vntRet, vntInParams
	
	with frmThis
		vntInParams = array("","",trim(.txtEMPNO.value),trim(.txtEMPNAME.value))
		vntRet = gShowModalWindow("../PDCO/PDCMEMPPOP.aspx",vntInParams, 413,425)
		
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEMPNO.value = vntRet(0,0)		             ' Code값 저장
			.txtEMPNAME.value = vntRet(1,0)             ' 코드명 표시
		'gSetChangeFlag .txtCUSTCODE1                      ' gSetChangeFlag objectID	 Flag 변경 알림
    end if
	end with
	
END SUB



'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)

	With frmThis
		'JOBNO에 넣기
		.txtJOBNO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",Row)
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		.txtCREPART.value = mobjSCGLSpr.GetTextBinding( .sprSht,"CREPART",Row)
		.txtDEPTCD.value = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTCD",Row)
		.txtDEPTNAME.value = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",Row)
		.txtEMPNAME.value = mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNAME",Row)
		.txtEMPNO.value = mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNO",Row)
		.txtJOBGUBN.value = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBN",Row)
	End With
End Sub  


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		
		vntData = mobjPDCO_TEST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNAMESEARCH.value)
		if not gDoErrorRtn ("SelectRtn_HDR") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			

   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			sprSht_Click 0,0
   			
   		end if
   	end with
end sub



sub ProcessRtn ()
	Dim intRtn
   	Dim strMasterData
	
	with frmThis
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
	
		intRtn = mobjPDCO_TEST.ProcessRtn(gstrConfigXml,strMasterData)

		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		
			SelectRtn
   		end if
   		
   	end with
end sub

'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	

	'서버업무객체 생성	
	set mobjPDCO_TEST = gCreateRemoteObject("cPDCO.ccPDCOTEST")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "210px"
	pnlTab1.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		'*********************************
		'수수료시트
		'*********************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK |JOBNO | JOBNAME | REQDAY | BUDGETAMT | BIGO |EMPNO |DEPTCD |JOBGUBN |CREPART |EMPNAME | DEPTNAME"
											 
		mobjSCGLSpr.SetHeader .sprSht,		 "선택 | JOBNO | JOBNAME | 의뢰일 | 예산금액  | 비고 |담당자|담당팀| 매체부문| 매체분류 | 담당자이름| 당당팀이름"
											    
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  6 | 20|        20|    20 |          30|   30|    0|     0|        0|        0|           0|          0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		
		
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "BUDGETAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "JOBNO | JOBNAME | BIGO|EMPNO |DEPTCD |JOBGUBN |CREPART |EMPNAME | DEPTNAME", -1, -1, 1000
		
	
    End With
    
	pnlTab1.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData
	
End Sub

Sub EndPage()
	set mobjPDCOTEST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	with frmThis
		.sprSht.MaxRows = 0
		.txtJOBNO.focus	
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	'초기 데이터 설정
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
				border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp; 제작 관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Top Define Table End-->
								<!--Input Define Table End--></TABLE>
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)">제작명&nbsp;</TD>
												<TD class="SEARCHDATA" style="WIDTH: 201px" width="201"><INPUT class="INPUT" id="txtJOBNAMESEARCH" title="제작명" style="WIDTH: 336px; HEIGHT: 22px"
														type="text" maxLength="20" size="50" name="txtJOBNAMESEARCH"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"
													width="80"><FONT face="굴림"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
															src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></FONT></TD>
												<TD class="SEARCHDATA" width="120">&nbsp;
												</TD>
												<TD class="SEARCHDATA" width="260"></TD>
												<TD class="SEARCHDATA">&nbsp;&nbsp;</TD>
												<td class="SEARCHDATA" width="50"></td>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="굴림"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;&nbsp;&nbsp;</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" height="20" alt="자료를 인쇄합니다."
																	src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
															<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></td>
															<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
																	name="imgPrint"></TD>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
										<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 55px" cellSpacing="1" cellPadding="0"
											align="right" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 84px; HEIGHT: 25px">
													<P align="center">JOB NO</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 18.24pt"><INPUT dataFld="JOBNO" class="INPUT_L" id="txtJOBNO" title="JOBNO" style="WIDTH: 176px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="24" name="txtJOBNO"></TD>
												<TD class="LABEL" style="HEIGHT: 25px" align="center">
													<P align="center">JOB 명</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="JOBNAME" class="INPUT_L" id="txtJOBNAME" title="JOB명" style="WIDTH: 176px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="24" name="txtJOBNAME"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 84px; HEIGHT: 25px">
													<P align="center">담당자</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="EMPNAME" class="INPUT_L" id="txtEMPNAME" title="담당자" style="WIDTH: 176px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="24" name="txtEMPNAME"><IMG id="ImgEMPNOCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgEMPNOCODE"><INPUT dataFld="EMPNO" class="NOINPUT" id="txtEMPNO" title="광고주코드" style="WIDTH: 72px"
														accessKey=",M" dataSrc="#xmlBind" readOnly type="text" size="6" name="txtEMPNO">
												</TD>
												<TD class="LABEL" style="HEIGHT: 25px" align="center">
													<P align="center">담당팀</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="DEPTNAME" class="INPUT_L" id="txtDEPTNAME" title="담당팀" style="WIDTH: 176px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="24" name="txtDEPTNAME"><IMG id="ImgDEPTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgDEPTCDCODE"><INPUT dataFld="DEPTCD" class="NOINPUT" id="txtDEPTCD" title="광고주코드" style="WIDTH: 72px"
														accessKey=",M" dataSrc="#xmlBind" readOnly type="text" size="6" name="txtDEPTCD"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 84px; HEIGHT: 25px">
													<P align="center">매체부문</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="JOBGUBN" class="INPUT_L" id="txtJOBGUBN" title="매체부문" style="WIDTH: 176px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="24" name="txtJOBGUBN"></TD>
												<TD class="LABEL" style="HEIGHT: 25px">
													<P align="center">매체분류</P>
												</TD>
												<TD class="DATA" style="HEIGHT: 19pt">
													<P align="left"><INPUT dataFld="CREPART" class="INPUT_L" id="txtCREPART" title="매체분류" style="WIDTH: 176px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" size="24" name="txtCREPART"></P>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 15px">
										<P>&nbsp;</P>
									</TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 654px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 1038px; HEIGHT: 630px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27464">
												<PARAM NAME="_ExtentY" VALUE="16669">
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
								<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;합 
							계 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="금액" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
								<!--List End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
