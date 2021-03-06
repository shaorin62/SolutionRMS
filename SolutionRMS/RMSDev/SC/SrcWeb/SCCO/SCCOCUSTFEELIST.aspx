<%@ Page CodeBehind="SCCOCUSTFEELIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="SC.SCCOCUSTFEELIST" %>
<HTML>
	<HEAD>
		<title>거래유형 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 광고처관리 (실적) 
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 거래처 대한 MAIN 정보를 조회/저장/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/25 By hwang duck-su
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
Dim mobjSCCOCUSTGREATLIST '공통코드, 클래스
Dim mobjMDCMGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' 명령 버튼 클릭 이벤트
'---------------------------------------------------

'-----------------------------------
'조회
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------
'HDR수정저장    -
'-----------------------------------
Sub imgSave_onclick ()
	IF frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub



'-----------------------------
' 거래선 팝업 조회 
'-----------------------------
Sub ImgCUST_ID_onclick	
	CLIENTPOP
End Sub

Sub CLIENTPOP
	Dim vntRet
	Dim vntInParams
	Dim strMEDFLAG
	strMEDFLAG =""
	with frmThis
	
	strMEDFLAG = "K"
	

	vntInParams = array(.txtCUSTCODE.value, .txtCUSTNAME.value,strMEDFLAG)
	vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)
		
	if isArray(vntRet) then
		if .txtCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		.txtCUSTCODE.value = vntRet(0,0)		             ' Code값 저장
		.txtCUSTNAME.value = vntRet(1,0)             ' 코드명 표시
		
		gSetChangeFlag .txtCUSTCODE                      ' gSetChangeFlag objectID	 Flag 변경 알림
    end if
			
	End with
	
	gSetChange
End Sub


'

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================

'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjSCCOCUSTGREATLIST = gCreateRemoteObject("cSCCO.ccSCCOCUSTGREATLIST")
	'set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET") 아직 쓰지 않음
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis

                
	gSetSheetColor mobjSCGLSpr, .sprSht	
	mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 3, 0,0
	mobjSCGLSpr.AddCellSpan  .sprSht, 2, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.SpreadDataField .sprSht, "BISINO|COMPANYNAME|CUSTNAME| CUSTCODE | DAY    |       A|       B|      C "
	mobjSCGLSpr.SetHeader .sprSht, " 사업자번호 |     상호명|거래처명|거래처코드|적용기간|거래유형|대행형태|적용매체"
	mobjSCGLSpr.SetColWidth .sprSht, "-1", "  20|         15|      10|        10|       10|      10|     10|      10"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
	'mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
	
	'mobjSCGLSpr.SetCellTypeComboBox2 .sprSht , -1,"2",2,1,false,false
	
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "", -1, -1, 200
	mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,0,2,false '왼쪽
	mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,2,2,false '가운데
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,""
	mobjSCGLSpr.colhidden .sprSht, "",true
	.sprSht.style.visibility = "visible"

    End With

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOCUSTGREATLIST = Nothing
	set mobjMDCMGET = Nothing
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
		.sprSht.MaxRows = 0
	End With

End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim Flag
   	
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOCUSTGREATLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, .txtCUSTNAME.value, .txtBUSINO.value )

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub


'------------------------------------------
' 수정/저장 처리 
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
   	Dim strRow
	Dim intCnt
	Dim strHIGHCUSTCODE

	with frmThis
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"COMPANYNAME | GREATCODE | BTN | GREATNAME")
		
	
		if  not IsArray(vntData) then	
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		strRow = .sprSht.ActiveRow
		strHIGHCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"HIGHCUSTCODE",strRow)
	

		'intRtn = mobjSCCOCUSTGREATLIST.ProcessRtn(gstrConfigXml,vntData, strHIGHCUSTCODE)
	
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			strRow = .sprSht.ActiveRow
			SelectRtn			
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		end if
   	end with
End Sub





'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	msgbox 1
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub



'-----------------------------------
'쉬트안의 팝업
'-----------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strRow
	Dim strGREATCODE
	
	with frmThis

		IF Col = 3 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			
			
			vntInParams = array("","")
			vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)

				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
			
		end if
		.sprSht.focus 
	End with
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
						<TD style="HEIGHT: 100%">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">거래처유형 관리</td>
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
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="90%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 17px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="1024" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 76px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,txtCUSTCODE)">거래처명</TD>
												<TD class="SEARCHDATA" style="WIDTH: 250px"><INPUT class="INPUT_L" id="txtCUSTNAME" title="코드명" style="HEIGHT: 22px" type="text" maxLength="100"
														align="left" size="18" name="txtCUSTNAME"><IMG id="ImgCUST_ID" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
														width="23" align="absMiddle" border="0" name="ImgCUST_ID"><INPUT class="INPUT_L" id="txtCUSTCODE" title="코드조회" style="WIDTH: 55px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="6" name="txtCUSTCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')">거래서차업자번호</TD>
												<TD class="SEARCHDATA" style="WIDTH: 50px"><INPUT dataFld="BUSINO" class="INPUT_L" id="txtBUSINO" title="코드조회" style="WIDTH: 128px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="15" align="left" size="16" name="txtBUSINO"></TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
														src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<tr>
									<TD>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27120">
												<PARAM NAME="_ExtentY" VALUE="14235">
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
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
								</TR>
								<!--Bottom Split End-->
							</TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
