<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMATTER.aspx.vb" Inherits="MD.MDCMMATTER" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>소재 일괄투입</title>
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
		<script language="vbscript" id="clientEventHandlersVBS">
'전역변수 설정
Dim mobjMDCMPREMATTER
Dim mobjMDCMMEDGet
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag

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
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
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
	Set mobjMDCMPREMATTER = gCreateRemoteObject("cMDET.ccMDETPREMATTER")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'화면 초기값 설정
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	Dim strComboList 
	strComboList =  "사용" & vbTab & "미사용"
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		'CC_CODE,CC_NAME,OC_CODE,OC_NAME,USE_YN,STDATE,EDATE
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0, 0
		'mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "CLIENTNAME|KOBACOCODE|MATTER|MATTERDIV|CMLAN|SISCUNO|DISCUDATE|MATTERVIEW|SUBSEQNAME|WASTECODE|LCODE|LNAME|MCODE|MNAME|SNAME|TELECASTLIMIT|TELECASTTIME|CUSER|CDATE|CLIENTCODE|ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht,		    "광고주명|광고주코드|소재|소재구분|초수|심의번호|심의일자|소재보기|품목|코드|업종대분류코드|업종대분류명|업종중분류코드|업종중분류명|업종소분류명|방영제한|시간|입력자|입력일자|CLIENTCODE|오류내용"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "10      |10        |10  |10      |10  |10      |10      |10      |10  |10  |12            |10          |12            |10          |10          |10      |10  |10    |10      |10         |10"         
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		'mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "SDATE|EDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME|KOBACOCODE|MATTER|MATTERDIV|CMLAN|SISCUNO|DISCUDATE|MATTERVIEW|SUBSEQNAME|WASTECODE|LCODE|LNAME|MCODE|MNAME|SNAME|TELECASTLIMIT|TELECASTTIME|CUSER|CDATE|CLIENTCODE|ERRMSG", -1, -1,200
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,2,2,false '중앙정렬
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,0,2,false '왼쪽정렬
		'mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CC_CODE|CC_NAME|OC_CODE|OC_NAME"
		'mobjSCGLSpr.SetCellTypeComboBox .sprSht,6,6,,,strComboList
		'mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", true
	End with

	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .chkAll.checked = True Then
		strCHK = ""
		Else
		strCHK = "All"
		End if
		
		vntData = mobjMDCMPREMATTER.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTCODE.value,.txtDEPTNAME.value,.cmbYN.value,strCHK)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtDEPTCODE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub
'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
End Sub







Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'여기까지 쉬트 버튼 클릭

'Validation
Function DataValidation ()
	DataValidation = false	
	With frmThis
		
	End With
	DataValidation = True
End Function
'저장로직

Sub ProcessRtn()
	Dim intRtn
   	dim vntData
	Dim intCnt
	Dim intDelCnt
	Dim strCODE
	Dim intErrCnt
	Dim intRtnYN
		with frmThis
   		'데이터 Validation 영역
		'if DataValidation =false then exit sub
		'DataErrorValidation
		'On error resume next
		'잔여Row 삭제 처리
		intRtnYN = gYesNoMsgbox("자료를 생성하시겠습니까?","자료삭제 확인")
		IF intRtnYN <> vbYes then exit Sub
		for intDelCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,1, intDelCnt) = "" Then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End IF
		Next
		
		'광고주코드 가져오기
		strCODE = ""
		for intCnt = 1 To .sprSht.MaxRows 
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"KOBACOCODE", intCnt)
		vntData = mobjMDCMPREMATTER.GetCLIENTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
			if mlngRowCnt > 0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",intCnt, vntData(0,1)
			Else
			mobjSCGLSpr.SetTextBinding .sprSht,"ERRMSG",intCnt, "광고주코드에러"
   			End if
		next
		
		for intErrCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"ERRMSG", intErrCnt) <> "" Then
				gErrorMsgbox "오류내역을 확인하십시오.","저장안내!"
				exit Sub 
			end if
		Next
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"MATTER|SUBSEQNAME|CLIENTCODE")
		'처리 업무객체 호출
		intRtn = mobjMDCMPREMATTER.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 1 Then
			gErrorMsgBox intRtn & " 건 이 저장되었습니다." & vbcrlf & "브랜드 코드와 소재 코드가 자동생성되었으니," & vbcrlf & "브랜드 코드 의 MC 부서를 연결하여 주시고," & vbcrlf & "소재코드 와 대대행사 를 연결 하여 주십시오.","저장안내"
			End If
			'SelectRtn
   		end if
   	end with
End Sub
Sub EndPage()
	set mobjMDCMPREMATTER = Nothing
	'set mobjMDCMMEDGet = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 2000
	End with
End Sub

sub DeleteRtn

	
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;소재 일괄투입</td>
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
									<TABLE id="tblButton" style="WIDTH: 183px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="183" border="0">
										<TR>
											<TD></TD>
											<TD width="54"></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD width="54"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
											<!--<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>--></TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 1040px; HEIGHT: 32px" cellSpacing="0" cellPadding="0"
							width="1040" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center"
									colSpan="2"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 14px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')">&nbsp;</TD>
												<TD class="SEARCHDATA" style="WIDTH: 172px"><INPUT class="INPUTL" id="txtDEPTCODE" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="8" size="10" name="txtDEPTCODE"></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 2px"></TD>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD class="DATAFRAME">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 683px"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 683px"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27490">
											<PARAM NAME="_ExtentY" VALUE="18071">
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
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR> 
			<!--List End-->
			<!--BodySplit Start-->
			<TR>
				<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
			</TR>
			<!--BodySplit End-->
			<!--Brench Start-->
			<TR>
				<TD class="BRANCHFRAME" style="WIDTH: 1040px"><FONT face="굴림" color="#666666" size="3"></FONT>
					<!--<INPUT class="BUTTON" id="btn1" style="WIDTH: 123px; HEIGHT: 16pt" type="button" value="분기버튼"
											name="Button">--></TD>
			</TR>
			<!--Brench End-->
			<!--Bottom Split Start-->
			<TR>
				<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
			</TR>
			<!--Bottom Split End--> </TABLE> 
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
