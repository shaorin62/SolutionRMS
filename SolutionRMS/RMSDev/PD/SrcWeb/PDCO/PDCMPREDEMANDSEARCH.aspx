<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPREDEMANDSEARCH.aspx.vb" Inherits="PD.PDCMPREDEMANDSEARCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>선청구 누계현황</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/마감 등록 화면
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
Dim mobjPDCMSEARCH
Dim mobjPDCMGET
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
Dim vntData

with frmThis
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End with
	
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub ImgExeConfirm_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub ImgExeConfirmCancel_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn_Cancel
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
	'mobjPDCMGET
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
	
    Call Grid_Layout()
    frmThis.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
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
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "GBN|CLIENTCODE|CLIENTNAME|YEARMON|JOBNO|JOBNAME|ADJAMT|EMPNAME"
		mobjSCGLSpr.SetHeader .sprSht,		    "구분|광고주코드|광고주명|청구월|JOBNO|JOB명|집행금액|기획담당"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "0|0|35|12|14|35|14|12"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|YEARMON|EMPNAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME|JOBNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CLIENTCODE|CLIENTNAME|YEARMON|JOBNO|JOBNAME|ADJAMT|EMPNAME"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"CLIENTNAME"
	End with
	'DateClean
	pnlTab1.style.visibility = "visible" 
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

Sub txtYEARMON_onchange
	gSetChange
End Sub
Sub txtTo_onchange
	gSetChange
End Sub


Sub SelectRtn ()

   	Dim vntData
   	Dim i, strCols
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMSEARCH.SelectRtn_PREDEMAND(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtCLIENTCODE.value,.txtCLIENTNAME.value )
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 1 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows 
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "2" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
				end if
				If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "3" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
				End If
			Next
			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			Else
   			initpageData
   			gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   			end If
   			
   		end if
   	end with
End Sub



Sub EndPage()
	set mobjPDCMSEARCH = Nothing
	set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
	End with
End Sub
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
		
				
     		'GetBrandDefaultFind	
 
			'.cmbSEARCHJOBGUBN.focus()					' 포커스 이동
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
     	
	End with

	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
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
												<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
											<td class="TITLE">선청구 누계현황&nbsp;</td>
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
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="110" border="0">
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
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="95%" border="0"> <!--TopSplit Start->
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="년도을삭제합니다." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">
												결산월
											</TD>
											<TD class="SEARCHDATA" width="424" style="WIDTH: 424px"><INPUT class="INPUT" id="txtYEARMON" title="년도을입력하세요" style="WIDTH: 100px; HEIGHT: 22px"
													type="text" maxLength="6" size="14" name="txtYEARMON" accessKey="NUM">
											</TD>
											<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">광고주
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 207px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31882">
											<PARAM NAME="_ExtentY" VALUE="17542">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
								</TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
			</TABLE>
		</FORM>
		</TR></TABLE>
	</body>
</HTML>
