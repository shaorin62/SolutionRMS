<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCUSTSEQPOP_ALL.aspx.vb" Inherits="MD.MDCMCUSTSEQPOP_ALL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>브랜드 조회</title> 
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/15 By Kim Tae Yub
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mobjMDCOGET
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
end sub

Sub txtCUSTCODE_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtSUBSEQ_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtSUBSEQNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick ()
	if frmThis.sprSht.ActiveRow > 0 then
		sprSht_DblClick frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	else
		call Window_OnUnload()
	end if
end sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

sub sprSht_DblClick (Col,Row)
	'선택된 로우 반환
	With frmThis
	if Row = 0 and Col >0 then
		mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
	Else
		window.returnvalue = mobjSCGLSpr.GetClip (.sprSht,1,.sprSht.ActiveRow,.sprSht.MaxCols,1,1)
		call Window_OnUnload()
	end if
	End With
end sub

Sub sprSht_Keydown(KeyCode, Shift)
    if KeyCode <> meCR then exit sub
	'시트에서 엔터시 확인 처리
	Call sprSht_DblClick (frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow)		
End Sub
'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMDCOGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : .txtSUBSEQ.value = vntInParam(i)	
				case 1 : .txtSUBSEQNAME.value = vntInParam(i)
				case 2 : .txtCUSTCODE.value = vntInParam(i)			'조회추가필드
				case 3 : .txtCUSTNAME.value = vntInParam(i)		'현재 사용중인 것만
				case 4 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 5 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
		'txtCUSTNAME,txtCUSTCODE
		'SpreadSheet 디자인
		gSetSheetDefaultColor	'Sheet 기본Color 지정
		gSetSheetColor mobjSCGLSpr, .sprSht	'Sheet 칼라 지정
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht, " SEQNO | SEQNAME | CUSTCODE | CUSTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | CLIENTSUBCODE | CLIENTSUBNAME | DEPT_CD | DEPT_NAME "
		mobjSCGLSpr.SetHeader .sprSht,	  "브랜드코드|브랜드명|광고주코드|광고주명|광고처코드|광고처명|팀코드|팀명|사업부코드|사업부|부서코드|부서명"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "    0|      10|         0|      12|         0|      12|     0|  10|         0|    10|       0|    15"
		mobjSCGLSpr.SetRowHeight .sprSht,"-1","13"
		mobjSCGLSpr.SetRowHeight .sprSht,"0","15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"SEQNO"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"SEQNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CUSTCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CUSTNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"GREATCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"GREATNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"TIMCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"TIMNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CLIENTSUBCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CLIENTSUBNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"DEPT_CD"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"DEPT_NAME"
		mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
		mobjSCGLSpr.ColHidden .sprSht, "SEQNO | CUSTCODE | GREATCODE | TIMCODE | CLIENTSUBCODE | DEPT_CD ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQNAME | CUSTNAME | GREATNAME | TIMNAME | CLIENTSUBNAME | DEPT_NAME" ,-1,-1,0,2,false

		
	end with	
	'자료조회	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCOGET.Get_BrandInfo_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtSUBSEQ.value,.txtSUBSEQNAME.value,.txtCUSTCODE.value,.txtCUSTNAME.value)

		if not gDoErrorRtn ("Get_BrandInfo") then
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			if mlngRowCnt <> 0 then
   				.sprSht.focus()
   			else
   				.sprSht.MaxRows = 0
   				.txtSUBSEQNAME.focus()
   			end if 
   		end if
   	end with
end sub

-->
		</script>
	</HEAD>
	<body class="base"  bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="620" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" valign=bottom>
												브랜드 검색
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 168px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgConfirmOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirm.gif'"
													height="20" alt="자료를 선택합니다." src="../../../images/imgConfirm.gif" width="54" border="0"
													name="imgConfirm"></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="화면을 닫습니다." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</tr>
						</table>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 382px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="DATA" vAlign="middle" height="21" style="WIDTH: 620px"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="0" cellPadding="0" width="620" align="right" border="0">
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,txtCUSTCODE)">
														광고주코드&nbsp;</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtCUSTCODE" type="text" size="9" name="txtCUSTCODE" style="WIDTH: 140px; HEIGHT: 22px"></TD>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,txtCUSTCODE)">광고주명</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtCUSTNAME" style="WIDTH: 140px; HEIGHT: 22px" type="text" size="18"
															name="txtCUSTNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME,txtSUBSEQ)">
														브랜드코드&nbsp;</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtSUBSEQ" type="text" size="9" name="txtSUBSEQ" style="WIDTH: 140px; HEIGHT: 22px"></TD>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME,txtSUBSEQ)">브랜드명</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtSUBSEQNAME" style="WIDTH: 140px; HEIGHT: 22px" type="text"
															size="18" name="txtSUBSEQNAME" tabIndex="1"></TD>
												</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 620px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD align="center" style="WIDTH: 620px"><FONT face="굴림">
										<OBJECT id="sprSht" style="WIDTH: 620px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="13229">
											<PARAM NAME="_ExtentY" VALUE="7250">
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
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 620px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="굴림"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
