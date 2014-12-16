<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMODY.aspx.vb" Inherits="PD.PDCMJOBMODY" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>PDCMJOBMODY</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet 정보 --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="456" border="1" style="WIDTH: 456px; HEIGHT: 689px">
			<TR>
				<TD style="WIDTH: 456px" valign="top">
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="57" cellSpacing="0" cellPadding="0" width="456" background="../../../images/PopupBG.gif"
							border="0" >
							<TR>
								<td  align="left" width="456">
									<table cellSpacing="0" cellPadding="0" width="100%" border="1">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/PopupIcon.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												사원&nbsp;조회
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
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="0" width="168"
										border="0">
										<TR>
											<TD><FONT face="굴림"></FONT></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><FONT face="굴림"></FONT></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD style="WIDTH: 1px"><FONT face="굴림"></FONT></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgConfirmOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirm.gif'"
													height="20" alt="자료를 선택합니다." src="../../../images/imgConfirm.gif" width="54" border="0"
													name="imgConfirm"></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="화면을 닫습니다." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="456" border="0" style="WIDTH: 456px">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 259px" vAlign="middle" height="259"><FONT face="굴림">
										<TABLE class="KEY" id="tblKey" style="WIDTH: 455px" cellSpacing="0" cellPadding="0" width="455"
											align="right" border="0">
											<TBODY>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField (txtDEPTCD,txtDEPTNAME)">광고주</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="txtDEPTCD" type="text" size="11" name="txtDEPTCD" style="WIDTH: 140px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">
														JOB&nbsp;명&nbsp;</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtDEPTNAME" style="WIDTH: 140px" type="text" size="18" name="txtDEPTNAME"
															tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														CIC/팀</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="txtEMPNO" type="text" size="11" name="txtEMPNO" style="WIDTH: 140px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														JOB&nbsp;NO&nbsp;</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 140px" type="text" size="18" name="txtEMPNAME"
															tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">브랜드</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text1" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">SUB NO</TD>
													<TD class="DATA" style="HEIGHT: 22px"><INPUT class="INPUT" id="Text8" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">예산</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text2" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">상태</TD>
													<TD class="DATA" style="HEIGHT: 22px"><INPUT class="INPUT" id="Text9" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">의뢰일</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text3" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">매체부문</TD>
													<TD class="DATA" style="HEIGHT: 22px"><INPUT class="INPUT" id="Text10" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">완료예정일</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text4" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">매체분류</TD>
													<TD class="DATA" style="HEIGHT: 23px"><INPUT class="INPUT" id="Text14" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">합의일</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text5" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">신규/수정</TD>
													<TD class="DATA" style="HEIGHT: 25px"><INPUT class="INPUT" id="Text13" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">청구일</TD>
													<TD class="DATA" style="WIDTH: 146px"><INPUT class="INPUT" id="Text6" style="WIDTH: 140px" type="text" size="11" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">정산대상</TD>
													<TD class="DATA" style="HEIGHT: 22px"><INPUT class="INPUT" id="Text12" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">결산일</TD>
													<TD class="DATA"><INPUT class="INPUT" id="Text7" style="WIDTH: 140px" type="text" size="14" name="txtEMPNO"></TD>
													<TD class="LABEL" style="WIDTH: 93px; CURSOR: hand">제작사/외주사</TD>
													<TD class="DATA"><INPUT class="INPUT" id="Text11" style="WIDTH: 140px" tabIndex="1" type="text" size="18"
															name="txtEMPNAME"></TD>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT id="sprSht1" style="WIDTH: 221px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="5847">
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
										<OBJECT id="sprSht2" style="WIDTH: 232px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="6138">
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
								<TD class="BOTTOMSPLIT" id="lblStatus">
									<table cellpadding="0" cellspacing="0" border="0" width="456">
										<tr>
											<td>
												<FONT face="굴림"><IMG id="ImgAddrow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" height="20"
														alt="자료를 선택합니다." src="../../../images/imgAddRow.gif" width="54" border="0" name="imgConfirm"></FONT>
														<br>
														<br>
											</td>
											<td>
												<FONT face="굴림"><IMG id="ImgAddrow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" height="20"
														alt="자료를 선택합니다." src="../../../images/imgAddRow.gif" width="54" border="0" name="imgConfirm"></FONT>
														<br>
														<br>
											</td>
										</tr>
									</table>
								</TD>
							</TR>
						</TABLE>
						<FONT face="굴림"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
