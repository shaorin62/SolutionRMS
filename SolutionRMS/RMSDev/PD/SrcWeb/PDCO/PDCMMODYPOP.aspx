<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMMODYPOP.aspx.vb" Inherits="PD.MODY" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>MODY</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet ���� --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
		
		option explicit
			
		Sub window_onload()
			InitPage
		end sub

		
		SUB InitPage ()
		with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		'vntInParam = window.dialogArguments
		'intNo = ubound(vntInParam)
		'�⺻�� ����
		'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		'for i = 0 to intNo
		'	select case i
		'		case 0 : .txtDEPTCD.value = vntInParam(i)	
		'		case 1 : .txtDEPTNAME.value = vntInParam(i)
		'		case 2 : .txtEMPNO.value = vntInParam(i)			'��ȸ�߰��ʵ�
		'		case 3 : .txtEMPNAME.value = vntInParam(i)		'���� ������� �͸�
		'		case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
		'		case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
		'	end select
		'next
		'SpreadSheet ������
		'gSetSheetDefaultColor()
        With frmThis
            gSetSheetColor mobjSCGLSpr, .sprSht1
			mobjSCGLSpr.SpreadLayout .sprSht1, 2, 0, 0, 0, , 2, 1, , , True
			mobjSCGLSpr.SpreadDataField .sprSht1, " EMP_NAME | EMP_NAME |"
			
			mobjSCGLSpr.SetHeader .sprSht1, "�����",0,1,true
			mobjSCGLSpr.SetHeader .sprSht1, "����ڸ�|�Ҽ���" ,SPREAD_HEADER + 1,1,true
			
			mobjSCGLSpr.AddCellSpan .sprSht1, 1, SPREAD_HEADER + 0, -1 ,1 
			
			
			
			
			gSetSheetColor mobjSCGLSpr, .sprSht2
			mobjSCGLSpr.SpreadLayout .sprSht2, 2, 0, 0, 0, , 2, 1, , , True
			mobjSCGLSpr.SpreadDataField .sprSht2, " EMP_NAME | EMP_NAME |"
			
			mobjSCGLSpr.SetHeader .sprSht2, "�����μ�",0,1,true
			mobjSCGLSpr.SetHeader .sprSht2, "�μ���|�й���" ,SPREAD_HEADER + 1,1,true
			
			mobjSCGLSpr.AddCellSpan .sprSht2, 1, SPREAD_HEADER + 0, -1 ,1 
			
			
			
			'mobjSCGLSpr.SpreadLayout .sprSht, intLayOutCnt, 0, 0, 0, , 2, 1, , , True
			'mobjSCGLSpr.SpreadDataField .sprSht, mstrField 
			'mobjSCGLSpr.SetHeader .sprSht,       strStartHead ,0,1,true
			'mobjSCGLSpr.SetHeader .sprSht,       strEndHead ,SPREAD_HEADER + 1,1,true
			
			'mobjSCGLSpr.AddCellSpan .sprSht, 1, SPREAD_HEADER + 0, 1    , 2      , -1 , true
			'mobjSCGLSpr.AddCellSpan .sprSht, 2, SPREAD_HEADER + 0, mvntDataEXCLIENTCNT    , 1      , -1 , true
			'                                 20��° ����            ����6���� 1���� 3�������� ������
			'mobjSCGLSpr.AddCellSpan .sprSht, intLayOutCnt-1, SPREAD_HEADER + 0, 1    , 2      , -1 , true
			'                                 ������ Ǯ���°� �� 44��°�̰� 2���� ���Ķ� -1 ��ü
			'mobjSCGLSpr.SetColWidth .sprSht, "-1", strEndWith
			'mobjSCGLSpr.SetCellTypeFloat2 .sprSht, mstrField, -1, -1, 0
			'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME|REAL_MED_NAME", , , 50, , ,0
			'mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
			'mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			'mobjSCGLSpr.SetCellsLock2 .sprSht,true,strField
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|REAL_MED_NAME",-1,-1,2,2,false

        End With
	end with	
	'�ڷ���ȸ	

		END SUB
		
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="392" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/PopupIcon.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												JOB&nbsp;���/����
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="168" border="0">
										<TR>
											<TD><FONT face="����"></FONT></TD>
											<TD style="WIDTH: 78px"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
													border="0" name="imgSave"></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="����">
										<TABLE class="DATA" id="tblKey" style="WIDTH: 392px" cellSpacing="1" cellPadding="0" width="392"
											align="right" border="0">
											<TBODY>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">
														�μ� �ڵ�</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtDEPTCD" type="text" size="9" name="txtDEPTCD" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">
														JOB��</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtDEPTNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtDEPTNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														CLC/��</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														JOBNO</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														�귣��</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														SUBNO&nbsp;</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														�Ƿ���</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														��ü�ι�</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														�ϷΌ����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														��ü�з�</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														������</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														�ű�/����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														û����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														������</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														�����</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 100px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="LABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
														���ۻ�/����ó</TD>
													<TD class="DATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="18"
															name="txtEMPNAME" tabIndex="1"></TD>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR width="392">
								<td id="SPRT">
									<TABLE border="0" cellpadding="0" cellspacing="0">
										<tr>
											<TD align="center"><FONT face="����">
													<OBJECT id="sprSht1" style="WIDTH: 196px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="5186">
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
											<td align="center"><font face="����">
													<OBJECT id="sprSht2" style="WIDTH: 196px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="5186">
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
												</font>
											</td>
										</tr>
									</TABLE>
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����">
										<table cellpadding="0" cellspacing="0" border="0" width="100%">
											<tr>
												<td>
													<FONT face="����"><IMG id="ImgAddrow1" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="�ڷḦ �����մϴ�."
															src="../../../images/imgAddRow.gif" width="54" border="0" name="imgConfirm"></FONT>
												</td>
												<td>
													<FONT face="����"><IMG id="ImgAddrow2" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="�ڷḦ �����մϴ�."
															src="../../../images/imgAddRow.gif" width="54" border="0" name="imgConfirm"></FONT>
												</td>
											</tr>
										</table>
									</FONT>
								</TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
