<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORPROGNAME.aspx.vb" Inherits="MD.MDCMOUTDOORPROGNAME" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ܱ��� �������</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/����/�����ڵ� �˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPOP1.aspx
'��      �� : JOBNO ��ȸ�� ���� �˾�
'�Ķ�  ���� : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , ��ȸ�߰��ʵ�, ���� ������� �͸� ��ȸ���� ����,
'			  �ڵ� ������, �ڵ�Like���� ����
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjMDCMGET
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

'-----------------------------
' �̺�Ʈ ���ν��� 
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


Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

Sub imgCalEndarFROM_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTBRDSTDATE,frmThis.imgCalEndarFROM,"txtTBRDSTDATE_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTBRDEDDATE,frmThis.imgCalEndarTO,"txtTBRDEDDATE_onchange()"
		gSetChange
	end with
End Sub

'��۽�����
Sub txtTBRDSTDATE_onchange
	gSetChange
End Sub

'���������
Sub txtTBRDEDDATE_onchange
	gSetChange
End Sub

sub sprSht_DblClick (Col,Row)
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		END IF
	End With
end sub

'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'����������ü ����	
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : .txtTBRDSTDATE.value = vntInParam(i)	
				case 1 : .txtTBRDEDDATE.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 2 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
		'SpreadSheet ������
		gSetSheetDefaultColor()
        With frmThis
			'������ �׸���
            
			'��ü�� �׸���
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 4, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | SEQ | OLDPROGNAME | NEWPROGNAME"
			mobjSCGLSpr.SetHeader .sprSht, "���|����|���������|��������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "8|6|30|30"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "YEARMON | SEQ | OLDPROGNAME | NEWPROGNAME"
        End With
	end with	
	'�ڷ���ȸ	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
'  	Dim vntData
' 	Dim i, strCols
'
'	On error resume next
'	with frmThis
'		'Long Type�� ByRef ������ �ʱ�ȭ
'		mlngRowCnt=clng(0)
'		mlngColCnt=clng(0)
'		
'		vntData = mobjMDCMGET.GetMED_DBLLIST(gstrConfigXml,lngRowCnt,lngColCnt,.txtYEAR.value,.txtMEDNAME.value)		
'  		
' 		if not gDoErrorRtn ("GetMED_DBLLIST") then
'			mobjSCGLSpr.SetClip .sprSht, vntData1, 1, 1, lngColCnt, lngRowCnt, True
'  		end if
' 	end with
end sub

-->
		</script>
	</HEAD>
	<BODY class="base">
		<table cellPadding="0" width="790" border="0" id="tblForm" cellSpacing="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 300px" align="left" width="300" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">
												&nbsp;���� ����� ����
											</td>
										</tr>
									</table>
								</td>
								<TD align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 300px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 168px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="168" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gIF" width="54" align="right" border="0" name="imgSave"></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="DATA" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD class="SEARCHLABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNO,'')">
												������</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtTBRDSTDATE" title="����Ⱓ" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="8" name="txtTBRDSTDATE"><IMG id="imgCalEndarFROM" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarFROM">&nbsp;~ <INPUT class="INPUT" id="txtTBRDEDDATE" title="����Ⱓ" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="1" name="txtTBRDEDDATE"><IMG id="imgCalEndarTO" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarTO"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="DATA" id="tblKey" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD class="LABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNO,'')">
												���� �����</TD>
											<TD class="DATA" width="305"><INPUT class="INPUT_L" id="txtOLDPROGNAME" title="�����" style="WIDTH: 304px; HEIGHT: 22px"
													type="text" maxLength="100000" size="45" name="txtOLDPROGNAME"></TD>
											<TD class="LABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,'')">
												�� �����</TD>
											<TD class="DATA" width="305"><INPUT class="INPUT_L" id="txtNEWPROGNAME" title="�����" style="WIDTH: 304px; HEIGHT: 22px"
													type="text" maxLength="100000" size="45" name="txtNEWPROGNAME"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"></TD>
							</TR>
							<TR>
								<TD align="center">
									<table>
										<tr>
											<td>
												<OBJECT id="sprSht" style="WIDTH: 790px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
													VIEWASTEXT>
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="20902">
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
											</td>
										</tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
				</TD>
				</FORM>
			</TR>
		</table>
	</BODY>
</HTML>
